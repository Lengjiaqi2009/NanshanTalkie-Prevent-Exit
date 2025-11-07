#include <windows.h>
#include <tlhelp32.h>
#include <tchar.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <combaseapi.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "mmdevapi.lib")

// 目标应用程序信息（学校教学用对讲软件）
#define APP_EXE_NAME L"nsptt_5.2.1.exe"        // 应用程序进程名
#define APP_FULL_PATH L"D:/nsptt_5.2.1.exe"   // 应用程序完整路径
#define WINDOW_CAPTION L"南山对讲"           // 应用程序窗口标题
#define VOLUME_CHECK_INTERVAL 1000  // 音量检查间隔(毫秒)
#define WINDOW_WAIT_DELAY 500       // 等待窗口创建的延迟(毫秒)

/**
 * 通过进程名获取进程ID
 * @param exeName 进程名（带.exe）
 * @return 进程ID，0表示未找到
 */
DWORD GetProcessIdByName(const wchar_t* exeName) {
    PROCESSENTRY32W pe32 = { sizeof(PROCESSENTRY32W) };
    HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnapshot == INVALID_HANDLE_VALUE) return 0;

    if (Process32FirstW(hSnapshot, &pe32)) {
        do {
            if (_wcsicmp(pe32.szExeFile, exeName) == 0) {
                CloseHandle(hSnapshot);
                return pe32.th32ProcessID;
            }
        } while (Process32NextW(hSnapshot, &pe32));
    }

    CloseHandle(hSnapshot);
    return 0;
}

/**
 * 通过窗口标题查找窗口并最小化（保留任务栏图标）
 * @param caption 窗口标题
 * @return 成功返回TRUE，失败返回FALSE
 */
BOOL MinimizeWindowByCaption(const wchar_t* caption) {
    // 查找目标窗口句柄
    HWND hWnd = FindWindowW(NULL, caption);
    if (hWnd == NULL) {
        return FALSE;
    }

    // 最小化窗口（SW_MINIMIZE：仅最小化窗口，保留任务栏图标）
    ShowWindow(hWnd, SW_MINIMIZE);
    return TRUE;
}

/**
 * 启动目标程序并最小化窗口（保留任务栏图标，不干扰课堂）
 */
void LaunchAndMinimize() {
    STARTUPINFOW si = { sizeof(STARTUPINFOW) };
    PROCESS_INFORMATION pi = { 0 };

    // 启动时设置窗口显示方式为最小化（保留任务栏图标）
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_MINIMIZE;  // 最小化启动，任务栏可见

    if (CreateProcessW(APP_FULL_PATH, NULL, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
        // 启动后等待5秒再最小化，确保程序完成初始化
        Sleep(5000);

        // 确保窗口处于最小化状态
        MinimizeWindowByCaption(WINDOW_CAPTION);

        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
    }
}

/**
 * 后台设置系统音量为100%（无窗口交互，不干扰课堂）
 */
HRESULT SetVolumeToMax() {
    HRESULT hr;
    IMMDeviceEnumerator* pEnumerator = NULL;
    IMMDevice* pDevice = NULL;
    IAudioEndpointVolume* pEndpointVolume = NULL;

    // 初始化COM（后台操作，无窗口界面）
    hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(hr)) goto cleanup;

    // 获取音频设备枚举器
    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, 
                        CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), 
                        (void**)&pEnumerator);
    if (FAILED(hr)) goto cleanup;

    // 获取默认音频输出设备
    hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
    if (FAILED(hr)) goto cleanup;

    // 激活音量控制接口
    hr = pDevice->Activate(__uuidof(IAudioEndpointVolume), 
                         CLSCTX_ALL, NULL, (void**)&pEndpointVolume);
    if (FAILED(hr)) goto cleanup;

    // 设置主音量为100%（确保对讲声音清晰）
    hr = pEndpointVolume->SetMasterVolumeLevelScalar(1.0f, NULL);

    // 如果系统处于静音状态（例如音量为0时自动静音），解除静音
    // 保持原有功能：仅在提升音量时同时取消静音，不改变其它行为
    if (SUCCEEDED(hr)) {
        // 忽略 SetMute 的返回值（非关键），但调用以确保取消静音
        pEndpointVolume->SetMute(FALSE, NULL);
    }

cleanup:
    // 释放资源，避免内存泄漏
    if (pEndpointVolume) pEndpointVolume->Release();
    if (pDevice) pDevice->Release();
    if (pEnumerator) pEnumerator->Release();
    CoUninitialize();
    return hr;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // 隐藏当前程序的控制台窗口（自身后台运行）
    HWND hWnd = GetConsoleWindow();
    if (hWnd != NULL) {
        ShowWindow(hWnd, SW_HIDE);
    }

    // 提高进程优先级，确保稳定运行
    HANDLE hSelf = GetCurrentProcess();
    if (hSelf) {
        SetPriorityClass(hSelf, HIGH_PRIORITY_CLASS);
        CloseHandle(hSelf);
    }

    while (true) {
        // 定期检查并维持最大音量
        SetVolumeToMax();

        // 检查目标程序是否运行
        DWORD pid = GetProcessIdByName(APP_EXE_NAME);
        
        if (pid == 0) {
            // 程序未运行，启动并最小化（保留任务栏图标）
            LaunchAndMinimize();
        } else {
            // 程序已运行，确保窗口处于最小化状态
            MinimizeWindowByCaption(WINDOW_CAPTION);
        }

        // 等待下次检查
        Sleep(VOLUME_CHECK_INTERVAL);
    }

    return 0;
}