using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace Auto_SLD.Service;

// Auto Drawing 중 사용자의 물리 마우스 입력만 차단하고 SendInput으로 생성한 자동 입력은 통과시킨다.
internal sealed class MouseInputGuard : IDisposable
{
    private const int WhMouseLowLevel = 14;
    private const uint MouseFlagInjected = 0x00000001;
    private const uint WmQuit = 0x0012;

    private readonly ManualResetEventSlim _ready = new ManualResetEventSlim(false);
    private readonly Thread _hookThread;
    private MouseHookCallback? _callback;
    private Exception? _startupException;
    private IntPtr _hook;
    private uint _hookThreadId;
    private bool _disposed;

    public MouseInputGuard()
    {
        _hookThread = new Thread(RunHook) { IsBackground = true, Name = "Auto SLD Mouse Input Guard" };
        _hookThread.Start();
        _ready.Wait();
        if (_startupException != null)
        {
            _hookThread.Join();
            _ready.Dispose();
            throw new InvalidOperationException($"Auto Drawing 중 마우스 입력을 차단하지 못했습니다: {_startupException.Message}", _startupException);
        }
    }

    private void RunHook()
    {
        _hookThreadId = GetCurrentThreadId();
        PeekMessage(out NativeMessage _, IntPtr.Zero, 0, 0, 0); // Dispose가 즉시 호출돼도 WM_QUIT을 받을 수 있게 큐를 먼저 만든다.
        _callback = OnMouseInput;
        _hook = SetWindowsHookEx(WhMouseLowLevel, _callback, GetModuleHandle(null), 0);
        if (_hook == IntPtr.Zero) _startupException = new Win32Exception(Marshal.GetLastWin32Error());
        _ready.Set();
        if (_hook == IntPtr.Zero) return;

        try
        {
            while (GetMessage(out NativeMessage message, IntPtr.Zero, 0, 0) > 0) { }
        }
        finally
        {
            UnhookWindowsHookEx(_hook);
            _hook = IntPtr.Zero;
        }
    }

    private IntPtr OnMouseInput(int code, IntPtr message, IntPtr data)
    {
        if (code >= 0)
        {
            MouseHookData input = Marshal.PtrToStructure<MouseHookData>(data);
            if ((input.Flags & MouseFlagInjected) == 0) return new IntPtr(1);
        }
        return CallNextHookEx(_hook, code, message, data);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        if (_hook != IntPtr.Zero)
        {
            if (!PostThreadMessage(_hookThreadId, WmQuit, UIntPtr.Zero, IntPtr.Zero)) Debug.WriteLine($"[마우스 입력 해제] 훅 스레드 종료 메시지 전달 실패: {new Win32Exception(Marshal.GetLastWin32Error()).Message}");
            if (!_hookThread.Join(5000))
            {
                Debug.WriteLine("[마우스 입력 해제] 훅 스레드가 5초 안에 종료되지 않아 훅을 직접 해제합니다.");
                UnhookWindowsHookEx(_hook);
                _hook = IntPtr.Zero;
            }
        }
        _ready.Dispose();
    }

    private delegate IntPtr MouseHookCallback(int code, IntPtr message, IntPtr data);

    [StructLayout(LayoutKind.Sequential)]
    private struct Point { public int X, Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MouseHookData { public Point Point; public uint MouseData, Flags, Time; public UIntPtr ExtraInfo; }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMessage { public IntPtr Window; public uint Message; public UIntPtr WParam; public IntPtr LParam; public uint Time; public Point Point; public uint Private; }

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)] private static extern IntPtr SetWindowsHookEx(int hookId, MouseHookCallback callback, IntPtr module, uint threadId);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool UnhookWindowsHookEx(IntPtr hook);
    [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr message, IntPtr data);
    [DllImport("user32.dll")] private static extern int GetMessage(out NativeMessage message, IntPtr window, uint minimum, uint maximum);
    [DllImport("user32.dll")] private static extern bool PeekMessage(out NativeMessage message, IntPtr window, uint minimum, uint maximum, uint removeMessage);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool PostThreadMessage(uint threadId, uint message, UIntPtr wParam, IntPtr lParam);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr GetModuleHandle(string? moduleName);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
}
