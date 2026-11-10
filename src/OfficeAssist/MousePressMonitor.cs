using System.Runtime.InteropServices;

namespace OfficeAssist;

// COM処理でUIが待っている間も、短時間でマウスフックから戻れるよう専用スレッドを使う。
internal sealed class MousePressMonitor : IDisposable
{
    const int LowLevelMouse = 14;
    const int LeftButtonDown = 0x0201;
    const uint Quit = 0x0012;
    readonly Thread worker;
    readonly DropGesture gesture;
    uint threadId;
    internal bool Installed => threadId != 0 && worker.IsAlive;

    internal MousePressMonitor(DropGesture gesture)
    {
        this.gesture = gesture;
        var ready = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        worker = new Thread(() => Run(ready)) { IsBackground = true, Name = "OfficeAssist.MousePress" };
        worker.Start();
        ready.Task.GetAwaiter().GetResult();
    }

    void Run(TaskCompletionSource ready)
    {
        HookProcedure callback = OnMouse;
        nint hook = 0;
        try
        {
            // PostThreadMessageを受け取れるよう、先にこのスレッドのキューを作る。
            PeekMessage(out _, 0, 0, 0, 0);
            hook = SetWindowsHookEx(LowLevelMouse, callback, GetModuleHandle(null), 0);
            if (hook != 0) threadId = GetCurrentThreadId();
        }
        finally { ready.SetResult(); }

        // フックが利用できない環境では、サービスが押下を従来のタイマーで補う。
        if (hook == 0) return;
        try { while (GetMessage(out _, 0, 0, 0) > 0) { } }
        finally
        {
            UnhookWindowsHookEx(hook);
            GC.KeepAlive(callback);
        }
    }

    nint OnMouse(int code, nint message, nint data)
    {
        if (code == 0 && message == LeftButtonDown)
        {
            // MSLLHOOKSTRUCTの先頭はスクリーン座標のPOINT。
            var position = new Point(Marshal.ReadInt32(data), Marshal.ReadInt32(data, sizeof(int)));
            gesture.RecordPress(position, Native.WindowFromPoint(position));
        }
        return CallNextHookEx(0, code, message, data);
    }

    public void Dispose()
    {
        if (Installed) PostThreadMessage(threadId, Quit, 0, 0);
        worker.Join(1000);
    }

    [StructLayout(LayoutKind.Sequential)]
    struct NativeMessage
    {
        internal nint Window;
        internal uint Message;
        internal nuint WParam;
        internal nint LParam;
        internal uint Time;
        internal Point Position;
        internal uint Private;
    }

    delegate nint HookProcedure(int code, nint message, nint data);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("kernel32.dll", EntryPoint = "GetModuleHandleW", CharSet = CharSet.Unicode)] static extern nint GetModuleHandle(string? name);
    [DllImport("user32.dll", EntryPoint = "SetWindowsHookExW")] static extern nint SetWindowsHookEx(int id, HookProcedure callback, nint module, uint thread);
    [DllImport("user32.dll")] static extern bool UnhookWindowsHookEx(nint hook);
    [DllImport("user32.dll")] static extern nint CallNextHookEx(nint hook, int code, nint message, nint data);
    [DllImport("user32.dll", EntryPoint = "PeekMessageW")] static extern bool PeekMessage(out NativeMessage message, nint window, uint min, uint max, uint remove);
    [DllImport("user32.dll", EntryPoint = "GetMessageW")] static extern int GetMessage(out NativeMessage message, nint window, uint min, uint max);
    [DllImport("user32.dll", EntryPoint = "PostThreadMessageW")] static extern bool PostThreadMessage(uint thread, uint message, nint wParam, nint lParam);
}
