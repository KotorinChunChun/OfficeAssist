namespace OfficeAssist;

internal sealed record ExcelWindowOwner(nint Handle) : IWin32Window;

// WinFormsの子ダイアログが開閉しても、外側のExcelロックを解放しない。
internal sealed class ExcelModalLock : IDisposable
{
    readonly nint window;
    readonly bool restoreEnabled;
    readonly uint threadId, processId;
    bool disposed;
    internal ExcelModalLock(nint window)
    {
        this.window = window;
        threadId = Native.GetWindowThreadProcessId(window, out processId);
        if (threadId == 0) return;
        restoreEnabled = Native.IsWindowEnabled(window);
        if (restoreEnabled) Native.EnableWindow(window, false);
    }
    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        if (restoreEnabled && IsSameWindow())
            Native.EnableWindow(window, true);
    }

    bool IsSameWindow() => threadId != 0 && Native.GetWindowThreadProcessId(window, out uint pid) == threadId && pid == processId;
    internal static void ShowMessage(nint sheetHandle, string message, string title, MessageBoxIcon icon)
    {
        nint owner = Native.GetAncestor(sheetHandle, 2);
        using var modalLock = new ExcelModalLock(owner);
        if (Native.IsWindow(owner)) MessageBox.Show(new ExcelWindowOwner(owner), message, title, MessageBoxButtons.OK, icon);
        else MessageBox.Show(message, title, MessageBoxButtons.OK, icon);
    }
}

internal sealed class ExcelDropLocks : IDisposable
{
    readonly Dictionary<nint, ExcelModalLock> locks = new();

    internal void LockExisting(nint window)
    {
        if (!locks.ContainsKey(window)) locks.Add(window, new(window));
    }

    internal void ReleaseAll()
    {
        foreach (var modalLock in locks.Values.Reverse()) modalLock.Dispose();
        locks.Clear();
    }

    public void Dispose() => ReleaseAll();
}
