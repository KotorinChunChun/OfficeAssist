namespace OfficeAssist;

internal sealed record ExcelWindowOwner(nint Handle) : IWin32Window;

// WinFormsの子ダイアログが開閉しても、外側のExcelロックを解放しない。
internal sealed class ExcelModalLock : IDisposable
{
    readonly nint window;
    readonly bool restoreEnabled;
    bool disposed;
    internal ExcelModalLock(nint window)
    {
        this.window = window;
        restoreEnabled = Native.IsWindow(window) && Native.IsWindowEnabled(window);
        if (restoreEnabled) Native.EnableWindow(window, false);
    }
    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        if (restoreEnabled && Native.IsWindow(window)) Native.EnableWindow(window, true);
    }
    internal static void ShowMessage(nint sheetHandle, string message, string title, MessageBoxIcon icon)
    {
        nint owner = Native.GetAncestor(sheetHandle, 2);
        using var modalLock = new ExcelModalLock(owner);
        if (Native.IsWindow(owner)) MessageBox.Show(new ExcelWindowOwner(owner), message, title, MessageBoxButtons.OK, icon);
        else MessageBox.Show(message, title, MessageBoxButtons.OK, icon);
    }
}
