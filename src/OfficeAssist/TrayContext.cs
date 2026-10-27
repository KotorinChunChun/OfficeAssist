using KotorinToys.Integration;

namespace OfficeAssist;

internal sealed class TrayContext : ApplicationContext
{
    readonly MainForm window = new();
    readonly Icon icon = AppIcon.Create();
    readonly NotifyIcon tray;
    readonly ContextMenuStrip menu = new();
    readonly ManagementEndpoint endpoint;
    readonly ExcelDropService dropService;
    public TrayContext(bool silent)
    {
        _ = window.Handle;
        var toggle = new ToolStripMenuItem("Excelへの画像ドロップ") { Checked = true, CheckOnClick = true };
        menu.Items.Add("開く", null, (_, _) => ShowWindow());
        menu.Items.Add(toggle);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("終了", null, (_, _) => Stop());
        tray = new NotifyIcon { Icon = icon, Text = "OfficeAssist", Visible = true, ContextMenuStrip = menu };
        tray.DoubleClick += (_, _) => ShowWindow();
        toggle.CheckedChanged += (_, _) => window.FeatureEnabled = toggle.Checked;
        dropService = new ExcelDropService(window, (message, failed) =>
        {
            window.SetResult(message);
        });
        window.FeatureChanged += enabled =>
        {
            toggle.Checked = enabled;
            dropService.Enabled = enabled;
            window.SetResult(enabled ? "待機中" : "一時停止中");
        };
        window.ExitRequested += Stop;
        endpoint = new ManagementEndpoint("office-assist", window, ShowWindow, Stop,
            () => dropService.Busy ? "画像を挿入中です。" : null);
        if (!silent) ShowWindow();
    }

    void ShowWindow()
    {
        window.Show();
        window.WindowState = FormWindowState.Normal;
        window.Activate();
    }

    void Stop() { if (dropService.Busy) return; window.AllowExit = true; window.Close(); ExitThread(); }

    protected override void Dispose(bool disposing)
    {
        if (disposing) { endpoint.Dispose(); dropService.Dispose(); tray.Visible = false; tray.Dispose(); menu.Dispose(); window.Dispose(); icon.Dispose(); }
        base.Dispose(disposing);
    }
}
