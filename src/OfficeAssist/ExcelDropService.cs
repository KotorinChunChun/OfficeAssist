namespace OfficeAssist;

internal sealed class ExcelDropService : IDisposable
{
    readonly System.Windows.Forms.Timer timer = new() { Interval = 15 };
    readonly DropOverlay overlay = new();
    readonly Control dispatcher;
    readonly Action<string, bool> report;
    bool wasDown, fromExplorer, suppressed, enabled = true;
    Point origin;
    nint target;
    int releasedTicks;
    internal bool Busy { get; private set; }
    internal bool Enabled { get => enabled; set { enabled = value; if (!value) { suppressed = true; overlay.Hide(); } } }

    internal ExcelDropService(Control dispatcher, Action<string, bool> report)
    {
        this.dispatcher = dispatcher;
        this.report = report;
        _ = overlay.Handle;
        overlay.Rejected += () => { suppressed = true; overlay.Hide(); };
        overlay.FilesDropped += QueueInsert;
        timer.Tick += (_, _) => Tick();
        timer.Start();
    }

    void Tick()
    {
        bool down = (Native.GetAsyncKeyState(1) & 0x8000) != 0;
        Native.GetCursorPos(out var point);
        if (down && !wasDown)
        {
            origin = point;
            fromExplorer = Native.IsExplorer(Native.WindowFromPoint(point));
            suppressed = false;
        }
        wasDown = down;
        if (!down)
        {
            // OLE側のDropが先に処理される猶予を持たせる。
            if (++releasedTicks >= 3) { overlay.Hide(); fromExplorer = false; }
            return;
        }
        releasedTicks = 0;
        if (!enabled || Busy || !fromExplorer || suppressed || (Native.GetAsyncKeyState(27) & 0x8000) != 0)
        { overlay.Hide(); return; }
        if (Math.Abs(point.X - origin.X) < SystemInformation.DragSize.Width
            && Math.Abs(point.Y - origin.Y) < SystemInformation.DragSize.Height) return;
        if (overlay.Visible)
        {
            if (!Native.IsWindow(target) || !Native.IsWindowVisible(target) || !overlay.Bounds.Contains(point)) overlay.Hide();
            return;
        }
        target = Native.ExcelSheetAt(point);
        if (target == 0 || !Native.GetWindowRect(target, out var rect)) return;
        overlay.Bounds = rect.Bounds;
        overlay.Show();
    }

    void QueueInsert(string[] files, Point point)
    {
        var handle = target;
        Busy = true;
        suppressed = true;
        overlay.Hide();
        // OLEのDrop呼出中にExcelへCOMを送るとビジーになるため、応答後にUIキューから実行する。
        dispatcher.BeginInvoke(new Action(() =>
        {
            try { report(ExcelPictureInserter.Insert(handle, point, files), false); }
            catch (Exception ex)
            {
                var message = ex is System.Runtime.InteropServices.COMException
                    ? "Excelへ接続できません。セルの編集やダイアログを閉じ、通常権限のExcelへ再度ドロップしてください。"
                    : ex.Message;
                report(message, true);
            }
            finally { Busy = false; }
        }));
    }

    public void Dispose() { timer.Stop(); timer.Dispose(); overlay.Dispose(); }
}
