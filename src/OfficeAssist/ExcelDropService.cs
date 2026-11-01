namespace OfficeAssist;

internal sealed class ExcelDropService : IDisposable
{
    readonly System.Windows.Forms.Timer timer = new() { Interval = 15 };
    readonly DropOverlay overlay = new();
    readonly Control dispatcher;
    readonly Action<string, bool> report;
    bool wasDown, externalSource, suppressed, enabled = true;
    bool allowLinkedImages;
    Point origin;
    nint target;
    int releasedTicks;
    internal bool Busy { get; private set; }
    internal bool Enabled { get => enabled; set { enabled = value; if (!value) { suppressed = true; overlay.Hide(); } } }

    internal ExcelDropService(Control dispatcher, Action<string, bool> report, bool startTimer = true)
    {
        this.dispatcher = dispatcher;
        this.report = report;
        _ = overlay.Handle;
        overlay.Rejected += () => { suppressed = true; overlay.Hide(); };
        overlay.FilesDropped += QueueInsert;
        timer.Tick += (_, _) => Tick();
        if (startTimer) timer.Start();
    }

    void Tick()
    {
        bool down = (Native.GetAsyncKeyState(1) & 0x8000) != 0;
        Native.GetCursorPos(out var point);
        if (down && !wasDown)
        {
            origin = point;
            nint source = Native.WindowFromPoint(point);
            externalSource = Native.IsExternalImageDragSource(source);
            allowLinkedImages = externalSource && Native.IsExplorer(source);
            suppressed = false;
        }
        wasDown = down;
        if (!down)
        {
            // OLE側のDropが先に処理される猶予を持たせる。
            if (++releasedTicks >= 3) { overlay.Hide(); externalSource = false; }
            return;
        }
        releasedTicks = 0;
        if (!enabled || Busy || !externalSource || suppressed || (Native.GetAsyncKeyState(27) & 0x8000) != 0)
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

    bool QueueInsert(string[] files, Point point)
    {
        if (Busy || !enabled) return false;
        suppressed = true;
        var handle = target;
        bool allowLinks = allowLinkedImages;
        ReceivedImageBatch batch;
        try { batch = ReceivedImageBatch.Capture(files); }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or ArgumentException
            or InvalidOperationException or System.Runtime.InteropServices.ExternalException)
        { report("画像を受け取れませんでした: " + ex.Message, true); return false; }
        Busy = true;
        overlay.Hide();
        // OLEのDrop呼出中にExcelへCOMを送るとビジーになるため、応答後にUIキューから実行する。
        try { dispatcher.BeginInvoke(new Action(() =>
        {
            using var received = batch;
            try { var result = ExcelPictureInserter.InsertInteractive(handle, point, received.Images, received.Originals, allowLinks); if (result is not null) report(result, false); }
            catch (Exception ex)
            {
                var message = ex is System.Runtime.InteropServices.COMException
                    ? "Excelへ接続できません。セルの編集やダイアログを閉じ、通常権限のExcelへ再度ドロップしてください。"
                    : ex.Message;
                report(message, true);
                MessageBox.Show(message, WindowTitle.Create("画像の挿入"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { Busy = false; }
        })); }
        catch (InvalidOperationException) { batch.Dispose(); Busy = false; return false; }
        return true;
    }

    public void Dispose() { timer.Stop(); timer.Dispose(); overlay.Dispose(); }
}
