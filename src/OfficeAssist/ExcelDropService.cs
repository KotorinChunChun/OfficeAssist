namespace OfficeAssist;

internal sealed class ExcelDropService : IDisposable
{
    readonly System.Windows.Forms.Timer timer = new() { Interval = 15 };
    readonly DropOverlay overlay = new();
    readonly DropGesture gesture = new();
    readonly MousePressMonitor? mouse;
    readonly Control dispatcher;
    readonly Action<string, bool> report;
    bool externalSource, suppressed, enabled = true;
    bool allowLinkedImages;
    Point origin;
    nint target;
    internal bool Busy { get; private set; }
    internal bool Enabled { get => enabled; set { enabled = value; if (!value) { suppressed = true; overlay.Reset(); } } }

    internal ExcelDropService(Control dispatcher, Action<string, bool> report, bool startTimer = true)
    {
        this.dispatcher = dispatcher;
        this.report = report;
        _ = overlay.Handle;
        overlay.Rejected += () => { suppressed = true; overlay.Reset(); };
        overlay.FilesDropped += QueueInsert;
        timer.Tick += (_, _) => Tick();
        if (startTimer)
        {
            mouse = new MousePressMonitor(gesture);
            timer.Start();
        }
    }

    void Tick()
    {
        bool down = (Native.GetAsyncKeyState(1) & 0x8000) != 0;
        if (!Native.GetCursorPos(out var point)) { overlay.Reset(); return; }
        if (gesture.TakePress(down, point, Native.WindowFromPoint) is { } press)
        {
            overlay.Reset();
            target = 0;
            origin = press.Position;
            externalSource = down && Native.IsExternalImageDragSource(press.Source);
            allowLinkedImages = externalSource && Native.IsExplorer(press.Source);
            suppressed = false;
        }
        if ((Native.GetAsyncKeyState(27) & 0x8000) != 0) suppressed = true;
        if (!enabled || Busy || !externalSource || suppressed)
        { overlay.Reset(); return; }
        if (!down)
        {
            if (!gesture.WaitForDrop(Environment.TickCount64, overlay.Receiving))
            { overlay.Reset(); externalSource = false; }
            return;
        }
        if (Math.Abs(point.X - origin.X) < SystemInformation.DragSize.Width
            && Math.Abs(point.Y - origin.Y) < SystemInformation.DragSize.Height) return;
        if (overlay.Visible)
        {
            if (!Native.IsWindow(target) || !Native.IsWindowVisible(target) || !overlay.Bounds.Contains(point)) overlay.Reset();
            return;
        }
        target = Native.ExcelRootAt(point);
        if (target == 0 || !Native.GetWindowRect(target, out var rect)) return;
        overlay.ImageBounds = Native.VisibleExcelSheet(target, point).Bounds;
        overlay.Bounds = rect.Bounds;
        overlay.Show();
    }

    bool QueueInsert(string[] files, Point point, int keyState)
    {
        if (Busy || !enabled || gesture.HasPendingPress) return false;
        suppressed = true;
        var handle = target;
        bool allowLinks = allowLinkedImages;
        ReceivedImageBatch? batch = null;
        ReceivedTabularBatch? tables = null;
        bool showOptions = false;
        try
        {
            bool tabular = TabularInputDiscovery.Accepts(files);
            showOptions = InsertSettings.LoadBehavior().Show(tabular, keyState);
            if (tabular) tables = ReceivedTabularBatch.Capture(files);
            else batch = ReceivedImageBatch.Capture(files);
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or ArgumentException
            or InvalidOperationException or System.Runtime.InteropServices.ExternalException)
        { report("ファイルを受け取れませんでした: " + ex.Message, true); return false; }
        Busy = true;
        // OLEのDrop呼出中にExcelへCOMを送るとビジーになるため、応答後にUIキューから実行する。
        try { dispatcher.BeginInvoke(new Action(() =>
        {
            using var received = batch;
            using var receivedTables = tables;
            try
            {
                if (Native.ExcelRootAt(point) != handle)
                {
                    report("ドロップ先のExcelを確認できません。対象のExcelを表示して、再度ドロップしてください。", true);
                    return;
                }
                var sheet = Native.VisibleExcelSheet(handle, point);
                if (tables is null && !sheet.Bounds.Contains(point)) return;
                var result = tables is not null ? ExcelTabularInserter.Insert(handle, sheet, point, tables, showOptions)
                    : ExcelPictureInserter.InsertInteractive(sheet.Handle, point, received!.Images, received.Originals, allowLinks, showOptions);
                if (result is not null)
                {
                    report(result.Warning is null ? result.Message : result.Message + Environment.NewLine + result.Warning, false);
                    if (result.Warning is not null && tables is null) ExcelModalLock.ShowMessage(handle, result.Warning, WindowTitle.Create("ファイルの展開"), MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                var message = ex is System.Runtime.InteropServices.COMException
                    ? "Excelへ接続できません。セルの編集やダイアログを閉じ、通常権限のExcelへ再度ドロップしてください。"
                    : ex.Message;
                report(message, true);
                ExcelModalLock.ShowMessage(handle, message, WindowTitle.Create("ファイルの展開"), MessageBoxIcon.Warning);
            }
            finally { Busy = false; }
        })); }
        catch (InvalidOperationException) { batch?.Dispose(); tables?.Dispose(); Busy = false; return false; }
        return true;
    }

    public void Dispose() { timer.Stop(); timer.Dispose(); mouse?.Dispose(); overlay.Dispose(); }
}
