using System.Collections.Concurrent;
using System.Diagnostics;

namespace OfficeAssist;

// COMは専用STAだけで使用する。取り込み先のCOMオブジェクトを渡さない。
internal sealed class ExcelFormatPreview : IExcelFormatPreview
{
    sealed record Request(IReadOnlyList<ExcelPreviewInput> Cells, ExcelPreviewEnvironment Environment, CancellationToken Token, TaskCompletionSource<IReadOnlyList<ExcelPreviewCell>> Completion);
    readonly BlockingCollection<Request> queue = new(4);
    readonly Thread worker;
    readonly bool cacheResults, batchStyles, literalText;
    bool disposed;
    internal int OwnedPid { get; private set; }
    internal ExcelPreviewMetrics LastMetrics { get; private set; } = new(0, 0, 0, 0, 0, 0);
    internal ExcelFormatPreview(bool cacheResults = true, bool batchStyles = true, bool literalText = true)
    {
        this.cacheResults = cacheResults; this.batchStyles = batchStyles; this.literalText = literalText;
        worker = new Thread(Run) { IsBackground = true, Name = "OfficeAssist.ExcelFormatPreview" };
        worker.SetApartmentState(ApartmentState.STA); worker.Start();
    }
    public Task<IReadOnlyList<ExcelPreviewCell>> RenderAsync(IReadOnlyList<ExcelPreviewInput> cells, ExcelPreviewEnvironment environment, CancellationToken token)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        var completion = new TaskCompletionSource<IReadOnlyList<ExcelPreviewCell>>(TaskCreationOptions.RunContinuationsAsynchronously);
        var request = new Request(cells.ToArray(), environment, token, completion);
        if (!queue.TryAdd(request)) completion.SetException(new InvalidOperationException("プレビューの更新待ちです。もう一度操作してください。"));
        return completion.Task;
    }
    void Run()
    {
        object? application = null, workbook = null;
        bool owned = false; DateTime started = default;
        bool oldSystemSeparators = true; string oldDecimal = ".", oldThousands = ",";
        string defaultFont = ""; double defaultFontSize = 11;
        ExcelPreviewSession? session = null;
        using var lifetime = new ComScope();
        try
        {
            foreach (var request in queue.GetConsumingEnumerable())
            {
                if (request.Token.IsCancellationRequested) { request.Completion.TrySetCanceled(request.Token); continue; }
                try
                {
                    if (literalText)
                    {
                        var literals = new List<ExcelPreviewCell>();
                        foreach (var input in request.Cells) { if (!ExcelPreviewSession.TryLiteral(input, out var cell)) break; literals.Add(cell); }
                        if (literals.Count == request.Cells.Count) { LastMetrics = new(literals.Count, 0, literals.Count, 0, 0, 0); request.Completion.TrySetResult(literals); continue; }
                    }
                    if (application is null)
                    {
                        var existing = Process.GetProcessesByName("EXCEL").Select(p => { using (p) return p.Id; }).ToHashSet();
                        dynamic app = lifetime.Own(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!);
                        Native.GetWindowThreadProcessId((nint)(int)app.Hwnd, out uint pid);
                        if (existing.Contains((int)pid)) throw new InvalidOperationException("専用Excelを作成できません。既存ブックへは接続しません。");
                        application = app; OwnedPid = (int)pid; owned = true;
                        using (var process = Process.GetProcessById(OwnedPid)) started = process.StartTime;
                        ownedStart = started; oldSystemSeparators = (bool)app.UseSystemSeparators; oldDecimal = (string)app.DecimalSeparator; oldThousands = (string)app.ThousandsSeparator;
                        app.Visible = false; app.DisplayAlerts = false; app.EnableEvents = false; app.AutomationSecurity = 3;
                        dynamic books = lifetime.Own(app.Workbooks); workbook = lifetime.Own(books.Add(-4167));
                        dynamic styles = lifetime.Own(((dynamic)workbook).Styles), normal = lifetime.Own(styles.Item("Normal")), font = lifetime.Own(normal.Font);
                        defaultFont = (string)font.Name; defaultFontSize = (double)font.Size;
                        app.ScreenUpdating = false; app.Calculation = -4135;
                        session = new(application, workbook, defaultFont, defaultFontSize, cacheResults, batchStyles, literalText);
                    }
                    var rendered = session!.Render(request.Cells, request.Environment, request.Token); LastMetrics = session.Metrics;
                    request.Completion.TrySetResult(rendered);
                }
                catch (OperationCanceledException) { request.Completion.TrySetCanceled(request.Token); }
                catch (Exception ex) { request.Completion.TrySetException(ex); }
            }
        }
        finally
        {
            if (owned && application is not null)
            {
                try { dynamic app = application; app.DecimalSeparator = oldDecimal; app.ThousandsSeparator = oldThousands; app.UseSystemSeparators = oldSystemSeparators; if (workbook is not null) { dynamic book = workbook; book.Close(false); } app.Quit(); } catch { }
            }
            foreach (var remaining in queue) remaining.Completion.TrySetCanceled();
            if (owned) ownedStart = started;
        }
    }
    DateTime ownedStart;
    public void Dispose()
    {
        if (disposed) return; disposed = true; queue.CompleteAdding();
        worker.Join(5000);
        // Excelが後からモーダル状態になりQuitを拒否した場合も、所有した専用PIDだけ回収する。
        if (OwnedPid > 0)
        {
            try { using var process = Process.GetProcessById(OwnedPid); if (ownedStart != default && process.StartTime == ownedStart && process.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) && !process.WaitForExit(1000)) process.Kill(); } catch { }
        }
    }
}
