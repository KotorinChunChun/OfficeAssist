using System.Collections.Concurrent;
using System.Diagnostics;

namespace OfficeAssist;

// COMは専用STAだけで使用する。取り込み先のCOMオブジェクトを渡さない。
internal sealed class ExcelFormatPreview : IExcelFormatPreview
{
    sealed record Request(IReadOnlyList<ExcelPreviewInput> Cells, ExcelPreviewEnvironment Environment, CancellationToken Token, TaskCompletionSource<IReadOnlyList<ExcelPreviewCell>> Completion);
    readonly BlockingCollection<Request> queue = new(4);
    readonly Thread worker;
    bool disposed;
    internal int OwnedPid { get; private set; }
    internal ExcelFormatPreview()
    {
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
        using var lifetime = new ComScope();
        try
        {
            foreach (var request in queue.GetConsumingEnumerable())
            {
                if (request.Token.IsCancellationRequested) { request.Completion.TrySetCanceled(request.Token); continue; }
                try
                {
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
                    }
                    request.Completion.TrySetResult(Render(application, workbook!, request, defaultFont, defaultFontSize));
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
    static IReadOnlyList<ExcelPreviewCell> Render(dynamic app, dynamic book, Request request, string defaultFont, double defaultFontSize)
    {
        using var com = new ComScope();
        var environment = request.Environment;
        app.UseSystemSeparators = environment.UseSystemSeparators;
        if (!environment.UseSystemSeparators) { app.DecimalSeparator = environment.DecimalSeparator; app.ThousandsSeparator = environment.ThousandsSeparator; }
        book.Date1904 = environment.Date1904;
        dynamic styles = com.Own(book.Styles), normal = com.Own(styles.Item("Normal")), normalFont = com.Own(normal.Font);
        normalFont.Name = environment.NormalFont ?? defaultFont;
        normalFont.Size = environment.NormalFontSize ?? defaultFontSize;
        dynamic sheets = com.Own(book.Worksheets), sheet = com.Own(sheets.Item(1)), cell = com.Own(sheet.Range["A1"]), column = com.Own(cell.EntireColumn);
        var widths = new Dictionary<double, double>(); var results = new List<ExcelPreviewCell>();
        double standardWidth = (double)sheet.StandardWidth;
        foreach (var input in request.Cells)
        {
            request.Token.ThrowIfCancellationRequested();
            string code = ExcelValueFormat.Code(input.Type, input.Decimals);
            try
            {
                object? value = ExcelValueFormat.Value(input.Raw, input.Type, false, environment.Date1904);
                cell.Clear(); var style = input.Style ?? new();
                dynamic font = com.Own(cell.Font);
                font.Name = style.FontName ?? (string)normalFont.Name; font.Size = style.FontSize ?? (double)normalFont.Size;
                font.Bold = style.Bold; font.Italic = style.Italic; cell.WrapText = style.Wrap; cell.ShrinkToFit = style.Shrink;
                cell.RowHeight = style.RowHeight ?? (double)sheet.StandardHeight;
                if (style.WidthPoints.HasValue)
                {
                    double points = style.WidthPoints.Value;
                    if (!widths.TryGetValue(points, out double characters))
                    {
                        double low = 0, high = 255; characters = 0;
                        for (int i = 0; i < 16; i++) { characters = (low + high) / 2; column.ColumnWidth = characters; double actual = (double)column.Width; if (Math.Abs(actual - points) < 0.2) break; if (actual < points) low = characters; else high = characters; }
                        widths[points] = (double)column.ColumnWidth;
                    }
                    column.ColumnWidth = widths[points];
                }
                else column.ColumnWidth = style.WidthCharacters ?? standardWidth;
                ExcelValueFormat.Apply(app, cell, ExcelValueFormat.CellCode(input.Type, input.Decimals, value));
                cell.Value2 = value;
                if (cell.HasFormula is not bool formula || formula) throw new InvalidOperationException("プレビューに意図しない数式があります。");
                string display = (string)cell.Text;
                string? warning = display.Length > 0 && display.All(ch => ch == '#') ? "Excelではこの列幅・値・書式の組合せが#表示になります。列幅、負の時刻などを確認してください。" : null;
                results.Add(new(input.Row, input.Column, input.Raw, display, (object?)cell.Value2, code, (string)cell.NumberFormat, (string)cell.PrefixCharacter, Warning: warning));
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            { results.Add(new(input.Row, input.Column, input.Raw, "変換エラー", null, code, code, "", ex.Message)); }
        }
        return results;
    }
    public void Dispose()
    {
        if (disposed) return; disposed = true; queue.CompleteAdding();
        if (!worker.Join(5000) && OwnedPid > 0)
        {
            try { using var process = Process.GetProcessById(OwnedPid); if (ownedStart != default && process.StartTime == ownedStart && process.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase)) process.Kill(); } catch { }
        }
    }
}
