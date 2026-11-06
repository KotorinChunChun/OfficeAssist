using OfficeAssist;
using System.Diagnostics;

internal static partial class Program
{
    static void FormattedPreviewFormTests()
    {
        var rows = new[] { new[] { "日付" } }.Concat(Enumerable.Range(0, 12).Select(_ => new[] { "2026/11/05" })).Append(["003"]).ToArray();
        var table = new TabularData("日付.csv", rows);
        using var data = table.Rows;
        var options = new TabularImportOptions(ImportDestination.NewBook, [[ColumnType.Date]], true);
        var validation = ImportPreviewValidation.Check([table], options, default);
        Check(validation.ErrorCount == 1 && validation.Examples[0].Row == 13 && validation.Cells == 14, "先頭10行の外にある日付変換エラーも全件検出する");
        options = options with { Types = [[ColumnType.TextFormat]] };
        Check(ImportPreviewValidation.Check([table], options, default).ErrorCount == 0, "文字列に戻すと先頭ゼロの値も変換エラーにならない");
        using (var cancellation = new CancellationTokenSource())
        {
            cancellation.Cancel(); bool cancelled = false;
            try { ImportPreviewValidation.Check([table], options, cancellation.Token); } catch (OperationCanceledException) { cancelled = true; }
            Check(cancelled, "全件検査をキャンセルできる");
        }
        using (var renderer = new ExcelFormatPreview())
        using (var form = new TabularImportForm([new("数値.csv", [["数値"], ["1234.56"], ["-0.12"]])], formatPreview: renderer,
            previewContext: (_, _, _, positions) => new(new(), positions.Select(_ => new ExcelPreviewStyle(WidthCharacters: 30)).ToArray())))
        {
            _ = form.Handle;
            var grid = (DataGridView)form.Controls.Find("ImportDataPreview", true).Single(); _ = grid.Handle;
            var digits = (NumericUpDown)form.Controls.Find("ColumnDecimals0", true).Single();
            var type = (ComboBox)form.Controls.Find("ColumnType0", true).Single();
            var code = (TextBox)form.Controls.Find("ColumnFormat0", true).Single();
            string Display() => Convert.ToString(grid[0, 0].Value) ?? "";
            PumpUntil(() => { string shown = Display(); return form.IsReady && form.IsFormattingReady && shown == "1234.56"; }, "初回の書式取得");
            Check(code.ReadOnly && code.Text == "0.00" && Display() == "1234.56", "非表示フォームでもExcelの書式付き表示とコピー可能な書式コードを得る");
            digits.Value = 0;
            PumpUntil(() => Display() == "1235" && form.IsReady && form.IsFormattingReady, "小数0桁への更新");
            Check(code.Text == "0" && Display() == "1235", "小数桁変更直後にコードとExcel表示のキャッシュを更新する");
            type.SelectedIndex = (int)ColumnType.Date;
            PumpUntil(() => form.IsReady && Display() == "変換エラー", "日付への変更");
            Check(code.Text == "yyyy/mm/dd" && !form.IsFormattingReady && !form.TryImport(), "変換不能な日付は画面にも表示し取り込みを止める");
            Check(grid[0, 0].ToolTipText.Contains("1234.56") && grid[0, 0].ToolTipText.Contains("yyyy/mm/dd"), "エラーセルの元データと書式をツールチップで確認できる");
            type.SelectedIndex = (int)ColumnType.TextFormat;
            PumpUntil(() => Display() == "1234.56" && form.IsReady && form.IsFormattingReady, "文字列への修正");
            Check(code.Text == "@" && form.IsFormattingReady, "形式を修正するとエラーと取り込み禁止を解除する");
            var mode = (ComboBox)form.Controls.Find("PreviewMode", true).Single(); mode.SelectedIndex = (int)PreviewMode.None;
            PumpUntil(() => form.IsReady && form.IsFormattingReady, "表示なしの検証");
            Check(grid.RowCount == 0 && form.IsFormattingReady, "プレビューなしでも全件の変換検査が完了する");
        }
        using (var renderer = new ExcelFormatPreview())
        using (var form = new TabularImportForm([table], formatPreview: renderer))
        {
            _ = form.Handle; var mode = (ComboBox)form.Controls.Find("PreviewMode", true).Single(); mode.SelectedIndex = (int)PreviewMode.None;
            var type = (ComboBox)form.Controls.Find("ColumnType0", true).Single(); type.SelectedIndex = (int)ColumnType.Date;
            PumpUntil(() => form.IsReady, "表示なしのエラー検査");
            Check(!form.IsFormattingReady && !form.TryImport() && renderer.OwnedPid == 0, "表示なしでも範囲外エラーを拒否し非表示Excelを不要起動しない");
        }
        using (var host = new Control())
        using (var renderer = new DelayedPreview())
        {
            _ = host.Handle;
            using var controller = new FormattedPreviewController(host, renderer, () => { });
            var number = new TabularData("stale.csv", [["12"]]);
            ExcelPreviewSnapshot Context(IReadOnlyList<(int Row, int Column)> p) => new(new(), p.Select(_ => new ExcelPreviewStyle()).ToArray());
            controller.Reset(number, [ColumnType.Number], [0], Context); controller.Get(0, 0);
            PumpUntil(() => renderer.Calls.Count == 1, "旧要求の開始");
            controller.Reset(number, [ColumnType.TextFormat], [0], Context); controller.Get(0, 0);
            renderer.Calls[0].Completion.SetResult([new(0, 0, "12", "古い結果", 12d, "0", "0", "")]);
            PumpUntil(() => renderer.Calls.Count == 2, "新要求の開始");
            Check(controller.Get(0, 0) is null && renderer.Calls[1].Input.Type == ColumnType.TextFormat, "遅延した旧形式の結果を捨てて新しい列設定を送る");
            renderer.Calls[1].Completion.SetResult([new(0, 0, "12", "新しい結果", "12", "@", "@", "")]);
            PumpUntil(() => controller.Get(0, 0)?.Display == "新しい結果", "新結果の適用");
            Check(!controller.Busy, "設定変更後の結果だけを表示し待機状態を解除する");
        }
    }
    static void PumpUntil(Func<bool> predicate, string name)
    {
        var timer = Stopwatch.StartNew();
        while (!predicate()) { if (timer.Elapsed > TimeSpan.FromSeconds(30)) throw new Exception("待機超過: " + name); Application.DoEvents(); Thread.Sleep(15); }
    }
    sealed class DelayedPreview : IExcelFormatPreview
    {
        internal List<(ExcelPreviewInput Input, TaskCompletionSource<IReadOnlyList<ExcelPreviewCell>> Completion)> Calls { get; } = new();
        public Task<IReadOnlyList<ExcelPreviewCell>> RenderAsync(IReadOnlyList<ExcelPreviewInput> cells, ExcelPreviewEnvironment environment, CancellationToken token)
        { var completion = new TaskCompletionSource<IReadOnlyList<ExcelPreviewCell>>(TaskCreationOptions.RunContinuationsAsynchronously); Calls.Add((cells[0], completion)); return completion.Task; }
        public void Dispose() { }
    }
}
