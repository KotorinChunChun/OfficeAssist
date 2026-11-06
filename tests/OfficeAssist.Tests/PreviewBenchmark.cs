using OfficeAssist;
using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

internal static partial class Program
{
    sealed record PreviewTiming(string Name, string Features, int Rows, int Columns, double[] Milliseconds, double MedianMs, long AllocatedBytes, string ResultHash);
    static void PreviewBenchmark(string output, int repeats)
    {
        var results = new List<PreviewTiming>();
        string report = Path.Combine(output, "benchmark.json");
        if (File.Exists(report)) throw new IOException("測定結果は既存ファイルへ上書きしません。");
        void Save() => File.WriteAllText(report, JsonSerializer.Serialize(new { schemaVersion = 1, syntheticData = true, repeats,
            runtime = Environment.Version.ToString(), os = Environment.OSVersion.ToString(), processors = Environment.ProcessorCount,
            source = typeof(ExcelFormatPreview).Assembly.GetCustomAttributes(false).OfType<System.Reflection.AssemblyInformationalVersionAttribute>().Single().InformationalVersion,
            results }, new JsonSerializerOptions { WriteIndented = true }));
        void Measure(string name, string features, int rows, int columns, Func<string> run, int? count = null)
        {
            var samples = new List<double>(); long allocated = 0; string hash = "";
            for (int i = 0; i < (count ?? repeats); i++)
            {
                long before = GC.GetTotalAllocatedBytes(false); var watch = Stopwatch.StartNew(); string value = run(); watch.Stop();
                samples.Add(watch.Elapsed.TotalMilliseconds); allocated += Math.Max(0, GC.GetTotalAllocatedBytes(false) - before);
                string current = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
                if (i > 0 && current != hash) throw new Exception("反復測定で結果が変化しました: " + name); hash = current;
            }
            var sorted = samples.Order().ToArray(); double median = sorted.Length % 2 == 1 ? sorted[sorted.Length / 2] : (sorted[sorted.Length / 2 - 1] + sorted[sorted.Length / 2]) / 2;
            results.Add(new(name, features, rows, columns, samples.ToArray(), median, allocated / samples.Count, hash)); Save();
            Console.WriteLine($"BENCH {name}: median={median:F2} ms [{string.Join(", ", samples.Select(s => s.ToString("F2", CultureInfo.InvariantCulture)))}]");
        }
        foreach (string kind in new[] { "small", "numbers", "dates-repeat", "dates-unique", "invalid-dates", "long-text", "wide" })
        {
            int height = kind == "small" ? 10 : 10000, width = kind == "wide" ? 64 : 8;
            var table = BenchmarkTable(kind, height, width);
            var types = Enumerable.Repeat(kind.StartsWith("dates", StringComparison.Ordinal) || kind == "invalid-dates" ? ColumnType.Date : kind == "long-text" ? ColumnType.TextFormat : ColumnType.Number, width).ToArray();
            var options = new TabularImportOptions(ImportDestination.NewBook, [types], true);
            Measure("raw/" + kind, "書式なし・先頭10行", height, width, () => string.Join("|", table.Rows.Skip(1).Take(10).SelectMany(r => r)));
            Measure("validation/" + kind, "全件検査ON", height, width, () => { var v = ImportPreviewValidation.Check([table], options, default); return $"{v.Cells}/{v.ErrorCount}"; });
            Measure("inference/" + kind, "全件列推定ON", height, width, () => { var v = ColumnInference.Scan(table, true); return string.Join(',', v.Types) + string.Join(',', v.Decimals); });
        }
        var left = BenchmarkTable("numbers", 10000, 8); var right = left with { Name = "second.tsv" };
        foreach (var pair in new[] { ("off", new TableTransformOptions()), ("trim", new TableTransformOptions(1, 1, true)),
            ("merge-order", new TableTransformOptions(Merge: true)), ("merge-name", new TableTransformOptions(Merge: true, MergeByName: true)) })
            Measure("transform/" + pair.Item1, pair.Item1, 20000, 8, () => { var p = TabularTransform.Prepare([left, right], pair.Item2, true); return string.Join("/", p.Tables.Select(t => t.RowCount + "x" + t.ColumnCount)); });
        var many = TabularTransform.Prepare([left, right], new(Merge: true), true).Tables[0];
        foreach (PreviewMode mode in Enum.GetValues<PreviewMode>())
            Measure("preview-rows/" + mode, mode.ToString(), 20000, 9, () => string.Join(',', TabularTransform.PreviewRows(many, true, mode)));
        string file = Path.Combine(output, "synthetic.csv");
        File.WriteAllLines(file, left.Rows.Select(r => string.Join(',', r)), new UTF8Encoding(true));
        foreach (var encoding in new[] { TableEncoding.Auto, TableEncoding.Utf8 })
            Measure("parse/" + encoding, encoding.ToString(), 10000, 8, () => { var parsed = TabularFiles.Read([file], encoding: encoding); try { return parsed[0].RowCount.ToString(); } finally { foreach (var t in parsed) t.Rows.Dispose(); } });

        using var renderer = new ExcelFormatPreview();
        Measure("excel/cold-start", "専用Excel初回起動を含む", 1, 1,
            () => PreviewDigest(renderer.RenderAsync([new(0, 0, "42", ColumnType.Number, 0)], new(), default).GetAwaiter().GetResult()), 1);
        foreach (var item in new[] { ("8-mixed", 1, false, false, false), ("80-mixed", 10, false, false, false),
            ("80-unique", 10, true, false, false), ("80-target-style", 10, true, true, false),
            ("320-target-style", 40, true, true, false), ("80-long-text", 10, true, false, true) })
        {
            var inputs = BenchmarkInputs(item.Item2, item.Item3, item.Item4, item.Item5);
            Measure("excel/" + item.Item1, item.Item4 ? "書式ON・対象属性ON" : "書式ON・対象属性OFF", item.Item2, 8,
                () => PreviewDigest(renderer.RenderAsync(inputs, new(), default).GetAwaiter().GetResult()));
        }
        var existing = Process.GetProcessesByName("EXCEL").Select(p => { using (p) return p.Id; }).ToHashSet();
        using var com = new ComScope(); dynamic app = com.Own(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!);
        Native.GetWindowThreadProcessId((nint)(int)app.Hwnd, out uint pid);
        if (existing.Contains((int)pid)) throw new Exception("試験専用Excelではありません。");
        try
        {
            app.Visible = false; app.DisplayAlerts = false; app.EnableEvents = false;
            dynamic book = com.Own(com.Own(app.Workbooks).Add(-4167)), sheet = com.Own(book.Worksheets.Item(1)), window = com.Own(app.ActiveWindow);
            var data = new TabularData("preview.csv", new[] { Enumerable.Range(0, 8).Select(c => "列" + c).ToArray() }.Concat(Enumerable.Range(0, 10).Select(r => BenchmarkInputs(10, true, false, false).Skip(r * 8).Take(8).Select(c => c.Raw).ToArray())).ToArray());
            var options = TabularImportOptions.Infer([data], header: true) with { BookId = (string)book.FullName, SheetName = (string)sheet.Name, CellAddress = "A1" };
            var positions = Enumerable.Range(1, 10).SelectMany(r => Enumerable.Range(0, 8).Select(c => (Row: r, Column: c))).ToArray();
            foreach (var destination in new[] { ImportDestination.NewBook, ImportDestination.NewSheet, ImportDestination.Selection })
                Measure("context/" + destination, "対象属性の読取りのみ", 10, 8, () => { var v = ExcelPreviewContext.Capture((object)window, [data], options with { Destination = destination }, 0, positions); return JsonSerializer.Serialize(v); });
            foreach (bool formatted in new[] { false, true })
            {
                TabularImportForm? form = null; DataGridView? grid = null;
                try
                {
                    Measure("form/" + (formatted ? "formatted" : "raw"), "初回表示・GUIは非表示", 10, 8, () =>
                    {
                        form?.Dispose(); form = new([data], formatPreview: formatted ? new BorrowedPreview(renderer) : null,
                            previewContext: (tables, o, f, p) => ExcelPreviewContext.Capture(window, tables, options with { Types = o.Types, DecimalPlaces = o.DecimalPlaces }, f, p));
                        _ = form.Handle; grid = (DataGridView)form.Controls.Find("ImportDataPreview", true).Single(); _ = grid.Handle;
                        return BenchmarkFormDisplay(form, grid);
                    });
                    Measure("form/" + (formatted ? "formatted-change" : "raw-change"), "小数桁変更", 10, 8, () =>
                    {
                        var digits = (NumericUpDown)form!.Controls.Find("ColumnDecimals1", true).Single(); digits.Value = digits.Value == 0 ? 2 : 0;
                        _ = BenchmarkFormDisplay(form, grid!); return "ready";
                    });
                }
                finally { form?.Dispose(); }
            }
            book.Close(false);
        }
        finally { app.Quit(); }
        Console.WriteLine("測定結果: " + report);
    }
    static string BenchmarkFormDisplay(TabularImportForm form, DataGridView grid)
    {
        string text = "";
        PumpUntil(() => { var values = grid.Rows.Cast<DataGridViewRow>().SelectMany(r => r.Cells.Cast<DataGridViewCell>().Select(c => Convert.ToString(c.Value) ?? "")).ToArray();
            text = string.Join('|', values); return form.IsReady && form.IsFormattingReady && values.All(v => v != "…"); }, "ベンチマークの表示完了");
        return text;
    }
    static TabularData BenchmarkTable(string kind, int height, int width)
    {
        var rows = new string[height + 1][]; rows[0] = Enumerable.Range(0, width).Select(c => "列" + c).ToArray();
        for (int r = 1; r <= height; r++) rows[r] = Enumerable.Range(0, width).Select(c => kind switch
        { "dates-repeat" => "2026/11/06", "dates-unique" => new DateTime(2000, 1, 1).AddDays(r + c).ToString("yyyy/MM/dd"),
            "invalid-dates" => "不正な日付" + (r % 100), "long-text" => new string('あ', 512) + (r % 100), _ => (r * 8.125 + c).ToString(CultureInfo.InvariantCulture) }).ToArray();
        return new(kind + ".csv", rows);
    }
    static ExcelPreviewInput[] BenchmarkInputs(int rows, bool unique, bool attributes, bool longText)
    {
        ColumnType[] types = [ColumnType.TextFormat, ColumnType.Number, ColumnType.Date, ColumnType.DateWeekday, ColumnType.DateTime, ColumnType.Time, ColumnType.Percent, ColumnType.Duration];
        return Enumerable.Range(0, rows * 8).Select(i => { int n = unique ? i : 0, column = i % 8;
            string day = new DateTime(2026, 1, 1).AddDays(n).ToString("yyyy/MM/dd");
            string value = column switch { 0 => "00" + n, 1 => (1234.56 + n).ToString(CultureInfo.InvariantCulture), 2 or 3 => day, 4 => day + " 08:09:10", 5 => "08:09:10", 6 => "12.5%", _ => "25:30:00" };
            return new ExcelPreviewInput(i / 8, column, longText ? new string('文', 32760) + i : value, longText ? ColumnType.TextFormat : types[column], 2,
                attributes ? new(WidthPoints: 50 + column * 8, FontName: "Yu Gothic UI", FontSize: 10 + i % 3) : new(WidthCharacters: 30)); }).ToArray();
    }
    static string PreviewDigest(IReadOnlyList<ExcelPreviewCell> cells) => JsonSerializer.Serialize(cells);
    sealed class BorrowedPreview(IExcelFormatPreview renderer) : IExcelFormatPreview
    {
        public Task<IReadOnlyList<ExcelPreviewCell>> RenderAsync(IReadOnlyList<ExcelPreviewInput> cells, ExcelPreviewEnvironment environment, CancellationToken token) => renderer.RenderAsync(cells, environment, token);
        public void Dispose() { }
    }
}
