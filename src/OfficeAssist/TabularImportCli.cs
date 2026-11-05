using System.Diagnostics;
using System.Text.Json;

namespace OfficeAssist;

internal static class TabularImportCli
{
    internal static int Run(string[] args)
    {
        object? application = null; bool owned = false, keep = false;
        IReadOnlyList<TabularData>? sources = null;
        using var com = new ComScope();
        try
        {
            if (args.Contains("--list-excel-targets"))
            {
                var targets = new List<object>();
                foreach (var connection in ExcelConnection.Windows())
                {
                    dynamic window = ExcelConnection.Attach(com, connection.Handle), listedApp = com.Own(window.Application), books = com.Own(listedApp.Workbooks);
                    for (int i = 1; i <= (int)books.Count; i++) { dynamic listedBook = com.Own(books.Item(i)), sheets = com.Own(listedBook.Worksheets); var names = new List<string>(); for (int s = 1; s <= (int)sheets.Count; s++) { dynamic sheet = com.Own(sheets.Item(s)); names.Add((string)sheet.Name); } targets.Add(new { instance = connection.Pid, bookId = (string)listedBook.FullName, name = (string)listedBook.Name, sheets = names }); }
                }
                Console.WriteLine(JsonSerializer.Serialize(targets, new JsonSerializerOptions { WriteIndented = true })); return 0;
            }
            int separator = Array.IndexOf(args, "--");
            if (separator < 0) separator = args.Length;
            string? report = null, saveAs = null, typeList = null, decimalsList = null, bookId = null, sheetName = null, cellAddress = null, inputList = null, includePattern = null;
            bool show = false, reportValues = false, listInputs = false; bool? header = null; int instance = 0;
            var destination = ImportDestination.NewBook; var transform = new TableTransformOptions();
            var explicitDestinations = new HashSet<ImportDestination>();
            for (int i = 0; i < separator; i++)
            {
                string Value() => ++i < separator ? args[i] : throw new ArgumentException("オプションの値がありません。");
                switch (args[i])
                {
                    case "--import-table": break;
                    case "--new-book": destination = ImportDestination.NewBook; explicitDestinations.Add(destination); break;
                    case "--new-sheet": destination = ImportDestination.NewSheet; explicitDestinations.Add(destination); break;
                    case "--append-same-sheet": destination = ImportDestination.AppendSameSheet; explicitDestinations.Add(destination); break;
                    case "--cell": cellAddress = Value(); destination = ImportDestination.Selection; explicitDestinations.Add(destination); break;
                    case "--excel-instance": instance = int.Parse(Value()); break;
                    case "--book-id": bookId = Value(); break;
                    case "--sheet": sheetName = Value(); break;
                    case "--show": show = true; break;
                    case "--report-values": reportValues = true; break;
                    case "--list-inputs": listInputs = true; break;
                    case "--input-list": inputList = Value(); break;
                    case "--include-pattern": includePattern = Value(); break;
                    case "--report": report = Path.GetFullPath(Value()); break;
                    case "--save-as": saveAs = Path.GetFullPath(Value()); break;
                    case "--types": typeList = Value(); break;
                    case "--decimals": decimalsList = Value(); break;
                    case "--skip-first": transform = transform with { SkipFirst = int.Parse(Value()) }; break;
                    case "--skip-last": transform = transform with { SkipLast = int.Parse(Value()) }; break;
                    case "--drop-empty-rows": transform = transform with { DropEmpty = true }; break;
                    case "--merge": transform = transform with { Merge = true }; break;
                    case "--merge-by": transform = transform with { MergeByName = Value() switch { "name" => true, "order" => false, _ => throw new ArgumentException("merge-byはorder・nameです。") } }; break;
                    case "--preview": _ = Value() switch { "none" => PreviewMode.None, "first10" => PreviewMode.First10, "all" => PreviewMode.All, _ => throw new ArgumentException("previewはnone・first10・allです。") }; break;
                    case "--header": header = Value() switch { "yes" => true, "no" => false, "auto" => null, _ => throw new ArgumentException("headerはauto・yes・noです。") }; break;
                    default: throw new ArgumentException("未対応のオプションです: " + args[i]);
                }
            }
            if (explicitDestinations.Count > 1) throw new ArgumentException("展開先は1種類だけ指定してください。");
            if (destination == ImportDestination.AppendSameSheet && transform.Merge) throw new ArgumentException("追記と結合は同時に指定できません。");
            string[] inputFiles = inputList is null ? separator < args.Length ? args[(separator + 1)..] : [] : JsonSerializer.Deserialize<string[]>(File.ReadAllText(inputList)) ?? throw new ArgumentException("input-listはフルパスのJSON配列です。");
            if (inputFiles.Length == 0) throw new ArgumentException("CSV・TSVかフォルダ、または--input-listを指定してください。");
            if (listInputs || inputFiles.Any(Directory.Exists))
            {
                var discovery = TabularInputDiscovery.Find(inputFiles);
                if (listInputs) { Console.WriteLine(JsonSerializer.Serialize(discovery, new JsonSerializerOptions { WriteIndented = true })); return discovery.Errors.Count == 0 ? 0 : 1; }
                if (discovery.Errors.Count != 0) throw new IOException("フォルダ検索で問題が発生しました。--list-inputsで詳細を確認してください。");
                inputFiles = discovery.Files.Select(f => f.Path).ToArray();
            }
            inputFiles = inputFiles.Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).Where(p => includePattern is null || TabularInputDiscovery.Matches(p, includePattern)).ToArray();
            if (destination == ImportDestination.NewBook && !show && saveAs is null) throw new ArgumentException("--show または --save-as を指定してください。");
            if (saveAs is not null && (File.Exists(saveAs) || !saveAs.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("保存先は未使用のxlsxパスを指定してください。");
            if (report is not null && File.Exists(report)) throw new ArgumentException("報告先が既に存在します。");
            sources = TabularFiles.Read(inputFiles);
            bool hasHeader = header ?? TabularImportOptions.Infer(TabularTransform.Prepare(sources, transform with { Merge = false, MergeByName = false }, false).Tables).Header;
            var prepared = TabularTransform.Prepare(sources, transform, hasHeader); var tables = prepared.Tables;
            var options = TabularImportOptions.Infer(tables, destination, hasHeader) with { BookId = bookId, SheetName = sheetName, CellAddress = cellAddress, Merged = transform.Merge };
            int offset = transform.Merge ? 1 : 0;
            if (typeList is not null)
            {
                var types = typeList.Split(',').Select(s => s.Trim() switch
                {
                    "text-format" => ColumnType.TextFormat, "text-quote" => ColumnType.TextQuote, "date" => ColumnType.Date, "general" => ColumnType.General, "number" => ColumnType.Number,
                    "date-weekday" => ColumnType.DateWeekday, "datetime" => ColumnType.DateTime, "time" => ColumnType.Time, "grouped-number" => ColumnType.GroupedNumber, "percent" => ColumnType.Percent, "yen" => ColumnType.Yen, "duration" => ColumnType.Duration,
                    _ => throw new ArgumentException("列形式の指定が不正です。")
                }).ToArray();
                if (tables.Any(t => t.ColumnCount - offset != types.Length)) throw new ArgumentException("列形式の数が出力データ列数と一致しません。");
                options = options with { Types = tables.Select(_ => offset == 1 ? new[] { ColumnType.TextFormat }.Concat(types).ToArray() : types.ToArray()).ToArray() };
            }
            if (decimalsList is not null)
            {
                int[] decimals = decimalsList.Split(',').Select(int.Parse).ToArray();
                if (decimals.Any(d => d is < 0 or > 15) || tables.Any(t => t.ColumnCount - offset != decimals.Length)) throw new ArgumentException("--decimalsは出力データ全列分の0～15の整数を指定してください。");
                options = options with { DecimalPlaces = tables.Select(_ => offset == 1 ? new[] { 0 }.Concat(decimals).ToArray() : decimals.ToArray()).ToArray() };
            }
            object? sourceWindow = null; dynamic app;
            if (destination == ImportDestination.NewBook)
            {
                var existing = Process.GetProcessesByName("EXCEL").Select(p => { using (p) return p.Id; }).ToHashSet();
                app = com.Own(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!);
                Native.GetWindowThreadProcessId((nint)(int)app.Hwnd, out uint processId);
                if (existing.Contains((int)processId)) throw new InvalidOperationException("既存Excelを保護するため中止しました。");
                application = app; owned = true; app.Visible = false; app.DisplayAlerts = false;
            }
            else
            {
                if (instance == 0 || bookId is null) throw new ArgumentException("既存ブックには--excel-instanceと--book-idが必要です。");
                var connection = ExcelConnection.Windows().SingleOrDefault(w => w.Pid == instance);
                if (connection.Handle == 0) throw new InvalidOperationException("指定Excelが見つかりません。");
                sourceWindow = ExcelConnection.Attach(com, connection.Handle); dynamic window = sourceWindow; app = com.Own(window.Application); application = app;
            }
            IReadOnlyList<ExcelImportEngine.Write>? writes = null;
            options = options with { Completed = result => writes = result };
            var result = ExcelTabularInserter.Import(com, app, sourceWindow, tables, options);
            if (writes is null) throw new InvalidOperationException("出力結果を取得できません。");
            var reports = Verify(com, writes, reportValues);
            dynamic book = com.Own(((dynamic)writes[0].Sheet).Parent);
            if (saveAs is not null) { Directory.CreateDirectory(Path.GetDirectoryName(saveAs)!); book.SaveAs(saveAs, 51); }
            Native.GetWindowThreadProcessId((nint)(int)app.Hwnd, out uint pid);
            var output = new { verified = true, pid, workbook = (string)book.FullName, header = options.Header, message = result.Message, notices = prepared.Notices, sheets = reports };
            if (report is not null) { Directory.CreateDirectory(Path.GetDirectoryName(report)!); File.WriteAllText(report, JsonSerializer.Serialize(output, new JsonSerializerOptions { WriteIndented = true })); }
            if (show) { app.Visible = true; if (owned) app.DisplayAlerts = true; keep = true; }
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { if (sources is not null) foreach (var table in sources) table.Rows.Dispose(); if (application is not null && owned && !keep) try { dynamic app = application; app.Quit(); } catch { } }
    }
    internal static List<object> Verify(ComScope com, IReadOnlyList<ExcelImportEngine.Write> writes, bool reportValues)
    {
        var reports = new List<object>();
        foreach (var write in writes)
        {
            dynamic sheet = write.Sheet; var sample = new List<object>();
            int step = Math.Max(1, 20000 / write.Table.ColumnCount);
            for (int start = 0; start < write.Table.RowCount; start += step)
            {
                int count = Math.Min(step, write.Table.RowCount - start);
                dynamic first = com.Own(sheet.Cells.Item(write.Row + start, write.Column)), range = com.Own(first.Resize[count, write.Table.ColumnCount]);
                object? raw = range.Value2; var values = raw as Array;
                if (range.HasFormula is not bool formulas || formulas) throw new InvalidOperationException("読戻しで意図しない数式を検出しました。");
                var expected = write.Table.ExcelValues(write.Types, write.Header, start, count);
                dynamic dateBook = com.Own(sheet.Parent); ExcelImportEngine.AdjustDateSystem(expected, write.Types, (bool)dateBook.Date1904);
                for (int r = 0; r < count; r++) for (int c = 0; c < write.Table.ColumnCount; c++)
                {
                    object? value = expected[r, c]; if (value is string text && text.StartsWith('\'')) value = text[1..];
                    object? actual = values is not null ? values.GetValue(r + 1, c + 1) : raw;
                    if (!Equals(value, actual)) throw new InvalidOperationException($"読戻しが一致しません: {sheet.Name} 行{write.Row + start + r} 列{write.Column + c}");
                    if (reportValues && start + r < 10 && c < 20) { dynamic cell = com.Own(sheet.Cells.Item(write.Row + start + r, write.Column + c)); sample.Add(new { row = write.Row + start + r, column = write.Column + c, value = actual, format = (string)cell.NumberFormat, prefix = (string)cell.PrefixCharacter, display = (string)cell.Text }); }
                }
            }
            reports.Add(new { name = (string)sheet.Name, startRow = write.Row, startColumn = write.Column, rows = write.Table.RowCount, columns = write.Table.ColumnCount, types = write.Types.Select(t => ColumnConversion.Labels[(int)t]).ToArray(), decimals = write.Decimals, sample });
        }
        return reports;
    }
}
