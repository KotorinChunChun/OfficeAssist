using System.Diagnostics;
using System.Text.Json;

namespace OfficeAssist;

internal static class TabularImportCli
{
    internal static int Run(string[] args)
    {
        object? application = null;
        bool keepOpen = false;
        using var com = new ComScope();
        try
        {
            int separator = Array.IndexOf(args, "--");
            if (separator < 0 || separator == args.Length - 1) throw new ArgumentException("-- の後にCSV・TSVを指定してください。");
            string? report = null, saveAs = null, typeList = null, decimalsList = null;
            bool show = false; bool? header = null;
            for (int i = 0; i < separator; i++)
            {
                string Value() => ++i < separator ? args[i] : throw new ArgumentException("オプションの値がありません。");
                switch (args[i])
                {
                    case "--import-table": case "--new-book": break;
                    case "--show": show = true; break;
                    case "--report": report = Path.GetFullPath(Value()); break;
                    case "--save-as": saveAs = Path.GetFullPath(Value()); break;
                    case "--types": typeList = Value(); break;
                    case "--decimals": decimalsList = Value(); break;
                    case "--header": header = Value() switch { "yes" => true, "no" => false, "auto" => null, _ => throw new ArgumentException("headerはauto・yes・noです。") }; break;
                    default: throw new ArgumentException("未対応のオプションです: " + args[i]);
                }
            }
            if (!show && saveAs is null) throw new ArgumentException("--show または --save-as を指定してください。");
            if (saveAs is not null && (File.Exists(saveAs) || !saveAs.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("保存先は未使用のxlsxパスを指定してください。");
            if (report is not null && File.Exists(report)) throw new ArgumentException("報告先が既に存在します。");
            var tables = TabularFiles.Read(args[(separator + 1)..]);
            var options = TabularImportOptions.Infer(tables, ImportDestination.NewBook, header);
            if (typeList is not null)
            {
                var types = typeList.Split(',').Select(s => s.Trim() switch
                {
                    "text-format" => ColumnType.TextFormat, "text-quote" => ColumnType.TextQuote,
                    "date" => ColumnType.Date, "general" => ColumnType.General, "number" => ColumnType.Number,
                    "date-weekday" => ColumnType.DateWeekday, "datetime" => ColumnType.DateTime, "time" => ColumnType.Time,
                    "grouped-number" => ColumnType.GroupedNumber, "percent" => ColumnType.Percent, "yen" => ColumnType.Yen, "duration" => ColumnType.Duration,
                    _ => throw new ArgumentException("列形式の指定が不正です。")
                }).ToArray();
                if (tables.Any(t => t.ColumnCount != types.Length)) throw new ArgumentException("列形式の数がファイルの列数と一致しません。");
                options = options with { Types = tables.Select(_ => types.ToArray()).ToArray(), DecimalPlaces = tables.Select(t => Enumerable.Range(0, t.ColumnCount)
                    .Select(c => ColumnConversion.InferDecimals(t.Rows.Skip(options.Header ? 1 : 0).Select(r => c < r.Length ? r[c] : ""), types[c])).ToArray()).ToArray() };
            }
            if (decimalsList is not null)
            {
                int[] decimals = decimalsList.Split(',').Select(s => int.Parse(s, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                if (decimals.Any(d => d is < 0 or > 15) || tables.Any(t => t.ColumnCount != decimals.Length)) throw new ArgumentException("--decimalsは全列分の0～15の整数を指定してください。");
                options = options with { DecimalPlaces = tables.Select(_ => decimals.ToArray()).ToArray() };
            }
            for (int t = 0; t < tables.Count; t++) { tables[t].ExcelValues(options.Types[t], options.Header); for (int c = 0; c < tables[t].ColumnCount; c++) options.Decimals(t, c); }
            var existing = Process.GetProcessesByName("EXCEL").Select(p => { using (p) return p.Id; }).ToHashSet();
            dynamic app = com.Own(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!);
            Native.GetWindowThreadProcessId((nint)(int)app.Hwnd, out uint pid);
            if (existing.Contains((int)pid)) throw new InvalidOperationException("既存Excelを保護するためCLI取り込みを中止しました。");
            application = app;
            app.Visible = false; app.DisplayAlerts = false;
            var result = ExcelTabularInserter.Import(com, app, null, tables, options);
            dynamic book = com.Own(app.ActiveWorkbook), sheets = com.Own(book.Worksheets);
            var sheetReports = new List<object>();
            for (int t = 0; t < tables.Count; t++)
            {
                dynamic sheet = com.Own(sheets.Item(t + 1));
                dynamic first = com.Own(sheet.Range["A1"]), range = com.Own(first.Resize[tables[t].RowCount, tables[t].ColumnCount]);
                object raw = range.Value2;
                var values = raw as Array;
                if (!(range.HasFormula is bool formulas) || formulas) throw new InvalidOperationException("読戻しで意図しない数式を検出しました。");
                dynamic cells = com.Own(range.Cells);
                dynamic columns = com.Own(range.Columns); columns.AutoFit();
                var sample = new List<object>();
                for (int r = 0; r < tables[t].RowCount; r++)
                    for (int c = 0; c < tables[t].ColumnCount; c++)
                    {
                        string source = c < tables[t].Rows[r].Length ? tables[t].Rows[r][c] : "";
                        var type = options.Types[t][c];
                        object? expected = ColumnConversion.Convert(source, type, options.Header && r == 0);
                        if (expected is string text && text.StartsWith('\'')) expected = text[1..];
                        object? actual = values is not null ? values.GetValue(r + values.GetLowerBound(0), c + values.GetLowerBound(1)) : raw;
                        if (!Equals(actual, expected)) throw new InvalidOperationException($"読戻しが一致しません: {sheet.Name} 行{r + 1} 列{c + 1}");
                        if (r < 10 && c < 20)
                        {
                            dynamic cell = com.Own(cells.Item(r + 1, c + 1));
                            sample.Add(new { row = r + 1, column = c + 1, value = actual, format = (string)cell.NumberFormat, prefix = (string)cell.PrefixCharacter, display = (string)cell.Text });
                        }
                    }
                sheetReports.Add(new { name = (string)sheet.Name, rows = tables[t].RowCount, columns = tables[t].ColumnCount,
                    types = options.Types[t].Select(type => ColumnConversion.Labels[(int)type]).ToArray(), decimals = Enumerable.Range(0, tables[t].ColumnCount).Select(c => options.Decimals(t, c)).ToArray(), sample });
            }
            if (saveAs is not null) { Directory.CreateDirectory(Path.GetDirectoryName(saveAs)!); book.SaveAs(saveAs, 51); }
            var output = new { verified = true, pid, workbook = (string)book.FullName, header = options.Header, message = result.Message, sheets = sheetReports };
            if (report is not null)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(report)!);
                File.WriteAllText(report, JsonSerializer.Serialize(output, new JsonSerializerOptions { WriteIndented = true }));
            }
            if (show)
            {
                dynamic firstSheet = com.Own(sheets.Item(1)); firstSheet.Activate();
                dynamic cell = com.Own(firstSheet.Range["A1"]); cell.Select();
                app.DisplayAlerts = true; app.Visible = true; keepOpen = true;
            }
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { if (application is not null && !keepOpen) { try { dynamic app = application; app.Quit(); } catch (Exception) { } } }
    }
}
