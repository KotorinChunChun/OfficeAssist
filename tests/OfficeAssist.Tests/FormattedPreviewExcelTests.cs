using OfficeAssist;
using System.Diagnostics;

internal static partial class Program
{
    static void FormattedPreviewExcelTests()
    {
        var existing = Process.GetProcessesByName("EXCEL").Select(p => { using (p) return p.Id; }).ToHashSet();
        using var com = new ComScope(); dynamic app = com.Own(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!);
        Native.GetWindowThreadProcessId((nint)(int)app.Hwnd, out uint pid);
        if (existing.Contains((int)pid)) throw new InvalidOperationException("専用Excelを確認できません。");
        app.Visible = false; app.DisplayAlerts = false;
        bool oldSystem = (bool)app.UseSystemSeparators; string oldDecimal = (string)app.DecimalSeparator, oldThousands = (string)app.ThousandsSeparator;
        try
        {
            dynamic books = com.Own(app.Workbooks), book = com.Own(books.Add(-4167)), sheet = com.Own(book.Worksheets.Item(1)), window = com.Own(app.ActiveWindow);
            dynamic wide = com.Own(sheet.Range["A:M"]); wide.ColumnWidth = 36;
            ColumnType[] types = [ColumnType.Number, ColumnType.Number, ColumnType.GroupedNumber, ColumnType.Percent, ColumnType.Yen, ColumnType.Date,
                ColumnType.DateWeekday, ColumnType.DateTime, ColumnType.Time, ColumnType.Duration, ColumnType.TextFormat, ColumnType.TextQuote, ColumnType.General];
            int[] decimals = [0, 2, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0];
            string[] values = ["1234.56", "1234.56", "1234.56", "12.5%", "1234.56", "2026/11/06", "2026/11/06", "2026/11/06 08:09:10", "08:09:10", "25:30:00", "001", "0007", "42.5"];
            var data = new TabularData("preview.csv", [types.Select(t => t.ToString()).ToArray(), values]);
            var options = new TabularImportOptions(ImportDestination.Selection, [types], true) { DecimalPlaces = [decimals], BookId = (string)book.FullName, SheetName = (string)sheet.Name, CellAddress = "A1" };
            ExcelTabularInserter.Import(com, app, window, new[] { data }, options);
            var positions = Enumerable.Range(0, types.Length).Select(c => (Row: 1, Column: c)).ToArray();
            ExcelPreviewSnapshot snapshot = ExcelPreviewContext.Capture(window, new[] { data }, options, 0, positions);
            var inputs = positions.Select((p, i) => new ExcelPreviewInput(p.Row, p.Column, values[i], types[i], decimals[i], snapshot.Styles[i])).ToArray();
            int previewPid;
            using (var renderer = new ExcelFormatPreview())
            {
                var results = renderer.RenderAsync(inputs, snapshot.Environment, default).GetAwaiter().GetResult(); previewPid = renderer.OwnedPid;
                Check(previewPid != pid && previewPid > 0 && !existing.Contains(previewPid), "プレビュー専用Excelは実書込み用・既存Excelと別プロセス");
                for (int c = 0; c < types.Length; c++)
                {
                    dynamic cell = com.Own(sheet.Cells.Item(2, c + 1)); var result = results[c];
                    Check(result.Error is null && result.Display == (string)cell.Text && Equals(result.StoredValue, (object?)cell.Value2) && result.NumberFormat == (string)cell.NumberFormat && result.Prefix == (string)cell.PrefixCharacter,
                        types[c] + ": プレビューの表示・格納値・書式・引用符が実書込みと一致");
                }
                Check(results[0].FormatCode == "0" && results[1].FormatCode == "0.00" && results[5].FormatCode == "yyyy/mm/dd", "適用するExcel書式コードを返す");
                sheet.Range["A41"].Font.Bold = true; sheet.Range["B42"].Font.Size = 22; sheet.Range["B42"].RowHeight = 30;
                var styleData = new TabularData("styles.csv", [["数値1", "数値2"], ["1", "2"], ["3", "4"]]);
                IReadOnlyList<TabularData> styleTables = new[] { styleData };
                var styleOptions = TabularImportOptions.Infer(styleTables, header: true) with { CellAddress = "A40" };
                var styleCache = new ExcelPreviewContextCache((object)window);
                var stylePositions = new[] { (Row: 1, Column: 0), (Row: 2, Column: 1) };
                var mixed = styleCache.Capture(styleTables, styleOptions, 0, stylePositions);
                Check(mixed.Styles[0].Bold && mixed.Styles[1].FontSize == 22 && mixed.Styles[1].RowHeight == 30, "範囲内の属性が混在した場合は個別値を正しく取得する");
                var reused = styleCache.Capture(styleTables, styleOptions with { DecimalPlaces = [[2, 2]] }, 0, stylePositions);
                Check(ReferenceEquals(mixed.Styles[0], reused.Styles[0]), "小数桁だけの変更は同じ出力属性を再取得しない");
                var changedTarget = styleCache.Capture(styleTables, styleOptions with { CellAddress = "D40" }, 0, stylePositions);
                Check(!changedTarget.Styles[0].Bold && changedTarget.Styles[1].FontSize != 22, "指定セル変更では属性キャッシュを破棄する");
                sheet.Range["A:A"].ColumnWidth = 2;
                ExcelPreviewSnapshot narrow = ExcelPreviewContext.Capture(window, new[] { data }, options, 0, new[] { (Row: 1, Column: 0) });
                var clipped = renderer.RenderAsync([new(1, 0, values[0], types[0], decimals[0], narrow.Styles[0])], narrow.Environment, default).GetAwaiter().GetResult()[0];
                Check(clipped.Display == (string)sheet.Range["A2"].Text && clipped.Warning is not null, "列幅不足の#表示も実際のExcelと一致");
                book.Date1904 = true;
                var dateOptions = options with { CellAddress = "A5" }; ExcelTabularInserter.Import(com, app, window, new[] { data }, dateOptions);
                ExcelPreviewSnapshot dateContext = ExcelPreviewContext.Capture(window, new[] { data }, dateOptions, 0, new[] { (Row: 1, Column: 5) });
                var date = renderer.RenderAsync([new(1, 5, values[5], types[5], 0, dateContext.Styles[0])], dateContext.Environment, default).GetAwaiter().GetResult()[0];
                Check(date.Error is null && date.Display == (string)sheet.Range["F6"].Text && Equals(date.StoredValue, (object)sheet.Range["F6"].Value2), "1904年日付方式も表示・格納値が一致");
                var invalid = renderer.RenderAsync([new(0, 0, "003", ColumnType.Date, 0), new(0, 1, "not-number", ColumnType.Number, 2), new(0, 2, "=1+2", ColumnType.TextFormat, 0)], new(), default).GetAwaiter().GetResult();
                Check(invalid[0].Error is not null && invalid[1].Error is not null && invalid[2].Error is null && invalid[2].StoredValue is string text && text == "=1+2", "変換エラーをセル別に返し数式風文字列は実行しない");
                Reject(() => ExcelImportEngine.Validate([new("invalid.csv", [["003", "not-number"]])], new(ImportDestination.NewBook, [[ColumnType.Date, ColumnType.Number]], false)), "プレビューで拒否したデータを実書込みの検査も拒否する");
                using var canceled = new CancellationTokenSource(); canceled.Cancel(); bool stopped = false;
                try { renderer.RenderAsync(inputs, snapshot.Environment, canceled.Token).GetAwaiter().GetResult(); } catch (OperationCanceledException) { stopped = true; }
                Check(stopped, "古い書式要求をキャンセルできる");
                Check((int)book.Worksheets.Count == 1 && (string)sheet.Range["K2"].Value2 == "001", "プレビューで既存ブックのシートとデータを変更しない");
                app.UseSystemSeparators = false; app.DecimalSeparator = ","; app.ThousandsSeparator = ".";
                var regionalOptions = options with { CellAddress = "A9" }; ExcelTabularInserter.Import(com, app, window, new[] { data }, regionalOptions);
                ExcelPreviewSnapshot regional = ExcelPreviewContext.Capture(window, new[] { data }, regionalOptions, 0, new[] { (Row: 1, Column: 2) });
                var regionalCell = renderer.RenderAsync([new(1, 2, values[2], types[2], decimals[2], regional.Styles[0])], regional.Environment, default).GetAwaiter().GetResult()[0];
                Check(regionalCell.Display == (string)sheet.Range["C10"].Text, "Excelの小数点・桁区切りのユーザー設定を再現する");
                dynamic appendSheet = com.Own(book.Worksheets.Add()); appendSheet.Name = "previewappend";
                appendSheet.Range["A1"].Value2 = "ID"; appendSheet.Range["B1"].Value2 = "値"; appendSheet.Range["A2"].Value2 = "既存";
                appendSheet.Range["B:B"].ColumnWidth = 20; appendSheet.Range["B3"].Font.Bold = true; appendSheet.Range["B3"].Font.Size = 18;
                var appendData = new TabularData("previewappend.csv", [["値", "ID"], ["1234.56", "003"]]);
                var appendOptions = TabularImportOptions.Infer([appendData], ImportDestination.AppendSameSheet, true);
                ExcelPreviewSnapshot appendContext = ExcelPreviewContext.Capture(window, new[] { appendData }, appendOptions, 0, new[] { (Row: 1, Column: 0) });
                var appendPreview = renderer.RenderAsync([new(1, 0, "1234.56", appendOptions.Types[0][0], appendOptions.Decimals(0, 0), appendContext.Styles[0])], appendContext.Environment, default).GetAwaiter().GetResult()[0];
                ExcelTabularInserter.Import(com, app, window, new[] { appendData }, appendOptions);
                Check(appendPreview.Error is null && appendPreview.Display == (string)appendSheet.Range["B3"].Text && appendContext.Styles[0].Bold && appendContext.Styles[0].FontSize == 18,
                    "追記の列名対応後の出力セルの幅・太字・サイズも再現する");
                var newSheetOptions = options with { Destination = ImportDestination.NewSheet };
                ExcelPreviewSnapshot newSheetContext = ExcelPreviewContext.Capture(window, new[] { data }, newSheetOptions, 0, new[] { (Row: 1, Column: 7) });
                var newSheetPreview = renderer.RenderAsync([new(1, 7, values[7], types[7], 0, newSheetContext.Styles[0])], newSheetContext.Environment, default).GetAwaiter().GetResult()[0];
                ExcelTabularInserter.Import(com, app, window, new[] { data }, newSheetOptions); dynamic newSheet = com.Own(app.ActiveSheet);
                Check(newSheetPreview.Error is null && newSheetPreview.Display == (string)newSheet.Range["H2"].Text, "新しいシートの既定幅と日付方式を再現する");
                renderer.RenderAsync([new(0, 0, "1234.56", ColumnType.Number, 2)], new(NormalFont: "Arial", NormalFontSize: 20), default).GetAwaiter().GetResult();
                var newBookOptions = options with { Destination = ImportDestination.NewBook };
                ExcelPreviewSnapshot newBookContext = ExcelPreviewContext.Capture(window, new[] { data }, newBookOptions, 0, new[] { (Row: 1, Column: 7) });
                var newBookPreview = renderer.RenderAsync([new(1, 7, values[7], types[7], 0, newBookContext.Styles[0])], newBookContext.Environment, default).GetAwaiter().GetResult()[0];
                ExcelTabularInserter.Import(com, app, window, new[] { data }, newBookOptions); dynamic newBook = com.Own(app.ActiveWorkbook);
                try { dynamic created = com.Own(newBook.Worksheets.Item(1)); Check(newBookPreview.Error is null && newBookPreview.Display == (string)created.Range["H2"].Text && !newBookContext.Environment.Date1904, "新しいブックは前のフォント設定を引き継がず既定の表示を再現する"); }
                finally { newBook.Close(false); }
            }
            bool exited;
            try { using var process = Process.GetProcessById(previewPid); exited = process.WaitForExit(5000); } catch (ArgumentException) { exited = true; }
            Check(exited && (string)sheet.Range["K2"].Value2 == "001", "プレビュー専用Excelだけを終了し参照用Excelを維持する");
            book.Close(false);
        }
        finally { try { app.DecimalSeparator = oldDecimal; app.ThousandsSeparator = oldThousands; app.UseSystemSeparators = oldSystem; app.Quit(); } catch { } }
    }
}
