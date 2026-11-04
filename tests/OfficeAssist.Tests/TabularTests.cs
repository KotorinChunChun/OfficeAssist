using OfficeAssist;
using System.Text;

internal static partial class Program
{
    static void TabularTests(string output)
    {
        Check(ColumnConversion.Infer(["001", "002"]) == ColumnType.TextFormat && ColumnConversion.Infer(["1234567890123456"]) == ColumnType.TextFormat, "先頭ゼロと長い番号は文字列書式を推定");
        Check(ColumnConversion.Infer(["12.5", "-3"]) == ColumnType.Number && ColumnConversion.Infer(["2026/11/05"]) == ColumnType.Date, "数値列と日付列を推定");
        Check(ColumnConversion.Infer(["12", "不明"]) == ColumnType.TextFormat, "混在列は文字列書式を推定");
        Reject(() => ColumnConversion.Convert("不明", ColumnType.Number, false), "数値の明示指定で不正値を拒否");
        var table = TabularFiles.Parse("名前,番号,備考\r\nちゅん,001,\"改行\r\nと,区切り\"\r\n", ',');
        Check(table.Length == 2 && table[1].SequenceEqual(["ちゅん", "001", "改行\r\nと,区切り"]), "CSVの日本語・引用符内改行・カンマを保持");
        Check(TabularFiles.Parse("\"a\"\"b\"\t\"c\td\"\t", '\t')[0].SequenceEqual(["a\"b", "c\td", ""]), "TSVの引用符と末尾空欄を保持");
        Check(TabularFiles.Parse("a,b\n\nc\n", ',').Select(r => r.Length).SequenceEqual([2, 1, 1]), "空行と不揃いの列を保持");
        Reject(() => TabularFiles.Parse("\"未終了", ','), "閉じない引用符を拒否");
        Reject(() => TabularFiles.Parse("\"a\"b", ','), "引用符後の不正文字を拒否");
        Reject(() => TabularFiles.Parse("a\0b", ','), "バイナリを拒否");
        Reject(() => TabularFiles.Parse(new string('a', 32768), ','), "セル文字数の超過を拒否");
        Reject(() => TabularFiles.Parse(new string(',', 16384), ','), "列数の超過を拒否");
        Check(TabularFiles.Decode(Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes("日本語")).ToArray()) == "日本語", "UTF-8 BOMを除外");
        Check(TabularFiles.Decode(Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("日本語")).ToArray()) == "日本語", "UTF-16 BOMを判定");
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Check(TabularFiles.Decode(Encoding.GetEncoding(932).GetBytes("日本語")) == "日本語", "Shift_JISを判定");
        var data = new DataObject(); data.SetData(DataFormats.FileDrop, false, new[] { "a.CSV", "b.tsv" });
        Check(ExcelDropData.Read(data, DragDropEffects.Copy)?.Length == 2, "CSV・TSVを独自ドロップ対象にする");
        Check(ExcelDropData.Read(data, DragDropEffects.Move) is null, "CSV・TSVの移動は受け付けない");
        Check(!TabularFiles.Accepts(["a.csv", "image.png"]), "画像とテキストの混在を拒否");
        string csv = Path.Combine(output, "fixture.csv"), tsv = Path.Combine(output, "fixture.tsv"), broken = Path.Combine(output, "broken.csv");
        File.WriteAllText(csv, "日本語,001", Encoding.GetEncoding(932));
        File.WriteAllText(tsv, "a\tb\n1\t2", new UTF8Encoding(true));
        File.WriteAllText(broken, "\"未終了");
        var received = TabularFiles.Read([csv, tsv]);
        Check(received.Count == 2 && received[0].Rows[0][0] == "日本語" && received[1].Rows[1][1] == "2", "複数CSV・TSVを原本順に解析する");
        Reject(() => TabularFiles.Read([csv, broken]), "後続ファイルの解析失敗でも一括受信しない");
        using var dispatcher = new Control();
        int failures = 0;
        using var service = new ExcelDropService(dispatcher, (_, failure) => { if (failure) failures++; }, startTimer: false);
        var queue = typeof(ExcelDropService).GetMethod("QueueInsert", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!;
        Check(queue.Invoke(service, [new[] { broken }, Point.Empty, 0]) is false && failures == 1 && !service.Busy, "CSV解析失敗をDrop受信経路から漏らさない");
    }

    static void TabularExcelTests(ComScope com, dynamic window, dynamic sheet)
    {
        dynamic selected = com.Own(sheet.Range["J30:L32,N30:P32"]); selected.Select();
        var input = new TabularData("test.csv", [["001", "=1+2", "2026/11/04"], ["'quote", "日本語", ""]]);
        var result = ExcelTabularInserter.InsertIntoWindow(com, window, new[] { input });
        dynamic actual = com.Own(sheet.Range["J30:L31"]);
        object[,] values = (object[,])actual.Value2;
        Check((string)values[1, 1] == "001" && (string)values[1, 2] == "=1+2" && (string)values[2, 1] == "'quote", "実Excelで先頭ゼロ・数式風文字列・引用符を保持");
        Check(actual.HasFormula is bool hasFormula && !hasFormula, "実ExcelでCSV内容を数式実行しない");
        dynamic textCell = com.Own(sheet.Range["J30"]);
        Check((string)textCell.NumberFormat == "@" && (string)textCell.PrefixCharacter == "", "文字列書式方式は先頭ゼロにシングルクォートを付けない");
        dynamic remaining = com.Own(window.RangeSelection);
        Check((string)remaining.Address == "$N$30:$P$32", "表を展開した選択範囲を除外");
        remaining.Select();
        var overflow = ExcelTabularInserter.InsertIntoWindow(com, window, new[] { input, input });
        dynamic lower = com.Own(sheet.Range["N33"]);
        Check(overflow.Warning is not null && (string)lower.Value2 == "001", "複数CSVは最後の展開先の下へ順に配置");
        selected.Select();
        Reject(() => ExcelTabularInserter.InsertIntoWindow(com, window, new[] { input }), "実Excelの既存データを上書きしない");
        dynamic empty = com.Own(sheet.Range["A50"]); empty.Select(); sheet.Protect();
        try { Reject(() => ExcelTabularInserter.InsertIntoWindow(com, window, new[] { input }), "実Excelの保護シートを拒否"); }
        finally { sheet.Unprotect(); }
        dynamic overlap = com.Own(sheet.Range["A60,C60"]); overlap.Select();
        Reject(() => ExcelTabularInserter.InsertIntoWindow(com, window, new[] { input, input }), "重なるCSV展開先は書込み前に拒否");
        dynamic untouched = com.Own(sheet.Range["A60"]);
        Check(untouched.Value2 is null, "展開先の検査失敗で先頭の表も書き込まない");
        dynamic edge = com.Own(sheet.Range["XFD1048576"]); edge.Select();
        Reject(() => ExcelTabularInserter.InsertIntoWindow(com, window, new[] { input }), "シート端からはみ出すCSVを拒否");
        TabularModeTests(com, window, sheet);
    }

    static void TabularModeTests(ComScope com, dynamic window, dynamic sheet)
    {
        var input = new TabularData("列形式.csv", [["番号", "日付", "金額", "文字列", "標準", "引用符方式"],
            ["001", "2026/11/05", "123.5", "=1+2", "42", "0007"], ["002", "2026/11/06", "-2", "日本語", "text", "0008"]]);
        var inferred = TabularImportOptions.Infer([input], ImportDestination.NewBook);
        Check(inferred.Header && inferred.Types[0][0] == ColumnType.TextFormat && inferred.Types[0][1] == ColumnType.Date && inferred.Types[0][2] == ColumnType.Number, "見出しを除外して列の既定形式を推定");
        inferred.Types[0][4] = ColumnType.General; inferred.Types[0][5] = ColumnType.TextQuote;
        dynamic app = com.Own(window.Application);
        ExcelTabularInserter.Import(com, app, window, new[] { input }, inferred);
        dynamic book = com.Own(app.ActiveWorkbook);
        try
        {
            dynamic imported = com.Own(book.ActiveSheet);
            dynamic id = com.Own(imported.Range["A2"]), date = com.Own(imported.Range["B2"]), number = com.Own(imported.Range["C2"]);
            Check((string)id.Value2 == "001" && (string)id.PrefixCharacter == "" && (string)id.NumberFormat == "@", "新しいブックの文字列は引用符なしの文字列書式");
            Check(date.Value2 is double && number.Value2 is double d && Math.Abs(d - 123.5) < 0.0001, "日付と数値はExcelの数値で格納");
            dynamic quote = com.Own(imported.Range["F2"]), general = com.Own(imported.Range["E2"]), formulaText = com.Own(imported.Range["D2"]);
            Check((string)quote.Value2 == "0007" && (string)quote.PrefixCharacter == "'", "明示したシングルクォート方式を維持");
            Check(general.Value2 is double && !(bool)formulaText.HasFormula && (string)formulaText.PrefixCharacter == "", "標準の数値変換と数式風文字列の非実行");
        }
        finally { book.Close(false); }
        dynamic original = com.Own(sheet.Parent), sheets = com.Own(original.Worksheets);
        int before = (int)sheets.Count;
        ExcelTabularInserter.Import(com, app, window, new[] { input }, inferred with { Destination = ImportDestination.NewSheet });
        Check((int)sheets.Count == before + 1, "既存ブックに新しいシートを作成して取り込む");
        dynamic invalid = com.Own(sheet.Range["T70"]); sheet.Activate(); invalid.Select();
        string oldFormat = (string)invalid.NumberFormat;
        var wrong = new TabularImportOptions(ImportDestination.Selection, [[ColumnType.Number]], false);
        Reject(() => ExcelTabularInserter.InsertIntoWindow(com, window, new[] { new TabularData("bad.csv", [["不正"]]) }, wrong), "列形式の不一致は書込み前に拒否");
        Check(invalid.Value2 is null && (string)invalid.NumberFormat == oldFormat, "変換拒否で値と書式を変更しない");
    }
}
