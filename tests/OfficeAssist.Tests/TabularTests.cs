using OfficeAssist;
using System.Text;

internal static partial class Program
{
    static void TabularTests(string output)
    {
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
    }
}
