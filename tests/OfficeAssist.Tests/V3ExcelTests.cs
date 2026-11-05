using OfficeAssist;

internal static partial class Program
{
    static void V3ExcelTests(ComScope com, dynamic window, dynamic sheet)
    {
        dynamic app = com.Own(window.Application), book = com.Own(sheet.Parent), sheets = com.Own(book.Worksheets);
        var source = new TabularData("追記試験.csv", [["氏名", "金額"], ["山田", "100"]]);
        var options = TabularImportOptions.Infer([source], ImportDestination.AppendSameSheet, true);
        ExcelTabularInserter.Import(com, app, window, new[] { source }, options);
        dynamic added = ExcelImportEngine.ResolveSheet(com, book, "追記試験");
        Check((string)added.Range["A2"].Value2 == "山田", "同名シートがなければ作成して見出し付きで書く");
        var reordered = new TabularData("追記試験.tsv", [["金額", "氏名"], ["200", "佐藤"]]);
        var second = TabularImportOptions.Infer([reordered, reordered], ImportDestination.AppendSameSheet, true);
        ExcelTabularInserter.Import(com, app, window, new[] { reordered, reordered }, second);
        Check((string)added.Range["A3"].Value2 == "佐藤" && (double)added.Range["B4"].Value2 == 200, "同名シートへ列名対応で連続追記し見出しを重複させない");
        var wrong = new TabularData("追記試験.csv", [["未知"], ["値"]]);
        Reject(() => ExcelTabularInserter.Import(com, app, window, new[] { source, wrong }, TabularImportOptions.Infer([source, wrong], ImportDestination.AppendSameSheet, true)), "後続の未知列を最初の書込み前に拒否する");
        Check(added.Range["A5"].Value2 is null && (string)added.Range["A2"].Value2 == "山田", "追記拒否で既存データを保持する");
        Reject(() => ExcelTabularInserter.Import(com, app, window, new[] { source }, options with { Merged = true }), "追記と結合の同時指定を拒否する");
        sheet.Activate();
        var explicitTarget = TabularImportOptions.Infer([source], header: true) with { BookId = (string)book.FullName, SheetName = (string)sheet.Name, CellAddress = "$D$180" };
        ExcelTabularInserter.Import(com, app, window, new[] { source }, explicitTarget);
        Check((string)sheet.Range["D181"].Value2 == "山田", "指定したブック・シート・絶対参照入力のセルへ書く");
        Check(ExcelAddress.Display(com.Own(sheet.Range["D180:E181"])) == "D180:E181", "表示用セルアドレスにドル記号を出さない");
        using var cancellation = new CancellationTokenSource();
        var many = new TabularData("cancel.csv", Enumerable.Range(0, 25000).Select(i => new[] { "001" }).ToArray());
        int chunks = 0;
        string originalFormat = (string)sheet.Range["Z200"].NumberFormat;
        var cancelOptions = TabularImportOptions.Infer([many], header: false) with { CellAddress = "Z200", Cancellation = cancellation.Token, Progress = _ => { if (++chunks == 2) cancellation.Cancel(); } };
        Reject(() => ExcelTabularInserter.Import(com, app, window, new[] { many }, cancelOptions), "分割書込み途中のキャンセルを受け付ける");
        Check(sheet.Range["Z200"].Value2 is null && (string)sheet.Range["Z200"].NumberFormat == originalFormat, "キャンセルで今回の値と書式を復元する");
        added.Range["A10"].Formula = "=\"\"";
        ExcelTabularInserter.Import(com, app, window, new[] { source }, options);
        Check((string)added.Range["A11"].Value2 == "山田", "空文字を返す数式の下へ追記する");
    }
}
