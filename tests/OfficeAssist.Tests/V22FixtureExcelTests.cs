using OfficeAssist;

internal static partial class Program
{
    static void V22FixtureExcelTests(ComScope com, dynamic window, dynamic sheet, string output)
    {
        string root = CasesRoot(output);
        dynamic app = com.Own(window.Application), book = com.Own(sheet.Parent);
        var tables = TabularFiles.Read([Path.Combine(root, "edge/09-trim.csv"), Path.Combine(root, "edge/10-trim-swapped.tsv")]);
        try
        {
            var merged = TabularTransform.Prepare(tables, new(2, 1, true, true, true), true).Tables;
            var options = TabularImportOptions.Infer(merged, header: true) with { BookId = (string)book.FullName, SheetName = (string)sheet.Name, CellAddress = "A400", Merged = true };
            ExcelTabularInserter.Import(com, app, window, merged, options);
            Check((string)sheet.Range["B401"].Value2 == "001" && (string)sheet.Range["B403"].Value2 == "003" && (string)sheet.Range["D403"].Value2 == "追加", "実ファイルの文字コード混在・行除外・列名和集合をExcelへ展開する");
            Reject(() => ExcelTabularInserter.Import(com, app, window, merged, options with { CellAddress = "A410,F410" }), "結合した表への複数アンカー指定を拒否する");
            sheet.Range["A420"].Value2 = "既存";
            Reject(() => ExcelTabularInserter.Import(com, app, window, merged, options with { CellAddress = "A420" }), "実ファイル結合後の出力範囲全体の衝突を検査する");
            Check((string)sheet.Range["A420"].Value2 == "既存" && sheet.Range["B421"].Value2 is null, "結合の衝突拒否で既存値と空セルを保持する");
        }
        finally { foreach (var table in tables) table.Rows.Dispose(); }
        var sameName = TabularFiles.Read([Path.Combine(root, "folders/a/common.csv"), Path.Combine(root, "folders/b/common.csv")]);
        try
        {
            ExcelTabularInserter.Import(com, app, window, sameName, TabularImportOptions.Infer(sameName, ImportDestination.AppendSameSheet, true));
            dynamic target = ExcelImportEngine.ResolveSheet(com, book, "common");
            Check((string)target.Range["A2"].Value2 == "001" && (string)target.Range["A3"].Value2 == "002", "同名別パスの2実ファイルを同じシートへ順次追記する");
            var unknown = TabularFiles.Read([Path.Combine(root, "edge/12-unknown-append.csv")]);
            try
            {
                var changed = unknown[0] with { Name = "common.csv" };
                Reject(() => ExcelTabularInserter.Import(com, app, window, new[] { changed }, TabularImportOptions.Infer([changed], ImportDestination.AppendSameSheet, true)), "未知列を持つ実ファイルの追記を拒否する");
                Check(target.Range["A4"].Value2 is null && (string)target.Range["A3"].Value2 == "002", "未知列拒否で先に追記したデータを保持する");
            }
            finally { foreach (var table in unknown) table.Rows.Dispose(); }
        }
        finally { foreach (var table in sameName) table.Rows.Dispose(); }
        var empty = TabularFiles.Read([Path.Combine(root, "edge/14-empty-space.csv")]);
        try
        {
            var filtered = TabularTransform.Prepare(empty, new(DropEmpty: true), true).Tables;
            Check(filtered[0].RowCount == 3 && filtered[0].Rows[2][1] == " ", "実ファイルの空欄行は除外し空白値を残す");
            var options = TabularImportOptions.Infer(filtered, header: true) with { BookId = (string)book.FullName, SheetName = (string)sheet.Name, CellAddress = "G440" };
            ExcelTabularInserter.Import(com, app, window, filtered, options);
            Check((string)sheet.Range["G442"].Value2 == "002" && (string)sheet.Range["H442"].Value2 == " ", "空欄除外後の先頭ゼロと空白値をExcelで保持する");
        }
        finally { foreach (var table in empty) table.Rows.Dispose(); }
        var literal = TabularFiles.Read([Path.Combine(root, "edge/16-formula-like.csv")]);
        try
        {
            var options = TabularImportOptions.Infer(literal, header: true) with { BookId = (string)book.FullName, SheetName = (string)sheet.Name, CellAddress = "M440" };
            ExcelTabularInserter.Import(com, app, window, literal, options);
            Check((string)sheet.Range["M441"].Value2 == "0000000000000001" && (string)sheet.Range["N441"].Value2 == "=1+2" && !(bool)sheet.Range["N441"].HasFormula && (string)sheet.Range["N442"].Value2 == "'quoted", "実ファイルの長い番号・数式風文字列・先頭引用符を保持する");
        }
        finally { foreach (var table in literal) table.Rows.Dispose(); }
    }
}
