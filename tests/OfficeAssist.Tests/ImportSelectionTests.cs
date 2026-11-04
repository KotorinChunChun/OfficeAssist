using OfficeAssist;

internal static partial class Program
{
    static void ImportSelectionTests()
    {
        TabularData[] tables = [new("first.csv", [["ID", "値", "追加"], ["001", "12", "長い列"]]),
            new("second.tsv", [["ID", "値"], ["002", "文字"]])];
        var selection = new TabularImportSelection(tables);
        selection.SetType(0, 0, ColumnType.TextQuote);
        Check(selection.Types[0][0] == ColumnType.TextQuote && selection.Types[1][0] == ColumnType.TextFormat, "ファイル別の列設定を分離する");
        selection.SetApplyToAll(true, 0);
        Check(selection.Types[1][0] == ColumnType.TextQuote && selection.Types[1][1] == selection.Types[0][1], "全ファイル適用は同じ列番号へ反映する");
        selection.SetType(0, 2, ColumnType.General);
        Check(selection.Types[1].Length == 2, "列数が異なるファイルへ存在しない列を追加しない");
        selection.SetType(1, 1, ColumnType.Date);
        Check(selection.Types[0][1] == ColumnType.Date, "全ファイル適用中の個別変更も反映する");
        selection.SetApplyToAll(false, 1); selection.SetType(1, 1, ColumnType.TextFormat);
        selection.SetHeader(!selection.Header);
        Check(selection.Types[0][1] == ColumnType.Date && selection.Types[1][1] == ColumnType.TextFormat, "全ファイル適用解除と見出し変更で手動設定を保持する");
        var options = selection.CreateOptions(ImportDestination.NewBook); selection.SetType(0, 0, ColumnType.Number);
        Check(options.Types[0][0] == ColumnType.TextQuote, "確定済み取り込み設定を後の画面操作から保護する");

        using var form = new TabularImportForm(tables);
        var file = (ComboBox)form.Controls.Find("ImportFile", true).Single();
        var grid = (DataGridView)form.Controls.Find("ImportDataPreview", true).Single();
        Check(file.Items.Count == 2 && grid.Columns.Count == 3 && grid.ReadOnly && grid.VirtualMode, "選択ファイルの実データを横方向の列で表示する");
        Check((string?)grid.Rows[0].Cells[0].Value == "001" && (string?)grid.Rows[0].Cells[1].Value == "12", "実データを元の列順でプレビューする");
        Check(grid.Columns.Cast<DataGridViewColumn>().All(c => c.SortMode == DataGridViewColumnSortMode.NotSortable), "データ見出しの並べ替えを無効にする");
        var type = (ComboBox)form.Controls.Find("ColumnType0", true).Single(); type.SelectedIndex = (int)ColumnType.TextQuote;
        file.SelectedIndex = 1;
        Check(grid.Columns.Count == 2, "ファイル切替でプレビューの列数を更新する");
        Check((string?)grid.Rows[0].Cells[0].Value == "002", "ファイル切替で実データを更新する");
        file.SelectedIndex = 0;
        Check(((ComboBox)form.Controls.Find("ColumnType0", true).Single()).SelectedIndex == (int)ColumnType.TextQuote, "ファイルを切り替えても列形式を保持する");
    }

    static void LargeImportPreviewTests(string output)
    {
        var tables = TabularFiles.Read([Path.Combine(output,"files","01-wide-40cols.csv"), Path.Combine(output,"files","02-wide-64cols.tsv")]);
        Check(tables[0].ColumnCount == 40 && tables[1].ColumnCount == 64 && tables.All(t => t.RowCount == 1501), "大型サンプルの列数と行数を保持");
        Check(tables[0].Rows[1][0] == "000001" && tables[1].Rows[1500][0] == "001500" && tables[0].Rows[1][4] == "前半1\n後半1", "大型サンプルの先頭ゼロ・最終行・改行を保持");
        using var form = new TabularImportForm(tables);
        form.Size = new Size(1500, 900); _ = form.Handle; form.PerformLayout();
        var files = (ComboBox)form.Controls.Find("ImportFile", true).Single();
        var grid = (DataGridView)form.Controls.Find("ImportDataPreview", true).Single(); _ = grid.Handle;
        Check(grid.Columns.Count == 40 && grid.RowCount == 1500, "大型プレビューを全行の仮想表示で構成");
        files.SelectedIndex = 1; form.PerformLayout();
        grid.HorizontalScrollingOffset = 64 * 240;
        Check(grid.Columns.Count == 64 && (string?)grid.Rows[1499].Cells[0].Value == "001500", "ファイル切替後も64列の最終行へアクセス可能");
        Check(form.Controls.Find("ColumnType63", true).Length == 1 && form.Controls.Find("ColumnType0", true).Length == 0,
            "横スクロールで末尾列の設定だけを表示");
        Check(form.Controls.Find("ColumnType63", true).Single().Parent!.Controls.OfType<ComboBox>().Count() < 10, "大型テーブルの設定コントロール数を表示範囲に制限");
    }
}
