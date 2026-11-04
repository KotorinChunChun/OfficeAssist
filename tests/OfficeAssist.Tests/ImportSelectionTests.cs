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
}
