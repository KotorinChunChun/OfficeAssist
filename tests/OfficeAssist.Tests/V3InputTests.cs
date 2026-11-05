using OfficeAssist;
using System.Diagnostics;

internal static partial class Program
{
    static void V3InputTests(string output)
    {
        string root = Path.Combine(output, "discovery"), sub = Path.Combine(root, "sub"); Directory.CreateDirectory(sub);
        string a = Path.Combine(root, "01.csv"), b = Path.Combine(sub, "02.TSV"), ignored = Path.Combine(sub, "notes.txt");
        File.WriteAllText(a, "ID,値\n001,1"); File.WriteAllText(b, "ID\t値\n002\t2"); File.WriteAllText(ignored, "対象外");
        var found = TabularInputDiscovery.Find([root, sub, a]);
        Check(found.Files.Count == 2 && found.Duplicates > 0 && found.Errors.Count == 0, "フォルダを再帰探索し親子ルートと直接指定の重複を除外する");
        Check(TabularInputDiscovery.Matches(b, "*SUB*?.tsv") && TabularInputDiscovery.Matches(b, "02"), "フルパスを大文字小文字を区別しないワイルドカードと部分一致で検索する");
        var model = new InputSelection(found.Files); var visible = model.Visible("*01.csv"); model.SetAll(visible, false);
        Check(model.Selected.Length == 1 && model.Selected[0].Path == b, "表示分だけ解除して非表示のチェックを保持する");
        model.SetAll(Enumerable.Range(0, model.Files.Count), true); model.Move(1, -1);
        Check(model.Selected[0].Path == b && model.Selected.Length == 2, "チェック状態を保って取込順を変更する");
        using var canceled = new CancellationTokenSource(); canceled.Cancel(); bool stopped = false;
        try { TabularInputDiscovery.Find([root], canceled.Token); } catch (OperationCanceledException) { stopped = true; }
        Check(stopped, "フォルダ探索のキャンセルで部分結果を返さない");
        using var batch = ReceivedTabularBatch.Capture([root]); var read = batch.Read(found.Files, default);
        try { Check(read.Count == 2 && read[1].Rows[1][0] == "002", "チェック済み候補だけを保持して解析する"); }
        finally { foreach (var table in read) table.Rows.Dispose(); }
        TabularData[] data = [new("a.csv", Enumerable.Range(0, 25).Select(i => new[] { i == 0 ? "ID" : i.ToString(), i == 0 ? "値" : "1" }).ToArray()),
            new("b.csv", [["値", "ID"], ["2", "002"]])];
        using var form = new TabularImportForm(data); _ = form.Handle;
        var grid = (DataGridView)form.Controls.Find("ImportDataPreview", true).Single();
        var modes = (ComboBox)form.Controls.Find("PreviewMode", true).Single();
        Check(grid.RowCount == 10, "取り込み画面の初期プレビューは各ファイル10行");
        modes.SelectedIndex = (int)PreviewMode.None;
        Check(grid.RowCount == 0 && form.Tables[0].RowCount == 25, "プレビューしない場合も出力データを保持する");
        modes.SelectedIndex = (int)PreviewMode.All; Check(grid.RowCount == 24, "すべてを選ぶと加工後の全データ行を参照できる");
        var header = (CheckBox)form.Controls.Find("ImportHeader", true).Single(); header.Checked = true; WaitReady(form);
        var merge = (CheckBox)form.Controls.Find("MergeFiles", true).Single(); merge.Checked = true; WaitReady(form);
        var by = (ComboBox)form.Controls.Find("MergeBy", true).Single(); by.SelectedIndex = 1; WaitReady(form);
        Check(form.Tables.Count == 1 && form.Tables[0].Rows[25][1] == "002", "非表示の画面部品から列名結合後の実データを確認する");
        var destination = (ComboBox)form.Controls.Find("ImportDestination", true).Single(); destination.SelectedIndex = (int)ImportDestination.AppendSameSheet;
        Check(destination.SelectedIndex != (int)ImportDestination.AppendSameSheet, "結合ONでは同名追記を選択できない");
    }
    static void WaitReady(TabularImportForm form)
    {
        var timeout = Stopwatch.StartNew();
        while (!form.IsReady && timeout.Elapsed < TimeSpan.FromSeconds(15)) { Application.DoEvents(); Thread.Sleep(10); }
        if (!form.IsReady) throw new Exception("画面モデルの再構築が終了しません。");
    }
}
