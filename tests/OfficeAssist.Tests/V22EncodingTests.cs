using OfficeAssist;

internal static partial class Program
{
    static string CasesRoot(string output) => Path.Combine(output, "import-cases");
    static void V22EncodingTests(string output)
    {
        string root = CasesRoot(output); var fixtures = ImportFixtureSet.Create(root);
        foreach (var fixture in fixtures.Where(f => f.RelativePath.StartsWith("valid/") || f.RelativePath.StartsWith("manual/")))
        {
            string path = Path.Combine(root, fixture.RelativePath); var manual = TabularFiles.Read([path], encoding: fixture.Encoding);
            try
            {
                Check(manual[0].Rows[1][0] == "001" && manual[0].Rows[1][1].StartsWith(fixture.RelativePath.Contains("emoji") ? "🕊" : "日本語"), fixture.RelativePath + ": 手動指定で日本語と先頭ゼロを保持");
                if (fixture.Auto)
                {
                    var automatic = TabularFiles.Read([path]);
                    try { Check(automatic[0].Rows.Select(r => string.Join("\u001f", r)).SequenceEqual(manual[0].Rows.Select(r => string.Join("\u001f", r))), fixture.RelativePath + ": 自動判定と手動指定の全行が一致"); }
                    finally { foreach (var table in automatic) table.Rows.Dispose(); }
                }
                else Reject(() => TabularFiles.Read([path]), fixture.RelativePath + ": BOMなしUTF-16は自動で無理に解釈しない");
            }
            finally { foreach (var table in manual) table.Rows.Dispose(); }
        }
        string utf8 = Path.GetFullPath(Path.Combine(root, "valid/01-utf8-bom.csv"));
        string tsv = Path.GetFullPath(Path.Combine(root, "valid/02-utf8-no-bom.tsv"));
        Reject(() => TabularFiles.Read([utf8], encoding: TableEncoding.Cp932), "BOMと手動指定が異なる場合は置換せず拒否する");
        Reject(() => TabularFiles.Read([Path.Combine(root, "reject/17-invalid-utf8.csv")], encoding: TableEncoding.Utf8), "不正UTF-8バイトを置換文字にせず拒否する");
        Reject(() => TabularFiles.Read([Path.Combine(root, "reject/18-empty.csv")]), "空ファイルを取り込まない");
        Reject(() => TabularFiles.Read([Path.Combine(root, "reject/15-malformed.csv")]), "構文不正を文字コード変更で通さない");
        var map = new Dictionary<string, TableEncoding>(StringComparer.OrdinalIgnoreCase);
        var snapshots = new List<IReadOnlyList<TabularData>>(); var initial = TabularFiles.Read([utf8, tsv]); snapshots.Add(initial); int warnings = 0;
        try
        {
            using var form = new TabularImportForm(initial, showMessage: (_, _, _) => warnings++, readEncoding: path =>
            {
                var values = new[] { utf8, tsv }.Where(p => path is null || p == path).Select(p => map.GetValueOrDefault(p)).Distinct().ToArray(); return values.Length == 1 ? (int)values[0] : 5;
            }, reencode: (path, choice) =>
            {
                var next = new Dictionary<string, TableEncoding>(map, StringComparer.OrdinalIgnoreCase);
                foreach (string file in new[] { utf8, tsv }) if (path is null || path == file) next[file] = choice;
                var loaded = TabularFiles.Read([utf8, tsv], encodings: next); snapshots.Add(loaded); map = next; return loaded;
            });
            _ = form.Handle;
            var encoding = (ComboBox)form.Controls.Find("ImportEncoding", true).Single();
            Check(encoding.SelectedIndex == 0 && encoding.Items.Count == 6, "取り込み画面は自動判定を初期選択し手動4種と混在表示を持つ");
            encoding.SelectedIndex = (int)TableEncoding.Utf8; WaitReady(form);
            Check(map[utf8] == TableEncoding.Utf8 && !map.ContainsKey(tsv), "文字コードを選択中のファイルだけに適用する");
            ((CheckBox)form.Controls.Find("ApplyToAllFiles", true).Single()).Checked = true;
            Check(encoding.SelectedIndex == 5, "ファイルごとの指定が異なると混在と表示する");
            encoding.SelectedIndex = (int)TableEncoding.Cp932;
            Check(warnings == 1 && map[utf8] == TableEncoding.Utf8, "誤った手動指定のエラーで元の入力と設定を保持する");
            encoding.SelectedIndex = (int)TableEncoding.Utf8; WaitReady(form);
            Check(map[tsv] == TableEncoding.Utf8 && form.Tables[1].Rows[1][1] == "日本語", "全ファイルへ手動指定して再解析できる");
        }
        finally { foreach (var snapshot in snapshots) foreach (var table in snapshot) table.Rows.Dispose(); }
        var discovered = TabularInputDiscovery.Find([Path.Combine(root, "folders")]);
        var selection = new InputSelection(discovered.Files); selection.SetEncoding([1], TableEncoding.Utf8); selection.Move(1, -1);
        Check(selection.Selected.Length == 2 && selection.Selected[0].Encoding == TableEncoding.Utf8, "同名別パスを残し文字コード指定も順序変更について移動する");
        var filtered = TabularFiles.Read([Path.Combine(root, "edge/09-trim.csv"), Path.Combine(root, "edge/10-trim-swapped.tsv")]);
        try
        {
            var prepared = TabularTransform.Prepare(filtered, new(2, 1, true, true, true), true).Tables[0];
            Check(prepared.RowCount == 4 && prepared.ColumnCount == 4 && prepared.Rows[3][1] == "003" && prepared.Rows[3][3] == "追加", "混在文字コードの実ファイルで行除外と列名結合を組み合わせる");
            foreach (var mode in Enum.GetValues<PreviewMode>()) Check(TabularTransform.PreviewRows(prepared, true, mode).Length == (mode == PreviewMode.None ? 0 : 3) && prepared.RowCount == 4, "結合後プレビュー" + mode + "で出力件数を変えない");
            Reject(() => TabularTransform.Prepare(filtered, new(100, 0), true), "実ファイルの全行除外を拒否する");
        }
        finally { foreach (var table in filtered) table.Rows.Dispose(); }
        var duplicate = TabularFiles.Read([Path.Combine(root, "reject/11-duplicate-headers.csv")]);
        try { Reject(() => TabularTransform.Prepare(duplicate, new(Merge: true, MergeByName: true), true), "実ファイルの重複見出しを列名結合前に拒否する"); }
        finally { foreach (var table in duplicate) table.Rows.Dispose(); }
        var headerOnly = TabularFiles.Read([Path.Combine(root, "edge/13-header-only.csv"), Path.Combine(root, "folders/a/common.csv")]);
        try
        {
            var result = TabularTransform.Prepare(headerOnly, new(Merge: true, MergeByName: true), true).Tables[0];
            Check(result.RowCount == 2 && result.ColumnCount == 4 && result.Rows[1][1] == "001" && result.Rows[1][3] == "", "見出しだけのファイルも結合スキーマに参加し空セルを補完する");
        }
        finally { foreach (var table in headerOnly) table.Rows.Dispose(); }
    }
}
