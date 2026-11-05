using OfficeAssist;
using System.Text;

internal static partial class Program
{
    static void V3ModelTests(string output)
    {
        TabularData[] source = [new("first.csv", [["説明"], ["ID", "値"], ["001", "10"], ["", ""], ["002", "20"], ["集計"]]),
            new("second.tsv", [["説明"], ["値", "ID", "追加"], ["30", "003", "備考"], ["集計"]])];
        var prepared = TabularTransform.Prepare(source, new(1, 1, true, true, true), true);
        var merged = prepared.Tables.Single();
        Check(merged.RowCount == 4 && merged.ColumnCount == 4, "行除外・空欄除去・見出し除去後に結合する");
        Check(merged.Rows[0].SequenceEqual(new[] { "ファイル名", "ID", "値", "追加" }), "列名結合は後続の列を右へ追加する");
        Check(merged.Rows[3].SequenceEqual(new[] { "second.tsv", "003", "30", "備考" }), "列名順序違いを対応させる");
        Check(merged.Rows[1][3] == "" && merged.Origin!(3).Contains("行3"), "不足列の空セル補完と元行番号を保持する");
        var raw = TabularTransform.Prepare(source, new(1, 1), true).Tables;
        Check(raw[0].RowCount == 4, "空欄除外OFFでは空欄レコードを残す");
        Reject(() => TabularTransform.Prepare(source, new(100, 0), true), "全行除外は書込み前に拒否する");
        Reject(() => TabularTransform.Prepare([new("a.csv", [["ID", " ID "], ["1", "2"]])], new(Merge: true, MergeByName: true), true), "列名の正規化後の重複を拒否する");
        Reject(() => TabularTransform.Prepare(source, new(Merge: true, MergeByName: true), false), "見出しOFFの列名結合を拒否する");
        var many = new TabularData("many.csv", Enumerable.Range(0, 30).Select(i => new[] { i.ToString() }).ToArray());
        Check(TabularTransform.PreviewRows(many, false, PreviewMode.None).Length == 0 && TabularTransform.PreviewRows(many, false, PreviewMode.First10).Length == 10 && TabularTransform.PreviewRows(many, false, PreviewMode.All).Length == 30, "プレビューの3モードは参照行だけを切り替える");
        var multiple = TabularTransform.Prepare([many, many with { Name = "other.tsv" }], new(Merge: true), false).Tables[0];
        Check(TabularTransform.PreviewRows(multiple, false, PreviewMode.First10).Length == 20, "結合プレビューは各ファイル10行ずつ示す");
        string path = Path.Combine(output, "stream.csv");
        File.WriteAllText(path, "ID,値\r\n001,\"前半\r\n後半\"\r\n002,末尾\r\n", new UTF8Encoding(true));
        var disk = TabularFiles.Read([path]);
        try
        {
            Check(disk[0].Rows[1][1] == "前半\r\n後半" && disk[0].RowCount == 3, "ストリーム解析で引用符内改行と末尾改行を保持する");
            var filtered = TabularTransform.Prepare(disk, new(SkipLast: 1), true).Tables[0];
            Check(filtered.RowCount == 2 && filtered.Rows[1][0] == "001", "ディスク保持から行ビューを構築する");
        }
        finally { foreach (var table in disk) table.Rows.Dispose(); }
        Check(ColumnInference.Scan(new("mixed.csv", [["列"], ["12.34"], ["4.5"]]), true).Decimals[0] == 2, "全件走査の型推定で小数桁を保持する");
    }
}
