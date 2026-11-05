namespace OfficeAssist;

internal enum PreviewMode { None, First10, All }
internal sealed record TableTransformOptions(int SkipFirst = 0, int SkipLast = 0, bool DropEmpty = false, bool Merge = false, bool MergeByName = false);
internal sealed record PreparedTables(IReadOnlyList<TabularData> Tables, IReadOnlyList<string> Notices);

internal static class TabularTransform
{
    internal static string Key(string value) => value.TrimStart('\uFEFF').Trim();
    internal static string[] HeaderKeys(string[] header)
    {
        var keys = header.Select(Key).ToArray();
        if (keys.Any(string.IsNullOrEmpty) || keys.Distinct(StringComparer.Ordinal).Count() != keys.Length)
            throw new InvalidDataException("列名が空欄または重複しています。列順を選ぶか見出しを修正してください。");
        return keys;
    }
    internal static PreparedTables Prepare(IReadOnlyList<TabularData> sources, TableTransformOptions options, bool header, CancellationToken token = default)
    {
        if (options.SkipFirst < 0 || options.SkipLast < 0 || options.MergeByName && (!options.Merge || !header)) throw new InvalidDataException("行除外または結合基準の設定が不正です。");
        var tables = new List<TabularData>(); var notices = new List<string>();
        foreach (var source in sources)
        {
            var indexes = new List<int>();
            for (int r = Math.Min(source.RowCount, options.SkipFirst); r < source.RowCount - (long)options.SkipLast; r++)
            { token.ThrowIfCancellationRequested(); if (!options.DropEmpty || source.Rows[r].Any(v => v.Length != 0)) indexes.Add(r); }
            if (indexes.Count == 0) { notices.Add($"{source.Name}: 全行を除外しました。"); continue; }
            int width = indexes.Max(i => source.Rows[i].Length);
            var view = new TabularData(source.Name, new TabularRows(indexes.Count, width, i => source.Rows[indexes[i]]))
            { Origin = i => source.Origin?.Invoke(indexes[i]) ?? $"{source.Name}: 行{indexes[i] + 1}" };
            tables.Add(view);
            if (indexes.Count != source.RowCount) notices.Add($"{source.Name}: {source.RowCount - indexes.Count}行を除外しました。");
        }
        if (tables.Sum(t => t.RowCount - (header ? 1 : 0)) == 0) throw new InvalidDataException("取り込むデータ行がありません。行除外と見出し設定を確認してください。");
        if (!options.Merge) return new(tables.Where(t => t.RowCount > (header ? 1 : 0)).ToArray(), notices);
        var titles = new List<string>(); var keys = new List<string>(); var mappings = new List<int[]>();
        foreach (var table in tables)
        {
            string[] names = header ? table.Rows[0] : Enumerable.Range(1, table.ColumnCount).Select(c => "列" + c).ToArray();
            if (options.MergeByName)
            {
                string[] normalized = HeaderKeys(names);
                if (names.Length != table.ColumnCount) throw new InvalidDataException($"{table.Name}: 見出しより長いデータ行があります。");
                var map = new int[names.Length];
                for (int c = 0; c < names.Length; c++)
                { int index = keys.IndexOf(normalized[c]); if (index < 0) { index = keys.Count; keys.Add(normalized[c]); titles.Add(names[c]); } map[c] = index + 1; }
                mappings.Add(map);
            }
            else
            {
                for (int c = 0; c < table.ColumnCount; c++)
                { string title = c < names.Length && names[c].Length != 0 ? names[c] : "列" + (c + 1); if (c >= titles.Count) titles.Add(title); else if (header && titles[c] != title) notices.Add($"{table.Name}: 列{c + 1}の見出しが異なります（列順で結合）。"); }
                mappings.Add(Enumerable.Range(1, table.ColumnCount).ToArray());
            }
        }
        string filename = "ファイル名"; while (titles.Contains(filename)) filename += "（元ファイル）";
        titles.Insert(0, filename);
        var ends = new List<int>(); int count = header ? 1 : 0;
        foreach (var table in tables) { count += table.RowCount - (header ? 1 : 0); ends.Add(count); }
        if (titles.Count > 16384 || (long)count * titles.Count > TabularFiles.MaximumCells) throw new InvalidDataException("結合後の列数または200万セルの上限を超えます（ファイル名列を含む）。");
        (int File, int Row) Locate(int index) { int f = ends.FindIndex(end => index < end); return (f, index - (f == 0 ? (header ? 1 : 0) : ends[f - 1]) + (header ? 1 : 0)); }
        var merged = new TabularData("結合", new TabularRows(count, titles.Count, index =>
        {
            if (header && index == 0) return titles.ToArray();
            var (f, r) = Locate(index); var source = tables[f].Rows[r]; var values = Enumerable.Repeat("", titles.Count).ToArray(); values[0] = tables[f].Name;
            for (int c = 0; c < source.Length; c++) values[mappings[f][c]] = source[c]; return values;
        })) { Origin = index => { if (header && index == 0) return "結合: 見出し"; var (f, r) = Locate(index); return tables[f].Origin!(r); } };
        return new([merged], notices);
    }
    internal static int[] PreviewRows(TabularData table, bool header, PreviewMode mode)
    {
        if (mode == PreviewMode.None) return [];
        int start = header ? 1 : 0;
        if (mode == PreviewMode.All) return Enumerable.Range(start, table.RowCount - start).ToArray();
        var seen = new Dictionary<string, int>(); var rows = new List<int>();
        for (int r = start; r < table.RowCount; r++)
        {
            string origin = table.Origin?.Invoke(r) ?? table.Name;
            string file = origin.Contains(": 行", StringComparison.Ordinal) ? origin[..origin.LastIndexOf(": 行", StringComparison.Ordinal)] : origin;
            seen.TryGetValue(file, out int count); if (count < 10) rows.Add(r); seen[file] = count + 1;
        }
        return rows.ToArray();
    }
}
