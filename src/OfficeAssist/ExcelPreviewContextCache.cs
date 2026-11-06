namespace OfficeAssist;

// 設定中はExcelの操作をロックしている。列形式・桁数だけの変更で同じ出力属性を読み直さない。
internal sealed class ExcelPreviewContextCache(object window)
{
    sealed record Key(IReadOnlyList<TabularData> Tables, int File, ImportDestination Destination, string? Book, string? Sheet, string? Cell, bool Header, bool Merged);
    Key? current;
    ExcelPreviewEnvironment environment = new();
    readonly Dictionary<(int Row, int Column), ExcelPreviewStyle> styles = new();
    internal ExcelPreviewSnapshot Capture(IReadOnlyList<TabularData> tables, TabularImportOptions options, int file, IReadOnlyList<(int Row, int Column)> positions)
    {
        var key = new Key(tables, file, options.Destination, options.BookId, options.SheetName, options.CellAddress, options.Header, options.Merged);
        if (key != current) { styles.Clear(); current = key; }
        var missing = positions.Distinct().Where(p => !styles.ContainsKey(p)).ToArray();
        if (missing.Length > 0)
        {
            var captured = ExcelPreviewContext.Capture(window, tables, options, file, missing); environment = captured.Environment;
            if (styles.Count + missing.Length > 4096) styles.Clear();
            for (int i = 0; i < missing.Length; i++) styles[missing[i]] = captured.Styles[i];
        }
        // 1要求は128セル以下。上限を超えて古い属性を破棄した場合だけ、その要求分を取り直す。
        if (positions.Any(p => !styles.ContainsKey(p)))
        {
            var captured = ExcelPreviewContext.Capture(window, tables, options, file, positions); environment = captured.Environment; styles.Clear();
            for (int i = 0; i < positions.Count; i++) styles[positions[i]] = captured.Styles[i];
        }
        return new(environment, positions.Select(p => styles[p]).ToArray());
    }
}
