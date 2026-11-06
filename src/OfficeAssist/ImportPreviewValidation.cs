namespace OfficeAssist;

internal sealed record ImportCellError(int File, int Row, int Column, string Origin, string Raw, string Reason);
internal sealed record ImportPreviewValidation(long Cells, long ErrorCount, IReadOnlyList<ImportCellError> Examples)
{
    internal static ImportPreviewValidation Check(IReadOnlyList<TabularData> tables, TabularImportOptions options, CancellationToken token)
    {
        var errors = new List<ImportCellError>(); long cells = 0, errorCount = 0;
        if (options.Types.Length != tables.Count) throw new InvalidDataException("列設定の数が一致しません。");
        for (int f = 0; f < tables.Count; f++)
        {
            token.ThrowIfCancellationRequested();
            var table = tables[f]; if (options.Types[f].Length != table.ColumnCount) throw new InvalidDataException("列設定の数が一致しません。");
            for (int c = 0; c < table.ColumnCount; c++) { if (!Enum.IsDefined(options.Types[f][c])) throw new InvalidDataException("列形式が不正です。"); _ = ExcelValueFormat.Code(options.Types[f][c], options.Decimals(f, c)); }
            // 文字列2方式と標準は任意の保持済み文字列を受理する。結果が常に成功する列を再読込みしない。
            var checkedColumns = Enumerable.Range(0, table.ColumnCount).Where(c => options.Types[f][c] is not (ColumnType.TextFormat or ColumnType.TextQuote or ColumnType.General)).ToArray();
            cells += (long)table.RowCount * table.ColumnCount;
            if (checkedColumns.Length == 0) continue;
            for (int r = options.Header ? 1 : 0; r < table.RowCount; r++)
            {
                if ((r & 127) == 0) token.ThrowIfCancellationRequested(); var row = table.Rows[r];
                foreach (int c in checkedColumns)
                {
                    string raw = c < row.Length ? row[c] : "";
                    if (!ColumnConversion.TryConvert(raw, options.Types[f][c], false, out _, materialize: false))
                    {
                        errorCount++;
                        if (errors.Count < 200) errors.Add(new(f, r, c, table.Origin?.Invoke(r) ?? $"{table.Name}: 行{r + 1}", raw.Length <= 256 ? raw : raw[..256] + "…", $"{ColumnConversion.Labels[(int)options.Types[f][c]]}へ変換できない値があります。"));
                    }
                }
            }
        }
        return new(cells, errorCount, errors);
    }
}
