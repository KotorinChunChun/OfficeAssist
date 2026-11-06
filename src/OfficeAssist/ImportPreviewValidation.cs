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
            var table = tables[f]; if (options.Types[f].Length != table.ColumnCount) throw new InvalidDataException("列設定の数が一致しません。");
            for (int c = 0; c < table.ColumnCount; c++) _ = ExcelValueFormat.Code(options.Types[f][c], options.Decimals(f, c));
            for (int r = 0; r < table.RowCount; r++)
            {
                if ((r & 127) == 0) token.ThrowIfCancellationRequested(); var row = table.Rows[r];
                for (int c = 0; c < table.ColumnCount; c++)
                {
                    string raw = c < row.Length ? row[c] : ""; cells++;
                    try { _ = ColumnConversion.Convert(raw, options.Types[f][c], options.Header && r == 0); }
                    catch (InvalidDataException ex)
                    {
                        errorCount++;
                        if (errors.Count < 200) errors.Add(new(f, r, c, table.Origin?.Invoke(r) ?? $"{table.Name}: 行{r + 1}", raw.Length <= 256 ? raw : raw[..256] + "…", ex.Message));
                    }
                }
            }
        }
        return new(cells, errorCount, errors);
    }
}
