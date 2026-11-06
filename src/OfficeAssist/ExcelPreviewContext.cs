namespace OfficeAssist;

internal sealed record ExcelPreviewSnapshot(ExcelPreviewEnvironment Environment, IReadOnlyList<ExcelPreviewStyle> Styles);

internal static class ExcelPreviewContext
{
    internal static ExcelPreviewSnapshot Capture(object sourceWindow, IReadOnlyList<TabularData> tables, TabularImportOptions options, int file, IReadOnlyList<(int Row, int Column)> positions)
    {
        using var com = new ComScope(); dynamic window = sourceWindow, app = com.Own(window.Application);
        bool system = (bool)app.UseSystemSeparators; string decimalSeparator = (string)app.DecimalSeparator, thousands = (string)app.ThousandsSeparator;
        if (options.Destination == ImportDestination.NewBook) return new(new(false, system, decimalSeparator, thousands), positions.Select(_ => new ExcelPreviewStyle()).ToArray());
        dynamic book = ExcelImportEngine.ResolveBook(com, app, window, options.BookId);
        dynamic normal = com.Own(com.Own(book.Styles).Item("Normal")), normalFont = com.Own(normal.Font);
        var environment = new ExcelPreviewEnvironment((bool)book.Date1904, system, decimalSeparator, thousands, (string)normalFont.Name, (double)normalFont.Size);
        object? sheetObject = null; int firstRow = 1, firstColumn = 1, skip = 0; int[]? columnMap = null;
        if (options.Destination == ImportDestination.Selection)
        {
            List<ExcelImportEngine.Write> writes = ExcelImportEngine.Selection(com, window, tables, options);
            var write = writes[file]; sheetObject = write.Sheet; firstRow = write.Row; firstColumn = write.Column;
        }
        else if (options.Destination == ImportDestination.AppendSameSheet)
        {
            string name = Path.GetFileNameWithoutExtension(tables[file].Name); ExcelAppendSchema.ValidateName(name);
            try { sheetObject = ExcelImportEngine.ResolveSheet(com, book, name); } catch (InvalidOperationException) { }
            (int NextRow, string[] Headers) state = sheetObject is null ? (1, []) : ExcelAppendSchema.Read(com, (dynamic)sheetObject, options.Header);
            for (int f = 0; f <= file; f++)
            {
                if (!Path.GetFileNameWithoutExtension(tables[f].Name).Equals(name, StringComparison.OrdinalIgnoreCase)) continue;
                bool existing = state.NextRow > 1;
                int[] mapping = existing ? ExcelAppendSchema.Map(tables[f], state.Headers, options.Header) : Enumerable.Range(0, tables[f].ColumnCount).ToArray();
                if (f == file) { firstRow = state.NextRow; skip = existing && options.Header ? 1 : 0; columnMap = mapping; break; }
                state = (state.NextRow + tables[f].RowCount - (existing && options.Header ? 1 : 0), existing ? state.Headers : options.Header ? tables[f].Rows[0] : new string[tables[f].ColumnCount]);
            }
        }
        if ((bool)book.ReadOnly) throw new InvalidOperationException("指定ブックは読み取り専用です。");
        if (sheetObject is null)
        {
            dynamic sheets = com.Own(book.Worksheets), sample = com.Own(sheets.Item(1));
            if ((bool)book.ProtectStructure) throw new InvalidOperationException("指定ブックの構成が保護されています。");
            double width = (double)sample.StandardWidth;
            return new(environment, positions.Select(_ => new ExcelPreviewStyle(WidthCharacters: width, FontName: environment.NormalFont, FontSize: environment.NormalFontSize)).ToArray());
        }
        dynamic sheet = sheetObject, sheetRows = com.Own(sheet.Rows), sheetColumns = com.Own(sheet.Columns);
        if ((bool)sheet.ProtectContents || (bool)sheet.ProtectDrawingObjects) throw new InvalidOperationException("指定シートは保護されています。");
        int rowLimit = (int)sheetRows.Count, columnLimit = (int)sheetColumns.Count;
        var addresses = positions.Select(p => (Row: firstRow + p.Row - skip,
            Column: firstColumn + (columnMap is null ? p.Column : Array.IndexOf(columnMap, p.Column)))).ToArray();
        foreach (var p in addresses) if (p.Row < 1 || p.Column < 1 || p.Row > rowLimit || p.Column > columnLimit) throw new InvalidOperationException("プレビューの出力先がExcelの行数・列数を超えます。");
        if (addresses.Length == 0) return new(environment, []);
        dynamic cells = com.Own(sheet.Cells);
        int minRow = addresses.Min(p => p.Row), maxRow = addresses.Max(p => p.Row), minColumn = addresses.Min(p => p.Column), maxColumn = addresses.Max(p => p.Column);
        // 通常の表は属性が共通。範囲で一度読み、混在する場合だけ個別読取りへ戻す。
        ExcelPreviewStyle? common = null;
        if (maxRow - minRow <= 256 && (long)(maxRow - minRow + 1) * (maxColumn - minColumn + 1) <= 8192)
        {
            dynamic first = com.Own(cells.Item(minRow, minColumn)), last = com.Own(cells.Item(maxRow, maxColumn)), range = com.Own(sheet.Range[first, last]);
            common = CommonStyle(com, range);
        }
        var widths = new Dictionary<int, double>(); var styles = new List<ExcelPreviewStyle>();
        foreach (var position in addresses)
        {
            int row = position.Row, column = position.Column;
            if (common is not null)
            {
                if (!widths.TryGetValue(column, out double width)) { dynamic source = com.Own(cells.Item(row, column)); widths[column] = width = (double)source.Width; }
                styles.Add(common with { WidthPoints = width }); continue;
            }
            dynamic cell = com.Own(cells.Item(row, column)), font = com.Own(cell.Font);
            styles.Add(new((double)cell.Width, FontName: (string)font.Name, FontSize: (double)font.Size, Bold: (bool)font.Bold, Italic: (bool)font.Italic,
                Wrap: (bool)cell.WrapText, Shrink: (bool)cell.ShrinkToFit, RowHeight: (double)cell.RowHeight));
        }
        return new(environment, styles);
    }
    static ExcelPreviewStyle? CommonStyle(ComScope com, dynamic range)
    {
        dynamic font = com.Own(range.Font);
        object? name = font.Name, size = font.Size, bold = font.Bold, italic = font.Italic, wrap = range.WrapText, shrink = range.ShrinkToFit, height = range.RowHeight;
        return name is string n && size is double s && bold is bool b && italic is bool i && wrap is bool w && shrink is bool sh && height is double h
            ? new(FontName: n, FontSize: s, Bold: b, Italic: i, Wrap: w, Shrink: sh, RowHeight: h) : null;
    }
}
