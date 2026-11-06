using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace OfficeAssist;

internal static class ExcelImportEngine
{
    internal sealed record Write(TabularData Table, ColumnType[] Types, int[] Decimals, object Sheet, int Row, int Column, bool Header);
    internal static object ResolveBook(ComScope com, dynamic app, object? window, string? id)
    {
        if (id is null) { dynamic source = window ?? throw new InvalidOperationException("対象ブックを指定してください。"); dynamic sheet = com.Own(source.ActiveSheet); return com.Own(sheet.Parent); }
        dynamic books = com.Own(app.Workbooks);
        for (int i = 1; i <= (int)books.Count; i++) { dynamic book = com.Own(books.Item(i)); if ((string)book.FullName == id) return book; }
        throw new InvalidOperationException("指定ブックが閉じられたか変更されました。");
    }
    internal static object ResolveSheet(ComScope com, dynamic book, string name)
    {
        dynamic sheets = com.Own(book.Worksheets);
        for (int i = 1; i <= (int)sheets.Count; i++) { dynamic sheet = com.Own(sheets.Item(i)); if (string.Equals((string)sheet.Name, name, StringComparison.OrdinalIgnoreCase)) return sheet; }
        throw new InvalidOperationException("指定シートが見つかりません。");
    }
    internal static string NormalizeAddress(string address)
    {
        string normalized = address.Replace("$", "").ToUpperInvariant();
        if (!Regex.IsMatch(normalized, @"^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?(,[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?)*$", RegexOptions.CultureInvariant))
            throw new InvalidDataException("セルはB3またはB3:D7のようなA1形式で指定してください。");
        return normalized;
    }
    internal static void Validate(IReadOnlyList<TabularData> tables, TabularImportOptions options)
    {
        if (tables.Count == 0 || tables.Count != options.Types.Length || !Enum.IsDefined(options.Destination) || options.Merged && options.Destination == ImportDestination.AppendSameSheet)
            throw new InvalidDataException("取り込み設定が不正です。追記と結合は同時に指定できません。");
        long cells = 0;
        for (int t = 0; t < tables.Count; t++)
        {
            if (tables[t].RowCount == 0) throw new InvalidDataException("取り込む行がありません。");
            cells += (long)tables[t].RowCount * tables[t].ColumnCount;
            for (int c = 0; c < tables[t].ColumnCount; c++) options.Decimals(t, c);
            for (int r = 0; r < tables[t].RowCount; r += ChunkRows(tables[t]))
            { options.Progress?.Invoke($"{tables[t].Name}: {r}/{tables[t].RowCount}行を検証中"); options.Cancellation.ThrowIfCancellationRequested(); tables[t].ExcelValues(options.Types[t], options.Header, r, Math.Min(ChunkRows(tables[t]), tables[t].RowCount - r)); }
        }
        if (cells > TabularFiles.MaximumCells) throw new InvalidDataException("出力は合計200万セルまでです。");
    }
    static int ChunkRows(TabularData table) => Math.Max(1, 20000 / Math.Max(1, table.ColumnCount));
    internal static List<Write> Selection(ComScope com, dynamic window, IReadOnlyList<TabularData> tables, TabularImportOptions options)
    {
        dynamic app = com.Own(window.Application);
        dynamic sheet;
        dynamic selected;
        if (options.BookId is not null || options.SheetName is not null || options.CellAddress is not null)
        {
            dynamic book = ResolveBook(com, app, window, options.BookId);
            sheet = options.SheetName is null ? com.Own(book.ActiveSheet) : ResolveSheet(com, book, options.SheetName);
            selected = com.Own(sheet.Range[NormalizeAddress(options.CellAddress ?? "A1")]);
        }
        else { sheet = com.Own(window.ActiveSheet); selected = com.Own(window.RangeSelection); }
        dynamic areas = com.Own(selected.Areas); var writes = new List<Write>(); int next = 0;
        if (options.Merged && (int)areas.Count != 1) throw new InvalidDataException("結合時は指定セルを1つの連続範囲にしてください。");
        for (int t = 0; t < tables.Count; t++)
        {
            dynamic area = com.Own(areas.Item(Math.Min(t + 1, (int)areas.Count)));
            int row = t < (int)areas.Count ? (int)area.Row : next;
            writes.Add(new(tables[t], options.Types[t], Enumerable.Range(0, tables[t].ColumnCount).Select(c => options.Decimals(t, c)).ToArray(), sheet, row, (int)area.Column, options.Header));
            next = row + tables[t].RowCount + 1;
        }
        return writes;
    }
    internal static List<Write> Append(ComScope com, dynamic app, dynamic book, IReadOnlyList<TabularData> tables, TabularImportOptions options, List<object> created)
    {
        dynamic sheets = com.Own(book.Worksheets); var writes = new List<Write>();
        var nextRows = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var headers = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        for (int t = 0; t < tables.Count; t++)
        {
            var source = tables[t]; string name = Path.GetFileNameWithoutExtension(source.Name);
            ExcelAppendSchema.ValidateName(name);
            dynamic sheet;
            try { sheet = ResolveSheet(com, book, name); }
            catch (InvalidOperationException)
            {
                if ((bool)book.ProtectStructure) throw new InvalidOperationException("ブックの構成が保護されているためシートを作れません。");
                sheet = com.Own(sheets.Add(After: com.Own(sheets.Item((int)sheets.Count)))); created.Add(sheet); sheet.Name = name;
            }
            if (!nextRows.TryGetValue(name, out int next))
            {
                (int NextRow, string[] Headers) state = ExcelAppendSchema.Read(com, sheet, options.Header);
                next = state.NextRow; if (next > 1) headers[name] = state.Headers;
            }
            var data = source; var types = options.Types[t]; var decimals = Enumerable.Range(0, source.ColumnCount).Select(c => options.Decimals(t, c)).ToArray(); bool outputHeader = options.Header;
            if (next > 1)
            {
                int[] map = ExcelAppendSchema.Map(source, headers[name], options.Header);
                int skip = options.Header ? 1 : 0;
                data = new(source.Name, new TabularRows(source.RowCount - skip, map.Length, r => { var row = source.Rows[r + skip]; return map.Select(c => c >= 0 && c < row.Length ? row[c] : "").ToArray(); })) { Origin = r => source.Origin?.Invoke(r + skip) ?? $"{source.Name}: 行{r + skip + 1}" };
                types = map.Select(c => c >= 0 ? options.Types[t][c] : ColumnType.TextFormat).ToArray();
                int fileIndex = t; decimals = map.Select(c => c >= 0 ? options.Decimals(fileIndex, c) : 0).ToArray(); outputHeader = false;
            }
            else headers[name] = options.Header ? source.Rows[0] : new string[source.ColumnCount];
            writes.Add(new(data, types, decimals, sheet, next, 1, outputHeader)); nextRows[name] = next + data.RowCount;
        }
        return writes;
    }
    internal static void Execute(ComScope com, dynamic app, IReadOnlyList<Write> writes, TabularImportOptions options)
    {
        if (writes.Sum(w => (long)w.Table.RowCount * w.Table.ColumnCount) > TabularFiles.MaximumCells) throw new InvalidDataException("列対応後の出力が200万セルを超えます。");
        var targets = new List<object>(); var saved = new List<(object Range, object Format)>(); var written = new List<object>();
        foreach (var write in writes)
        {
            dynamic sheet = write.Sheet, book = com.Own(sheet.Parent), rows = com.Own(sheet.Rows), columns = com.Own(sheet.Columns);
            if ((bool)book.ReadOnly || (bool)sheet.ProtectContents || (bool)sheet.ProtectDrawingObjects) throw new InvalidOperationException("読み取り専用ブックや保護シートには展開できません。");
            if ((long)write.Row + write.Table.RowCount - 1 > (int)rows.Count || (long)write.Column + write.Table.ColumnCount - 1 > (int)columns.Count) throw new InvalidOperationException("展開先がExcelの行数・列数を超えます。");
            dynamic cells = com.Own(sheet.Cells), first = com.Own(cells.Item(write.Row, write.Column)), range = com.Own(first.Resize[write.Table.RowCount, write.Table.ColumnCount]);
            dynamic functions = com.Own(app.WorksheetFunction);
            if (range.MergeCells is not bool merged || merged || (double)functions.CountA(range) != 0 || range.HasArray is not bool array || array)
                throw new InvalidOperationException($"展開先 {ExcelAddress.Display(range)} に値・数式・結合セルがあります。空のセルを指定してください。");
            if (HasSpill(com, app, range)) throw new InvalidOperationException("展開先はスピル範囲と重なっています。");
            dynamic tables = com.Own(sheet.ListObjects), pivots = com.Own(sheet.PivotTables());
            for (int i = 1; i <= (int)tables.Count; i++) { dynamic table = com.Own(tables.Item(i)); if (app.Intersect(range, com.Own(table.Range)) is object overlap) { com.Own(overlap); throw new InvalidOperationException("展開先はExcelテーブルと重なっています。"); } }
            for (int i = 1; i <= (int)pivots.Count; i++) { dynamic pivot = com.Own(pivots.Item(i)); if (app.Intersect(range, com.Own(pivot.TableRange2)) is object overlap) { com.Own(overlap); throw new InvalidOperationException("展開先はピボットテーブルと重なっています。"); } }
            for (int i = 0; i < targets.Count; i++)
            { dynamic previousSheet = writes[i].Sheet; if ((string)previousSheet.Name == (string)sheet.Name && app.Intersect(targets[i], range) is object overlap) { com.Own(overlap); throw new InvalidOperationException("複数ファイルの展開先が重なっています。"); } }
            targets.Add(range);
        }
        bool events = (bool)app.EnableEvents, updating = (bool)app.ScreenUpdating;
        try
        {
            app.EnableEvents = false; app.ScreenUpdating = false;
            for (int i = 0; i < writes.Count; i++)
            {
                var write = writes[i]; dynamic target = targets[i], targetColumns = com.Own(target.Columns);
                options.Cancellation.ThrowIfCancellationRequested(); written.Add(target);
                for (int c = 0; c < write.Table.ColumnCount; c++)
                {
                    dynamic column = com.Own(targetColumns.Item(c + 1)); SaveFormat(com, column, saved);
                    ExcelValueFormat.Apply(app, column, ExcelValueFormat.Code(write.Types[c], write.Decimals[c]));
                }
                int step = ChunkRows(write.Table);
                for (int r = 0; r < write.Table.RowCount; r += step)
                {
                    options.Progress?.Invoke($"{write.Table.Name}: {r}/{write.Table.RowCount}行を書込み中"); options.Cancellation.ThrowIfCancellationRequested();
                    int count = Math.Min(step, write.Table.RowCount - r);
                    var values = write.Table.ExcelValues(write.Types, write.Header, r, count);
                    dynamic dateBook = com.Own(((dynamic)write.Sheet).Parent);
                    AdjustDateSystem(values, write.Types, (bool)dateBook.Date1904);
                    dynamic cells = com.Own(target.Cells), first = com.Own(cells.Item(r + 1, 1)), chunk = com.Own(first.Resize[count, write.Table.ColumnCount]);
                    for (int c = 0; c < write.Table.ColumnCount; c++)
                    {
                        if (write.Types[c] is ColumnType.TextFormat or ColumnType.TextQuote) continue;
                        for (int row = 0; row < count; row++)
                        {
                            if (values[row, c] is not string) continue; int start = row;
                            while (row + 1 < count && values[row + 1, c] is string) row++;
                            dynamic textRange = com.Own(cells.Item(r + start + 1, c + 1).Resize[row - start + 1, 1]); ExcelValueFormat.Apply(app, textRange, "@");
                        }
                    }
                    chunk.Value2 = values;
                }
            }
        }
        catch (Exception ex)
        {
            bool restored = true;
            foreach (dynamic target in written) try { target.ClearContents(); } catch { restored = false; }
            foreach (var format in saved) try { dynamic range = format.Range; range.NumberFormat = format.Format; } catch { restored = false; }
            throw new InvalidOperationException(restored ? "取り込みを中止しました。今回の書込みを取り消しました。" : "書込みの取り消しに失敗しました。対象シートを確認してください。", ex);
        }
        finally { app.EnableEvents = events; app.ScreenUpdating = updating; }
        options.Completed?.Invoke(writes);
    }
    static void SaveFormat(ComScope com, dynamic range, List<(object Range, object Format)> saved)
    {
        object? format = range.NumberFormat;
        if (format is string) { saved.Add((range, format)); return; }
        dynamic rows = com.Own(range.Rows), cells = com.Own(range.Cells); int count = (int)rows.Count;
        if (count == 1) { saved.Add((range, "General")); return; }
        int half = count / 2;
        SaveFormat(com, com.Own(cells.Item(1, 1).Resize[half, 1]), saved);
        SaveFormat(com, com.Own(cells.Item(half + 1, 1).Resize[count - half, 1]), saved);
    }
    static bool HasSpill(ComScope com, dynamic app, dynamic range)
    {
        try { object? value = range.HasSpill; if (value is bool spill) return spill; }
        catch (COMException ex) when ((uint)ex.HResult is 0x800A03EC or 0x80020003 or 0x80020006) { }
        object capability = app.Evaluate("IFERROR(ROWS(SEQUENCE(1)),0)");
        double capabilityValue = capability is double or int or short or decimal ? Convert.ToDouble(capability, System.Globalization.CultureInfo.InvariantCulture) : double.NaN;
        if (capabilityValue is not (0 or 1))
        {
            // 小数点がカンマのユーザー設定ではEvaluateの引数区切りも変わる場合がある。
            capability = app.Evaluate("IFERROR(ROWS(SEQUENCE(1));0)");
            capabilityValue = capability is double or int or short or decimal ? Convert.ToDouble(capability, System.Globalization.CultureInfo.InvariantCulture) : double.NaN;
        }
        if (capabilityValue == 0) return false;
        if (capabilityValue != 1) throw new InvalidOperationException($"Excelの動的配列への対応を確認できませんでした（{capability.GetType().Name}:{capabilityValue}）。");
        // HasSpillが利用できないExcelでは、数式アンカーの#参照を検査する。
        dynamic sheet = com.Own(range.Parent), used = com.Own(sheet.UsedRange);
        object formulaObject;
        try { formulaObject = used.SpecialCells(-4123); }
        catch (COMException ex) when ((uint)ex.HResult == 0x800A03EC) { return false; }
        dynamic formulas = com.Own(formulaObject), areas = com.Own(formulas.Areas);
        for (int a = 1; a <= (int)areas.Count; a++)
        {
            dynamic area = com.Own(areas.Item(a)), cells = com.Own(area.Cells);
            for (int i = 1; i <= (int)cells.Count; i++)
            {
                dynamic cell = com.Own(cells.Item(i));
                object isReference = sheet.Evaluate("ISREF(" + ExcelAddress.Display(cell) + "#)");
                if (isReference is not bool supported) throw new InvalidOperationException("数式のスピル範囲を確認できませんでした。");
                if (!supported) continue;
                dynamic spilled = com.Own(cell.SpillingToRange);
                if (app.Intersect(range, spilled) is object overlap) { com.Own(overlap); return true; }
            }
        }
        return false;
    }
    internal static void AdjustDateSystem(object?[,] values, ColumnType[] types, bool date1904)
    {
        if (!date1904) return;
        for (int r = 0; r < values.GetLength(0); r++) for (int c = 0; c < types.Length; c++)
            values[r, c] = ExcelValueFormat.AdjustSerial(values[r, c], types[c], date1904);
    }
}
