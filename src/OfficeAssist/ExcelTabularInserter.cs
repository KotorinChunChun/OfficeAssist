using System.Runtime.InteropServices;

namespace OfficeAssist;

internal static class ExcelTabularInserter
{
    sealed record ExcelOwner(nint Handle) : IWin32Window;
    internal static ExcelPictureInserter.Result? Insert(nint handle, Point point, IReadOnlyList<TabularData> tables)
    {
        if (!Native.IsWindow(handle) || Native.ClassName(handle) != "EXCEL7" || !Native.IsWindowEnabled(Native.GetAncestor(handle, 2)))
            throw new InvalidOperationException("対象のExcelが閉じられたか、ダイアログを表示しています。");
        using var com = new ComScope();
        var iid = new Guid("00020400-0000-0000-C000-000000000046");
        Marshal.ThrowExceptionForHR(Native.AccessibleObjectFromWindow(handle, 0xFFFFFFF0, ref iid, out object nativeWindow));
        dynamic window = com.Own(nativeWindow);
        object? hitObject = window.RangeFromPoint(point.X, point.Y);
        if (hitObject is null) throw new InvalidOperationException("ワークシートのセルへドロップしてください。");
        dynamic hit = com.Own(hitObject);
        try
        {
            dynamic rows = com.Own(hit.Rows), columns = com.Own(hit.Columns);
            if ((int)rows.Count != 1 || (int)columns.Count != 1) throw new InvalidOperationException("見出しではなくセルへドロップしてください。");
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        { throw new InvalidOperationException("ワークシートのセルへドロップしてください。"); }
        dynamic originalSheet = com.Own(window.ActiveSheet);
        dynamic originalBook = com.Own(originalSheet.Parent);
        dynamic originalSelection = com.Own(window.RangeSelection);
        string address = (string)originalSelection.Address;
        string sheetName = (string)originalSheet.Name;
        using var dialog = new TabularImportForm(tables) { StartPosition = FormStartPosition.CenterParent };
        var owner = new ExcelOwner(Native.GetAncestor(handle, 2));
        if (dialog.ShowDialog(owner) != DialogResult.OK || dialog.Options is null) return null;
        if (!Native.IsWindow(handle) || !Native.IsWindowEnabled(Native.GetAncestor(handle, 2))) throw new InvalidOperationException("対象Excelを確認できません。もう一度ドロップしてください。");
        dynamic currentSheet = com.Own(window.ActiveSheet);
        dynamic currentBook = com.Own(currentSheet.Parent);
        dynamic currentSelection = com.Own(window.RangeSelection);
        if ((string)currentSheet.Name != sheetName || (string)currentBook.FullName != (string)originalBook.FullName || (string)currentSelection.Address != address)
            throw new InvalidOperationException("設定中にブック・シート・選択範囲が変わりました。もう一度ドロップしてください。");
        dynamic app = com.Own(window.Application);
        return Import(com, app, window, tables, dialog.Options);
    }

    internal static ExcelPictureInserter.Result Import(ComScope com, dynamic app, object? sourceWindow, IReadOnlyList<TabularData> tables, TabularImportOptions options)
    {
        Validate(tables, options);
        if (options.Destination == ImportDestination.Selection)
            return InsertIntoWindow(com, sourceWindow ?? throw new InvalidOperationException("対象ウィンドウが必要です。"), tables, options);
        object? createdBook = null;
        var createdSheets = new List<object>();
        bool events = (bool)app.EnableEvents;
        dynamic book;
        try
        {
            app.EnableEvents = false;
            if (options.Destination == ImportDestination.NewBook)
            {
                dynamic books = com.Own(app.Workbooks);
                createdBook = com.Own(books.Add(-4167)); book = createdBook;
                book.Date1904 = false;
            }
            else
            {
                dynamic window = sourceWindow ?? throw new InvalidOperationException("対象ウィンドウが必要です。");
                dynamic sourceSheet = com.Own(window.ActiveSheet); book = com.Own(sourceSheet.Parent);
                if ((bool)book.ReadOnly || (bool)book.ProtectStructure) throw new InvalidOperationException("読み取り専用または構成保護中のブックにはシートを追加できません。");
            }
            dynamic sheets = com.Own(book.Worksheets);
            for (int t = 0; t < tables.Count; t++)
            {
                dynamic sheet;
                if (createdBook is not null && t == 0) sheet = com.Own(sheets.Item(1));
                else { dynamic last = com.Own(sheets.Item((int)sheets.Count)); sheet = com.Own(sheets.Add(After: last)); }
                createdSheets.Add(sheet);
                string stem = string.Concat(Path.GetFileNameWithoutExtension(tables[t].Name).Where(c => !"[]:*?/\\'".Contains(c))).Trim();
                if (stem.Length == 0) stem = "取込";
                stem = stem[..Math.Min(stem.Length, 24)];
                for (int suffix = 0; suffix < 10000; suffix++)
                {
                    string candidate = suffix == 0 ? stem : stem + "_" + suffix;
                    bool exists = false;
                    for (int i = 1; i <= (int)sheets.Count; i++) { dynamic other = com.Own(sheets.Item(i)); if ((string)other.Name == candidate) { exists = true; break; } }
                    if (!exists) { sheet.Name = candidate; break; }
                }
                sheet.Activate(); dynamic cell = com.Own(sheet.Range["A1"]); cell.Select();
                dynamic window = com.Own(app.ActiveWindow);
                InsertIntoWindow(com, window, new[] { tables[t] }, new TabularImportOptions(ImportDestination.Selection, [options.Types[t]], options.Header)
                { DecimalPlaces = [Enumerable.Range(0, tables[t].ColumnCount).Select(c => options.Decimals(t, c)).ToArray()] });
            }
            return new($"{tables.Count}ファイルを{(options.Destination == ImportDestination.NewBook ? "新しいブック" : "新しいシート")}に取り込みました（{book.Name}）。");
        }
        catch (Exception ex)
        {
            bool alerts = (bool)app.DisplayAlerts, restored = true;
            try
            {
                app.DisplayAlerts = false;
                if (createdBook is not null) { dynamic failedBook = createdBook; failedBook.Close(false); }
                else foreach (dynamic sheet in createdSheets.AsEnumerable().Reverse()) sheet.Delete();
            }
            catch (Exception) { restored = false; }
            finally { app.DisplayAlerts = alerts; }
            throw new InvalidOperationException(restored ? "取り込みに失敗しました。今回作成したブック・シートは取り消しました。" : "取り込みと作成物の取り消しに失敗しました。対象ブックを確認してください。", ex);
        }
        finally { app.EnableEvents = events; }
    }

    static void Validate(IReadOnlyList<TabularData> tables, TabularImportOptions options)
    {
        if (tables.Count == 0 || tables.Count != options.Types.Length || !Enum.IsDefined(options.Destination)) throw new InvalidDataException("取り込み設定が不正です。");
        for (int t = 0; t < tables.Count; t++)
        {
            tables[t].ExcelValues(options.Types[t], options.Header);
            for (int c = 0; c < tables[t].ColumnCount; c++) options.Decimals(t, c);
        }
    }

    internal static ExcelPictureInserter.Result InsertIntoWindow(ComScope com, dynamic window, IReadOnlyList<TabularData> tables, TabularImportOptions? settings = null)
    {
        if (tables.Count == 0) throw new InvalidOperationException("展開するデータがありません。");
        var options = settings ?? TabularImportOptions.Infer(tables);
        Validate(tables, options);
        var values = tables.Select((t, i) => t.ExcelValues(options.Types[i], options.Header)).ToArray();
        dynamic sheet = com.Own(window.ActiveSheet);
        if ((int)sheet.Type != -4167) throw new InvalidOperationException("ワークシートを選択してください。");
        dynamic book = com.Own(sheet.Parent);
        if ((bool)book.ReadOnly || (bool)sheet.ProtectContents || (bool)sheet.ProtectDrawingObjects)
            throw new InvalidOperationException("読み取り専用ブックや保護シートには展開できません。");
        dynamic app = com.Own(window.Application);
        dynamic selected = com.Own(window.RangeSelection);
        dynamic areas = com.Own(selected.Areas);
        var destinations = new List<object>();
        for (int i = 1; i <= (int)areas.Count; i++) destinations.Add(com.Own(areas.Item(i)));
        dynamic sheetCells = com.Own(sheet.Cells);
        dynamic functions = com.Own(app.WorksheetFunction);
        var targets = new List<object>();
        int nextRow = 0;
        for (int i = 0; i < tables.Count; i++)
        {
            dynamic area = destinations[Math.Min(i, destinations.Count - 1)];
            int row = i < destinations.Count ? (int)area.Row : nextRow;
            int column = (int)area.Column;
            var table = tables[i];
            if ((long)row + table.RowCount - 1 > 1048576 || (long)column + table.ColumnCount - 1 > 16384)
                throw new InvalidOperationException("展開先がExcelの行数・列数を超えます。");
            dynamic first = com.Own(sheetCells.Item(row, column));
            dynamic target = com.Own(first.Resize[table.RowCount, table.ColumnCount]);
            nextRow = row + table.RowCount + 1;
            object? merged = target.MergeCells;
            if (merged is not bool hasMerged || hasMerged || (double)functions.CountA(target) != 0)
                throw new InvalidOperationException($"展開先 {target.Address} に値・数式・結合セルがあります。空のセルを選択してください。");
            foreach (dynamic previous in targets)
            {
                object? overlap = app.Intersect(previous, target);
                if (overlap is not null) { com.Own(overlap); throw new InvalidOperationException("複数ファイルの展開先が重なっています。選択範囲を離してください。"); }
            }
            targets.Add(target);
        }
        bool eventsEnabled = (bool)app.EnableEvents;
        var written = new List<object>();
        var formats = new List<(object Range, object Format)>();
        try
        {
            app.EnableEvents = false;
            for (int i = 0; i < targets.Count; i++)
            {
                dynamic target = targets[i];
                if ((double)functions.CountA(target) != 0) throw new InvalidOperationException("処理中に展開先が変更されました。");
                written.Add(target);
                dynamic targetColumns = com.Own(target.Columns);
                for (int c = 0; c < tables[i].ColumnCount; c++)
                {
                    dynamic column = com.Own(targetColumns.Item(c + 1));
                    object? previous = column.NumberFormat;
                    if (previous is string) formats.Add((column, previous));
                    else
                    {
                        dynamic columnCells = com.Own(column.Cells);
                        for (int r = 1; r <= tables[i].RowCount; r++) { dynamic cell = com.Own(columnCells.Item(r, 1)); formats.Add((cell, (object)cell.NumberFormat)); }
                    }
                    var type = options.Types[i][c];
                    string format = ColumnConversion.Format(type, options.Decimals(i, c));
                    try
                    {
                        if (type is ColumnType.General or ColumnType.TextQuote) column.NumberFormatLocal = (string)app.International[26]; // xlGeneralFormatName
                        else column.NumberFormat = format;
                    }
                    catch (COMException ex) { throw new InvalidOperationException($"列{c + 1}の書式（{format}）を設定できません。", ex); }
                    dynamic cells = com.Own(column.Cells);
                    for (int r = 0; r < tables[i].RowCount; r++)
                    {
                        if (values[i][r, c] is string && type != ColumnType.TextQuote && type != ColumnType.TextFormat)
                        { dynamic textCell = com.Own(cells.Item(r + 1, 1)); textCell.NumberFormat = "@"; }
                    }
                }
                target.Value2 = values[i];
            }
        }
        catch (Exception ex)
        {
            bool restored = true;
            foreach (dynamic target in written) { try { target.ClearContents(); } catch (Exception) { restored = false; } }
            foreach (var saved in formats) { try { dynamic range = saved.Range; range.NumberFormat = saved.Format; } catch (Exception) { restored = false; } }
            throw new InvalidOperationException(restored ? "CSV・TSVを展開できませんでした。今回の書込みを取り消しました。"
                : "CSV・TSVの展開と取り消しに失敗しました。対象シートを確認してください。", ex);
        }
        finally { app.EnableEvents = eventsEnabled; }
        var warnings = new List<string>();
        if (tables.Count > destinations.Count) warnings.Add("選択範囲がファイル数より少ないため、残りを最後の範囲の下へ１行空けて順に展開しました。");
        try
        {
            if (tables.Count < destinations.Count)
            {
                dynamic remaining = destinations[tables.Count];
                for (int i = tables.Count + 1; i < destinations.Count; i++) remaining = com.Own(app.Union(remaining, destinations[i]));
                remaining.Select();
            }
            else
            {
                dynamic last = targets[^1];
                dynamic lastRows = com.Own(last.Rows);
                int next = (int)last.Row + (int)lastRows.Count;
                if (next <= 1048576) { dynamic cell = com.Own(sheetCells.Item(next, (int)last.Column)); cell.Select(); }
            }
        }
        catch (Exception) { warnings.Add("展開は完了しましたが、次のセルの選択に失敗しました。展開先を選び直してください。"); }
        return new($"{tables.Count}ファイルを展開しました（{sheet.Name}）。", warnings.Count == 0 ? null : string.Join(Environment.NewLine, warnings));
    }
}
