using System.Runtime.InteropServices;

namespace OfficeAssist;

internal static class ExcelTabularInserter
{
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
        var owner = new ExcelWindowOwner(Native.GetAncestor(handle, 2));
        using var modalLock = new ExcelModalLock(owner.Handle);
        using var dialog = new TabularImportForm(tables, options =>
        {
            if (!Native.IsWindow(handle)) throw new InvalidOperationException("対象Excelを確認できません。もう一度ドロップしてください。");
            dynamic currentSheet = com.Own(window.ActiveSheet), currentBook = com.Own(currentSheet.Parent), currentSelection = com.Own(window.RangeSelection);
            if ((string)currentSheet.Name != sheetName || (string)currentBook.FullName != (string)originalBook.FullName || (string)currentSelection.Address != address)
                throw new InvalidOperationException("設定中にブック・シート・選択範囲が変わりました。もう一度ドロップしてください。");
            dynamic app = com.Own(window.Application);
            return Import(com, app, window, tables, options);
        }) { StartPosition = FormStartPosition.CenterParent };
        return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.Result : null;
    }

    internal static ExcelPictureInserter.Result Import(ComScope com, dynamic app, object? sourceWindow, IReadOnlyList<TabularData> tables, TabularImportOptions options)
    {
        ExcelImportEngine.Validate(tables, options);
        if (options.Destination == ImportDestination.Selection)
            return InsertIntoWindow(com, sourceWindow ?? throw new InvalidOperationException("対象ウィンドウが必要です。"), tables, options);
        object? createdBook = null;
        var created = new List<object>();
        bool events = (bool)app.EnableEvents;
        dynamic book;
        try
        {
            app.EnableEvents = false;
            if (options.Destination == ImportDestination.NewBook)
            { dynamic books = com.Own(app.Workbooks); createdBook = com.Own(books.Add(-4167)); book = createdBook; book.Date1904 = false; }
            else book = ExcelImportEngine.ResolveBook(com, app, sourceWindow, options.BookId);
            if ((bool)book.ReadOnly) throw new InvalidOperationException("読み取り専用のブックには書き込めません。");
            List<ExcelImportEngine.Write> writes;
            if (options.Destination == ImportDestination.AppendSameSheet) writes = ExcelImportEngine.Append(com, app, book, tables, options, created);
            else
            {
                if ((bool)book.ProtectStructure) throw new InvalidOperationException("構成保護中のブックにはシートを追加できません。");
                dynamic sheets = com.Own(book.Worksheets); writes = new();
                for (int t = 0; t < tables.Count; t++)
                {
                    dynamic sheet;
                    if (createdBook is not null && t == 0) sheet = com.Own(sheets.Item(1));
                    else sheet = com.Own(sheets.Add(After: com.Own(sheets.Item((int)sheets.Count))));
                    created.Add(sheet);
                    string stem = string.Concat(Path.GetFileNameWithoutExtension(tables[t].Name).Where(c => !"[]:*?/\\'".Contains(c))).Trim();
                    if (stem.Length == 0) stem = "取込"; stem = stem[..Math.Min(stem.Length, 24)];
                    for (int suffix = 0; suffix < 10000; suffix++)
                    {
                        string candidate = suffix == 0 ? stem : stem + "_" + suffix; bool exists = false;
                        for (int i = 1; i <= (int)sheets.Count; i++) { dynamic other = com.Own(sheets.Item(i)); if ((string)other.Name == candidate) { exists = true; break; } }
                        if (!exists) { sheet.Name = candidate; break; }
                    }
                    int index = t;
                    writes.Add(new(tables[t], options.Types[t], Enumerable.Range(0, tables[t].ColumnCount).Select(c => options.Decimals(index, c)).ToArray(), sheet, 1, 1, options.Header));
                }
            }
            ExcelImportEngine.Execute(com, app, writes, options);
            dynamic last = writes[^1].Sheet; last.Activate();
            return new($"{tables.Count}ファイルを{(options.Destination == ImportDestination.AppendSameSheet ? "同名のシートに追記しました" : "取り込みました")}（{book.Name}）。");
        }
        catch (Exception ex)
        {
            bool alerts = (bool)app.DisplayAlerts, restored = true;
            try { app.DisplayAlerts = false; if (createdBook is not null) { dynamic failed = createdBook; failed.Close(false); } else foreach (dynamic sheet in created.AsEnumerable().Reverse()) sheet.Delete(); }
            catch { restored = false; }
            finally { app.DisplayAlerts = alerts; }
            throw new InvalidOperationException(restored ? "取り込みに失敗しました。今回作成したブック・シートは取り消しました。 " + ex.Message : "取り込みと取消しに失敗しました。対象ブックを確認してください。", ex);
        }
        finally { app.EnableEvents = events; }
    }

    internal static ExcelPictureInserter.Result InsertIntoWindow(ComScope com, dynamic window, IReadOnlyList<TabularData> tables, TabularImportOptions? settings = null)
    {
        var options = settings ?? TabularImportOptions.Infer(tables);
        ExcelImportEngine.Validate(tables, options);
        List<ExcelImportEngine.Write> writes = ExcelImportEngine.Selection(com, window, tables, options);
        dynamic app = com.Own(window.Application);
        dynamic initialSheet = com.Own(window.ActiveSheet), initialBook = com.Own(initialSheet.Parent), initialSelection = com.Own(window.RangeSelection);
        bool originalSelection = (options.BookId is null || options.BookId == (string)initialBook.FullName)
            && (options.SheetName is null || options.SheetName == (string)initialSheet.Name)
            && (options.CellAddress is null || ExcelImportEngine.NormalizeAddress(options.CellAddress) == ExcelAddress.Display(initialSelection));
        ExcelImportEngine.Execute(com, app, writes, options);
        string? warning = null;
        try
        {
            dynamic sheet = writes[^1].Sheet;
            if (originalSelection)
            {
                dynamic selected = com.Own(window.RangeSelection), areas = com.Own(selected.Areas);
                int consumed = Math.Min(tables.Count, (int)areas.Count);
                if (tables.Count < (int)areas.Count)
                { dynamic remaining = com.Own(areas.Item(consumed + 1)); for (int i = consumed + 2; i <= (int)areas.Count; i++) remaining = com.Own(app.Union(remaining, com.Own(areas.Item(i)))); remaining.Select(); }
                else { dynamic cell = com.Own(sheet.Cells.Item(Math.Min(1048576, writes[^1].Row + writes[^1].Table.RowCount), writes[^1].Column)); cell.Select(); }
                if (tables.Count > (int)areas.Count) warning = "選択範囲がファイル数より少ないため、残りを最後の範囲の下へ１行空けて順に展開しました。";
            }
            else { sheet.Activate(); dynamic cell = com.Own(sheet.Cells.Item(Math.Min(1048576, writes[^1].Row + writes[^1].Table.RowCount), writes[^1].Column)); cell.Select(); }
        }
        catch { warning = "展開は完了しましたが、次のセルの選択に失敗しました。"; }
        dynamic lastSheet = writes[^1].Sheet;
        return new($"{tables.Count}ファイルを展開しました（{lastSheet.Name}）。", warning);
    }
}
