using System.Runtime.InteropServices;

namespace OfficeAssist;

internal static class ExcelTabularInserter
{
    internal static ExcelPictureInserter.Result Insert(nint handle, Point point, IReadOnlyList<TabularData> tables)
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
        return InsertIntoWindow(com, window, tables);
    }

    internal static ExcelPictureInserter.Result InsertIntoWindow(ComScope com, dynamic window, IReadOnlyList<TabularData> tables)
    {
        if (tables.Count == 0) throw new InvalidOperationException("展開するデータがありません。");
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
        try
        {
            app.EnableEvents = false;
            for (int i = 0; i < targets.Count; i++)
            {
                dynamic target = targets[i];
                if ((double)functions.CountA(target) != 0) throw new InvalidOperationException("処理中に展開先が変更されました。");
                written.Add(target);
                target.Value2 = tables[i].ExcelValues();
            }
        }
        catch (Exception ex)
        {
            bool restored = true;
            foreach (dynamic target in written) { try { target.ClearContents(); } catch (Exception) { restored = false; } }
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
