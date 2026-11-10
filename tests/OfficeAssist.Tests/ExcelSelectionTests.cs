using OfficeAssist;
using System.Diagnostics;

internal static partial class Program
{
    static void ExcelSelectionTests(string output)
    {
        var existing = Process.GetProcessesByName("EXCEL").Select(p => { using (p) return p.Id; }).ToHashSet();
        using var com = new ComScope();
        dynamic app = com.Own(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!);
        Native.GetWindowThreadProcessId((nint)(int)app.Hwnd, out uint pid);
        Console.WriteLine($"専用の非表示Excel PID={pid}");
        if (existing.Contains((int)pid)) throw new InvalidOperationException("既存Excelを保護するため選択範囲の検証を中止しました。");
        try
        {
            app.Visible = false;
            app.DisplayAlerts = false;
            dynamic books = com.Own(app.Workbooks);
            dynamic book = com.Own(books.Add());
            try
            {
                dynamic sheet = com.Own(book.ActiveSheet);
                dynamic window = com.Own(app.ActiveWindow);
                dynamic selected = com.Own(sheet.Range["A1:B2,D1:E2,G1:H2"]);
                selected.Select();
                dynamic areas = com.Own(selected.Areas);
                var destinations = new List<object>();
                for (int i = 1; i <= (int)areas.Count; i++) destinations.Add(com.Own(areas.Item(i)));
                ExcelPictureInserter.AdvanceSelection(com, window, sheet, selected, destinations, 1);
                dynamic after = com.Own(window.RangeSelection);
                dynamic afterAreas = com.Own(after.Areas);
                Check((int)afterAreas.Count == 2 && (string)after.Address == "$D$1:$E$2,$G$1:$H$2", "実Excelで使用済みの先頭範囲を選択から外す");
                ExcelPictureInserter.AdvanceSelection(com, window, sheet, selected, destinations, 3);
                dynamic exhausted = com.Own(window.RangeSelection);
                dynamic intersection = app.Intersect(selected, exhausted);
                Check(intersection is null, "実Excelで全範囲を使い切ったら選択外へ移す");
                if (intersection is not null) com.Own(intersection);
                TabularExcelTests(com, window, sheet);
                TabularFormatExcelTests(com, window, sheet);
                TabularModalExcelTests(com, window, sheet);
                NewBookModalExcelTests(com, window, sheet);
                HiddenBookExcelTests(com, window, sheet);
                V3ExcelTests(com, window, sheet);
                V22FixtureExcelTests(com, window, sheet, output);
                DropDialogExcelTests(com, window, sheet);
            }
            finally { try { book.Close(false); } catch (System.Runtime.InteropServices.COMException) { } }
        }
        finally { try { app.Quit(); } catch (System.Runtime.InteropServices.COMException) { } }
    }
}
