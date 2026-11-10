using OfficeAssist;

internal static partial class Program
{
    static void TabularModalExcelTests(ComScope com, dynamic window, dynamic sheet)
    {
        dynamic app = com.Own(window.Application), book = com.Own(sheet.Parent), sheets = com.Own(book.Worksheets);
        sheet.Activate(); dynamic cell = com.Own(sheet.Range["A140"]); cell.Value2 = "既存データ"; cell.Select();
        nint handle = (nint)(int)app.Hwnd;
        bool enabled = Native.IsWindowEnabled(handle);
        int warnings = 0, before = (int)sheets.Count;
        using (var modalLock = new ExcelModalLock(handle))
        {
            Check(!Native.IsWindowEnabled(handle), "設定処理中の実Excelをロックする");
            using (var nested = new ExcelModalLock(handle)) { }
            Check(!Native.IsWindowEnabled(handle), "子エラーのロック終了で外側のExcelロックを解除しない");
            TabularData[] data = [new("retry.csv", [["値"], ["001"]])];
            using var form = new TabularImportForm(data, options => ExcelTabularInserter.Import(com, app, window, data, options), (owner, message, icon) =>
            {
                warnings++;
                Check(owner is TabularImportForm f && f.TopMost && !f.IsDisposed && !Native.IsWindowEnabled(handle), "エラーの所有元は前面の設定画面でExcelロックを維持");
                Check(message.Contains("既存") || message.Contains("値・数式"), "セル衝突のエラー内容を渡す");
            });
            Check(!form.TryImport() && warnings == 1 && !form.IsDisposed && form.DialogResult == DialogResult.None, "衝突で設定画面を閉じずに再試行可能にする");
            Check((string)cell.Value2 == "既存データ" && !Native.IsWindowEnabled(handle), "エラー後も既存データとExcelロックを維持");
            var destination = (ComboBox)form.Controls.Find("ImportDestination", true).Single(); destination.SelectedIndex = (int)ImportDestination.NewSheet;
            Check(form.TryImport() && (int)sheets.Count == before + 1 && form.Result is not null, "エラー後に新しいシートへ変更して取り込める");
            Check(!Native.IsWindowEnabled(handle), "取り込み完了までExcelロックを維持");
        }
        Check(Native.IsWindowEnabled(handle) == enabled, "外側の処理終了で実Excelの操作状態を復元する");
    }

    static void NewBookModalExcelTests(ComScope com, dynamic window, dynamic sheet)
    {
        sheet.Activate();
        dynamic app = com.Own(window.Application);
        nint sourceHandle = (nint)(int)window.Hwnd;
        using var unrelated = new Control();
        nint disabledHandle = unrelated.Handle;
        Native.EnableWindow(disabledHandle, false);
        TabularData[] data = [new("new-book-lock.csv", [["番号"], ["001"]])];
        foreach (bool inheritedDisabled in new[] { false, true })
        {
            nint createdHandle = 0;
            using (var locks = new ExcelDropLocks())
            {
                locks.LockExisting(sourceHandle);
                locks.LockExisting(disabledHandle);
                Action<nint> track = handle =>
                {
                    if (inheritedDisabled) Native.EnableWindow(handle, false);
                    locks.LockCreated(handle);
                    createdHandle = handle;
                };
                var options = new TabularImportOptions(ImportDestination.NewBook, [[ColumnType.TextFormat]], true);
                if (inheritedDisabled) ExcelTabularInserter.Import(com, app, window, data, options, track);
                else
                {
                    bool? unlockedBeforeCreation = null;
                    options = options with { Progress = _ => unlockedBeforeCreation ??= Native.IsWindowEnabled(sourceHandle) };
                    ExcelTabularInserter.ImportOutsideDialog(com, app, window, data, options, locks);
                    dynamic createdWindow = com.Own(app.ActiveWindow);
                    createdHandle = (nint)(int)createdWindow.Hwnd;
                    Check(unlockedBeforeCreation == true, "モーダル画面終了後は元のExcelを解除してから新しいブックを作る");
                }
                Check(createdHandle != 0 && createdHandle != sourceHandle && !Native.IsWindowEnabled(createdHandle), "作成したExcelウィンドウも取り込み中は操作をロックする");
            }
            dynamic book = com.Own(app.ActiveWorkbook), resultSheet = com.Own(book.Worksheets.Item(1));
            try
            {
                Check(Native.IsWindowEnabled(createdHandle) && Native.IsWindowEnabled(sourceHandle) && !Native.IsWindowEnabled(disabledHandle),
                    inheritedDisabled ? "作成時に無効だった新しいブックを操作可能にして別の無効状態を保持する" : "取り込み終了後に元と新しいブックの操作を戻す");
                Check((string)resultSheet.Range["A2"].Value2 == "001", "ロックを復元した新しいブックにデータを保持する");
            }
            finally { book.Close(false); }
        }
    }
}
