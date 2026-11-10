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

    static void HiddenBookExcelTests(ComScope com, dynamic window, dynamic sheet)
    {
        dynamic app = com.Own(window.Application), books = com.Own(app.Workbooks);
        dynamic hidden = com.Own(books.Add(-4167)), hiddenWindow = com.Own(hidden.Windows.Item(1));
        dynamic addIn = com.Own(books.Add(-4167));
        try
        {
            addIn.IsAddin = true;
            hiddenWindow.Visible = false;
            sheet.Activate();
            using var target = new ExcelTargetEditor(com, window);
            var choices = (ComboBox)target.Controls.Find("TargetBook", true).Single();
            Check(!choices.Items.Contains((string)hidden.Name) && !choices.Items.Contains((string)addIn.Name)
                && !ExcelTargetEditor.VisibleWindows(com, (object)addIn).Any(), "アドインと非表示ブックを展開先・ロック対象に含めない");
            using var locks = new ExcelDropLocks();
            TabularData[] data = [new("hidden-book.csv", [["値"], ["001"]])];
            var options = new TabularImportOptions(ImportDestination.NewBook, [[ColumnType.TextFormat]], true);
            int countBefore = (int)books.Count;
            ExcelTabularInserter.ImportOutsideDialog(com, app, window, data, options, locks);
            dynamic created = com.Own(books.Item(countBefore + 1));
            try
            {
                Check((bool)app.Interactive && !(bool)hiddenWindow.Visible && (bool)addIn.IsAddin,
                    "非表示ブックとアドインが存在しても新規取込後にExcelの操作状態を復元する");
            }
            finally { created.Close(false); }
        }
        finally { addIn.IsAddin = false; addIn.Close(false); hidden.Close(false); }
    }

    static void NewBookModalExcelTests(ComScope com, dynamic window, dynamic sheet)
    {
        sheet.Activate();
        dynamic app = com.Own(window.Application);
        dynamic books = com.Own(app.Workbooks), otherBook = com.Own(books.Add(-4167));
        dynamic otherWindow = com.Own(app.ActiveWindow);
        nint sourceHandle = (nint)(int)window.Hwnd;
        nint otherHandle = (nint)(int)otherWindow.Hwnd;
        sheet.Activate();
        using var unrelated = new Control();
        nint disabledHandle = unrelated.Handle;
        Native.EnableWindow(disabledHandle, false);
        TabularData[] data = [new("new-book-lock.csv", [["番号"], ["001"]])];
        var options = new TabularImportOptions(ImportDestination.NewBook, [[ColumnType.TextFormat]], true);
        try
        {
            using var locks = new ExcelDropLocks();
            locks.LockExisting(sourceHandle);
            locks.LockExisting(otherHandle);
            locks.LockExisting(disabledHandle);
            int initialBooks = (int)books.Count;
            bool? unlockedBeforeCreation = null, interactiveDuringWrite = null;
            ExcelTabularInserter.ImportOutsideDialog(com, app, window, data, options with
            {
                Progress = stage =>
                {
                    if (stage.Contains("検証中")) unlockedBeforeCreation ??= Native.IsWindowEnabled(sourceHandle) && Native.IsWindowEnabled(otherHandle) && (bool)app.Interactive;
                    if (stage.Contains("書込み中")) interactiveDuringWrite = (bool)app.Interactive;
                }
            }, locks);
            dynamic book = com.Own(books.Item(initialBooks + 1)), resultSheet = com.Own(book.Worksheets.Item(1));
            dynamic createdWindow = com.Own(book.Windows.Item(1));
            nint createdHandle = (nint)(int)createdWindow.Hwnd;
            try
            {
                Check(unlockedBeforeCreation == true && interactiveDuringWrite == false, "新規作成前に既存ブックを解除し、書込み中だけExcelが入力を止める");
                Check((bool)app.Interactive && createdHandle != sourceHandle && createdHandle != otherHandle
                    && Native.IsWindowEnabled(createdHandle) && Native.IsWindowEnabled(sourceHandle) && Native.IsWindowEnabled(otherHandle)
                    && !Native.IsWindowEnabled(disabledHandle), "新しいブックを開いたまま既存の2ブックも操作可能にする");
                dynamic sourceCell = com.Own(sheet.Range["A141"]);
                sourceCell.Value2 = "元のブック";
                Check((string)sourceCell.Value2 == "元のブック" && (int)books.Count == initialBooks + 1,
                    "新しいブックを開いたまま元のブックのデータを操作できる");
                dynamic resultCell = com.Own(resultSheet.Range["A2"]);
                Check((string)resultCell.Value2 == "001", "新しいブックに書込み結果を保持する");
            }
            finally { book.Close(false); }

            int before = (int)books.Count;
            using var cancellation = new CancellationTokenSource();
            locks.LockExisting(sourceHandle);
            bool failed = false;
            try
            {
                ExcelTabularInserter.ImportOutsideDialog(com, app, window, data, options with
                {
                    Cancellation = cancellation.Token,
                    Progress = stage => { if (stage.Contains("書込み中")) cancellation.Cancel(); }
                }, locks);
            }
            catch (InvalidOperationException ex) when (ex.GetBaseException() is OperationCanceledException) { failed = true; }
            Check(failed && (int)books.Count == before && (bool)app.Interactive && Native.IsWindowEnabled(sourceHandle) && Native.IsWindowEnabled(otherHandle),
                "書込み中のキャンセルで新規ブックを取り消し、既存ブックの操作を戻す");
        }
        finally { otherBook.Close(false); }
    }
}
