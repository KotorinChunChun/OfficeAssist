using System.Runtime.InteropServices;

namespace OfficeAssist;

internal static class ExcelTabularInserter
{
    internal static ExcelPictureInserter.Result? Insert(nint handle, Point point, ReceivedTabularBatch received, bool showOptions = true)
    {
        if (!Native.IsWindow(handle) || Native.ClassName(handle) != "EXCEL7" || !Native.IsWindowEnabled(Native.GetAncestor(handle, 2)))
            throw new InvalidOperationException("対象のExcelが閉じられたか、ダイアログを表示しています。");
        using var com = new ComScope(); dynamic window = ExcelConnection.Attach(com, handle), app = com.Own(window.Application);
        object? hit = window.RangeFromPoint(point.X, point.Y);
        if (hit is null) throw new InvalidOperationException("ワークシートのセルへドロップしてください。");
        com.Own(hit);
        var defaults = TabularDropSettings.Load();
        try { dynamic cell = hit, rows = com.Own(cell.Rows), columns = com.Own(cell.Columns); if ((int)rows.Count != 1 || (int)columns.Count != 1) throw new InvalidOperationException("見出しではなくセルへドロップしてください。"); }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { throw new InvalidOperationException("ワークシートのセルへドロップしてください。"); }
        var owner = new ExcelWindowOwner(Native.GetAncestor(handle, 2));
        using var locks = new ExcelDropLocks();
        locks.LockExisting(owner.Handle);
        IReadOnlyList<TabularData>? sources = null;
        InputSelection? inputSelection = null;
        try
        {
            ExcelPictureInserter.Result ImportWithoutDialog(IReadOnlyList<TabularData> tables, TabularImportOptions options)
            {
                using var cancellation = new CancellationTokenSource();
                return ImportOutsideDialog(com, app, window, tables, options with { Cancellation = cancellation.Token,
                    Progress = _ => { Application.DoEvents(); if ((Native.GetAsyncKeyState(27) & 0x8000) != 0) cancellation.Cancel(); } }, locks);
            }
            if (received.HasFolders)
            {
                using var chooser = new FileSelectionForm(received.Inputs) { StartPosition = FormStartPosition.CenterParent };
                if (chooser.ShowDialog(owner) != DialogResult.OK) return null;
                inputSelection = chooser.Selection;
            }
            else
            {
                inputSelection = new(received.Inputs.Select(path => { var info = new FileInfo(path); return new InputCandidate(path, info.Length, info.LastWriteTimeUtc) { Encoding = defaults.Encoding }; }));
            }
            IReadOnlyList<TabularData>? Load(InputCandidate[] files, IWin32Window parent)
            {
                using var load = new TableLoadForm<IReadOnlyList<TabularData>>("CSV・TSVの読込み", token => received.Read(files, token)) { StartPosition = FormStartPosition.CenterParent };
                if (load.ShowDialog(parent) != DialogResult.OK) { if (load.Error is not null) throw load.Error; return null; }
                return load.Result;
            }
            while (sources is null)
            {
                try { sources = Load(inputSelection!.Selected, owner); if (sources is null) return null; }
                catch (Exception ex)
                {
                    MessageBox.Show(owner, ex.Message + Environment.NewLine + "対象ファイルと文字コードを確認してください。", WindowTitle.Create("CSV・TSVの読込み"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    using var chooser = new FileSelectionForm(received.Inputs, inputSelection) { StartPosition = FormStartPosition.CenterParent };
                    if (chooser.ShowDialog(owner) != DialogResult.OK) return null;
                    inputSelection = chooser.Selection;
                }
            }
            if (!showOptions)
            {
                ExcelPictureInserter.Result? automatic = null;
                try
                {
                    var prepared = defaults.Prepare(sources);
                    automatic = ImportWithoutDialog(prepared.Tables, prepared.Options);
                }
                catch (Exception ex)
                {
                    if (defaults.Destination == ImportDestination.NewBook)
                    {
                        locks.ReleaseAll();
                        locks.LockExisting(owner.Handle);
                    }
                    MessageBox.Show(owner, ex.Message + Environment.NewLine + "取り込み設定を確認してください。", WindowTitle.Create("CSV・TSVの取り込み"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                if (automatic is not null)
                {
                    if (automatic.Warning is not null) MessageBox.Show(owner, automatic.Warning, WindowTitle.Create("CSV・TSVの取り込み"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return automatic;
                }
            }
            var targets = new ExcelTargetEditor(com, window);
            targets.TargetWindowChanged = locks.LockExisting;
            targets.LockCurrentWindows();
            var previewContext = new ExcelPreviewContextCache(window);
            TabularImportForm? dialog = null;
            using (dialog = new TabularImportForm(sources, options =>
            {
                if (!Native.IsWindow(handle)) throw new InvalidOperationException("対象Excelを確認できません。");
                // 新しいブックだけはモーダル画面が閉じてから作る。
                if (options.Destination == ImportDestination.NewBook) return null;
                return Import(com, app, window, dialog!.Tables, options);
            }, targets: targets, reselect: () =>
            {
                using var chooser = new FileSelectionForm(received.Inputs, inputSelection) { StartPosition = FormStartPosition.CenterParent };
                if (chooser.ShowDialog(dialog) != DialogResult.OK) return null;
                var replacement = Load(chooser.Selected, dialog!); if (replacement is null) return null;
                foreach (var table in sources) table.Rows.Dispose(); sources = replacement; inputSelection = chooser.Selection; return replacement;
            }, defer: true, readEncoding: path =>
            {
                var choices = inputSelection!.Selected.Where(f => path is null || f.Path.Equals(path, StringComparison.OrdinalIgnoreCase)).Select(f => f.Encoding).Distinct().ToArray();
                return choices.Length == 1 ? (int)choices[0] : 5;
            }, reencode: (path, encoding) =>
            {
                var changed = inputSelection!.Files.Select(f => inputSelection.IsChecked(f.Path) && (path is null || f.Path.Equals(path, StringComparison.OrdinalIgnoreCase)) ? f with { Encoding = encoding } : f).ToArray();
                var replacement = Load(changed.Where(f => inputSelection.IsChecked(f.Path)).ToArray(), dialog!);
                if (replacement is null) return null;
                var next = new InputSelection(changed); foreach (var file in changed) next.Check(file.Path, inputSelection.IsChecked(file.Path));
                foreach (var table in sources) table.Rows.Dispose(); sources = replacement; inputSelection = next; return replacement;
            }, previewContext: previewContext.Capture, defaults: defaults))
            {
                if (dialog.ShowDialog(owner) != DialogResult.OK || dialog.Options is not { } options) return null;
                var result = options.Destination == ImportDestination.NewBook ? ImportWithoutDialog(dialog.Tables, options) : dialog.Result;
                if (result is null) return null;
                var encodings = inputSelection!.Selected.Select(f => f.Encoding).Distinct().ToArray();
                try { TabularDropSettings.Save(dialog.CreateDropProfile(encodings.Length == 1 ? encodings[0] : TableEncoding.Auto, encodings.Length > 1)); }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException)
                {
                    string warning = "取り込みは完了しましたが、次回の設定を保存できませんでした。 " + ex.Message;
                    MessageBox.Show(owner, warning, WindowTitle.Create("CSV・TSVの取り込み"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return result with { Warning = warning };
                }
                return result;
            }
        }
        finally { if (sources is not null) foreach (var table in sources) table.Rows.Dispose(); }
    }

    internal static ExcelPictureInserter.Result ImportOutsideDialog(ComScope com, object application, object sourceWindow,
        IReadOnlyList<TabularData> tables, TabularImportOptions options, ExcelDropLocks locks)
    {
        if (options.Destination == ImportDestination.NewBook) locks.ReleaseAll();
        return Import(com, application, sourceWindow, tables, options, locks.LockCreated);
    }

    internal static ExcelPictureInserter.Result Import(ComScope com, dynamic app, object? sourceWindow,
        IReadOnlyList<TabularData> tables, TabularImportOptions options, Action<nint>? lockCreated = null)
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
            {
                dynamic books = com.Own(app.Workbooks);
                createdBook = com.Own(books.Add(-4167));
                LockCreatedWindows(com, createdBook, lockCreated);
                book = createdBook;
                book.Date1904 = false;
            }
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
            if (createdBook is not null) LockCreatedWindows(com, createdBook, lockCreated);
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

    static void LockCreatedWindows(ComScope com, object book, Action<nint>? lockCreated)
    {
        if (lockCreated is null) return;
        dynamic workbook = book, windows = com.Own(workbook.Windows);
        for (int i = 1; i <= (int)windows.Count; i++)
        {
            dynamic window = com.Own(windows.Item(i));
            lockCreated((nint)(int)window.Hwnd);
        }
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
