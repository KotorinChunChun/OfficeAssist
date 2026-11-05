using System.Runtime.InteropServices;

namespace OfficeAssist;

internal sealed class ComScope : IDisposable
{
    readonly List<object> objects = [];
    internal dynamic Own(object value) { objects.Add(value); return value; }
    public void Dispose()
    {
        for (int i = objects.Count - 1; i >= 0; i--)
            if (Marshal.IsComObject(objects[i])) Marshal.ReleaseComObject(objects[i]);
    }
}

internal static class ExcelPictureInserter
{
    internal sealed record Result(string Message, string? Warning = null);
    internal static string Insert(nint sheetHandle, Point screenPoint, IReadOnlyList<string> files) =>
        InsertCore(sheetHandle, screenPoint, files, false)!.Message;

    internal static Result? InsertInteractive(nint sheetHandle, Point screenPoint, IReadOnlyList<string> files,
        IReadOnlyList<string>? linkedFiles = null, bool allowLinkedImages = true, bool showOptions = false) =>
        InsertCore(sheetHandle, screenPoint, files, true, linkedFiles, allowLinkedImages, showOptions);

    static Result? InsertCore(nint sheetHandle, Point screenPoint, IReadOnlyList<string> files, bool interactive,
        IReadOnlyList<string>? linkedFiles = null, bool allowLinkedImages = true, bool showOptions = false)
    {
        ImageFiles.Validate(files);
        if (!Native.IsWindow(sheetHandle) || Native.ClassName(sheetHandle) != "EXCEL7"
            || !Native.IsWindowEnabled(Native.GetAncestor(sheetHandle, 2)))
            throw new InvalidOperationException("対象のExcelが閉じられたか、ダイアログを表示しています。");

        using var com = new ComScope();
        var iid = new Guid("00020400-0000-0000-C000-000000000046");
        int hr = Native.AccessibleObjectFromWindow(sheetHandle, 0xFFFFFFF0, ref iid, out object nativeWindow);
        Marshal.ThrowExceptionForHR(hr);
        dynamic window = com.Own(nativeWindow);
        dynamic sheet = com.Own(window.ActiveSheet);
        if ((int)sheet.Type != -4167) throw new InvalidOperationException("ワークシートのセルへドロップしてください。");
        dynamic book = com.Own(sheet.Parent);
        if ((bool)book.ReadOnly || (bool)sheet.ProtectContents || (bool)sheet.ProtectDrawingObjects)
            throw new InvalidOperationException("読み取り専用ブックや保護されたシートには挿入できません。");
        object? hitObject = window.RangeFromPoint(screenPoint.X, screenPoint.Y);
        if (hitObject is null) throw new InvalidOperationException("シートのセルへドロップしてください。");
        dynamic hit = com.Own(hitObject);
        dynamic anchor;
        try { anchor = com.Own(hit.Cells); }
        catch (Exception ex) when (ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        { anchor = com.Own(hit.TopLeftCell); }
        // 行列ヘッダーはセル範囲として返る場合があるので、全行・全列を拒否する。
        dynamic rows = com.Own(anchor.Rows);
        dynamic columns = com.Own(anchor.Columns);
        if ((int)rows.Count > 1 || (int)columns.Count > 1)
            throw new InvalidOperationException("行列の見出しではなく、1つのセルへドロップしてください。");
        string sheetName = (string)sheet.Name;
        dynamic selected = com.Own(window.RangeSelection);
        dynamic areas = com.Own(selected.Areas);
        string selectedAddress = (string)selected.Address;
        InsertOptions options;
        if (interactive)
        {
            options = InsertSettings.Load();
            if (showOptions)
            {
                using var dialog = new InsertOptionsForm(ExcelAddress.Display(selected), files.Count, allowLinkedImages, options);
                if (dialog.ShowDialog() != DialogResult.OK || dialog.Options is null) return null;
                options = dialog.Options;
                InsertSettings.Save(options);
            }
            if (!Native.IsWindow(sheetHandle) || !Native.IsWindowEnabled(Native.GetAncestor(sheetHandle, 2)))
                throw new InvalidOperationException("対象のExcelが閉じられたか、ダイアログを表示しています。");
            dynamic currentSheet = com.Own(window.ActiveSheet);
            dynamic currentSelection = com.Own(window.RangeSelection);
            if ((string)currentSheet.Name != sheetName || (string)currentSelection.Address != selectedAddress)
                throw new InvalidOperationException("確認中にシートまたは選択範囲が変わりました。もう一度ドロップしてください。");
            if ((bool)book.ReadOnly || (bool)sheet.ProtectContents || (bool)sheet.ProtectDrawingObjects)
                throw new InvalidOperationException("読み取り専用ブックや保護されたシートには挿入できません。");
        }
        else options = new InsertOptions(Destination.DropPosition, PictureKind.Embedded, []);
        if (options.Kind == PictureKind.ImageFormula && options.Urls.Length != files.Count)
            throw new InvalidOperationException($"IMAGE関数には画像URLが必要です。{InsertSettings.LoadModifier()}を押しながらドロップし、URLを指定してください。");
        var destinations = new List<object>();
        if (options.Destination == Destination.DropPosition) destinations.Add(anchor);
        else for (int i = 1; i <= (int)areas.Count; i++) destinations.Add(com.Own(areas.Item(i)));
        string address = options.Destination == Destination.DropPosition ? ExcelAddress.Display(anchor) : ExcelAddress.Display(selected);
        var assignments = ImageAssignment.Create(destinations.Count, files.Count);
        var warnings = new List<string>();
        if (options.Destination != Destination.DropPosition && files.Count > destinations.Count)
            warnings.Add("選択範囲が画像数より少ないため、残りの画像を最後の選択範囲にまとめて挿入しました。");
        string message;
        if (options.Kind == PictureKind.ImageFormula)
        {
            message = InsertFormulas(com, destinations, assignments, options, sheetName, address);
            return Finish(com, window, sheet, selected, destinations, assignments, options, message, warnings);
        }
        var insertFiles = options.Kind == PictureKind.Linked ? linkedFiles ?? files : files;
        if (options.Kind == PictureKind.Linked)
        {
            if (!allowLinkedImages) throw new InvalidOperationException("外部アプリから受け取った画像は埋め込みを選択してください。");
            ImageFiles.Validate(insertFiles);
        }
        dynamic shapes = com.Own(sheet.Shapes);
        var added = new List<object>();
        try
        {
            var offsets = new int[destinations.Count];
            var nextTops = new float[destinations.Count];
            foreach (var (area, fileIndex) in assignments)
            {
                dynamic destination = destinations[area];
                int countInArea = assignments.Count(x => x.Area == area);
                int offset = offsets[area]++;
                float left = (float)destination.Left, top = (float)destination.Top;
                if (options.Destination == Destination.SelectionRange) top += (float)destination.Height / countInArea * offset;
                else if (offset > 0) top = nextTops[area];
                dynamic picture = com.Own(shapes.AddPicture(Path.GetFullPath(insertFiles[fileIndex]), options.Kind == PictureKind.Linked ? -1 : 0,
                    options.Kind == PictureKind.Linked ? 0 : -1, left, top, -1, -1));
                added.Add(picture);
                float width = (float)picture.Width, height = (float)picture.Height;
                float ratio = interactive ? 1 : Math.Min(1, 480 / Math.Max(width, height));
                if (options.Destination == Destination.SelectionRange)
                    ratio = Math.Min((float)destination.Width / width, ((float)destination.Height / countInArea) / height);
                if (!float.IsFinite(ratio) || ratio <= 0) throw new InvalidOperationException("選択範囲のサイズが不正です。");
                picture.LockAspectRatio = -1;
                picture.Width = width * ratio;
                picture.Placement = 3; // xlFreeFloating: 行列サイズ変更で縦横比を壊さない。
                nextTops[area] = top + height * ratio + 12;
            }
        }
        catch (Exception error)
        {
            bool rolledBack = true;
            foreach (dynamic picture in added.AsEnumerable().Reverse())
            {
                try { picture.Delete(); }
                catch (Exception) { rolledBack = false; }
            }
            throw new InvalidOperationException(rolledBack
                ? "画像を挿入できませんでした。今回の追加は取り消しました。Excelの編集状態と画像を確認してください。"
                : "挿入に失敗し、一部の画像を取り消せませんでした。対象シートを確認してください。", error);
        }
        message = $"{assignments.Count}枚を挿入しました（{sheetName} {address}）。";
        return Finish(com, window, sheet, selected, destinations, assignments, options, message, warnings);
    }

    static Result Finish(ComScope com, dynamic window, dynamic sheet, dynamic selected, List<object> destinations,
        List<(int Area, int File)> assignments, InsertOptions options, string message, List<string> warnings)
    {
        if (options.Destination != Destination.DropPosition)
        {
            try { AdvanceSelection(com, window, sheet, selected, destinations, Math.Min(destinations.Count, assignments.Count)); }
            catch (Exception ex) when (ex is COMException or InvalidOperationException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            { warnings.Add("画像は挿入しましたが選択範囲を更新できませんでした。次の挿入先を選び直してください。"); }
        }
        return new Result(message, warnings.Count == 0 ? null : string.Join(Environment.NewLine, warnings));
    }

    internal static void AdvanceSelection(ComScope com, dynamic window, dynamic sheet, dynamic selected,
        List<object> destinations, int consumed)
    {
        dynamic app = com.Own(window.Application);
        if (consumed < destinations.Count)
        {
            dynamic remaining = destinations[consumed];
            for (int i = consumed + 1; i < destinations.Count; i++) remaining = com.Own(app.Union(remaining, destinations[i]));
            remaining.Select();
            return;
        }
        // Excelには空のセル選択がないため、元の選択範囲に重ならない隣接セルへ移る。
        dynamic last = destinations[^1];
        dynamic rows = com.Own(last.Rows);
        dynamic columns = com.Own(last.Columns);
        int row = (int)last.Row, column = (int)last.Column;
        int bottom = row + (int)rows.Count - 1, right = column + (int)columns.Count - 1;
        (int Row, int Column)[] candidates = [(row, right + 1), (bottom + 1, column), (row, column - 1), (row - 1, column)];
        dynamic sheetCells = com.Own(sheet.Cells);
        foreach (var candidate in candidates)
        {
            if (candidate.Row < 1 || candidate.Row > 1048576 || candidate.Column < 1 || candidate.Column > 16384) continue;
            dynamic cell = com.Own(sheetCells.Item(candidate.Row, candidate.Column));
            object? intersect = app.Intersect(selected, cell);
            if (intersect is not null) { com.Own(intersect); continue; }
            cell.Select();
            return;
        }
        throw new InvalidOperationException("選択外のセルが見つかりません。");
    }

    static string InsertFormulas(ComScope com, List<object> destinations, List<(int Area, int File)> assignments, InsertOptions options, string sheetName, string address)
    {
        var targets = new List<object>();
        var offsets = new int[destinations.Count];
        var addresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (area, _) in assignments)
        {
            dynamic destination = destinations[area];
            dynamic cells = com.Own(destination.Cells);
            dynamic rows = com.Own(destination.Rows);
            dynamic columns = com.Own(destination.Columns);
            int rowCount = (int)rows.Count, columnCount = (int)columns.Count;
            int offset = offsets[area]++;
            bool withinRange = options.Destination == Destination.SelectionRange;
            if (withinRange && (long)rowCount * columnCount <= offset)
                throw new InvalidOperationException("選択範囲のセル数が画像数より少ないため挿入できません。");
            dynamic cell = com.Own(cells.Item(withinRange ? offset / columnCount + 1 : offset + 1,
                withinRange ? offset % columnCount + 1 : 1));
            if (!addresses.Add((string)cell.Address)) throw new InvalidOperationException("IMAGE関数の挿入先が重複しています。選択範囲を変更してください。");
            if ((bool)cell.MergeCells || cell.Value2 is not null || (bool)cell.HasFormula)
                throw new InvalidOperationException("結合セルや値の入ったセルにはIMAGE関数を挿入できません。空のセルを選択してください。");
            targets.Add(cell);
        }
        var written = new List<object>();
        try
        {
            for (int i = 0; i < assignments.Count; i++)
            {
                dynamic cell = targets[i];
                string url = options.Urls[assignments[i].File].Replace("\"", "\"\"");
                cell.Formula = "=IMAGE(\"" + url + "\",\"\"," + (options.Destination == Destination.SelectionRange ? "0" : "2") + ")";
                written.Add(cell);
            }
        }
        catch (Exception ex)
        {
            bool restored = true;
            foreach (dynamic cell in written) { try { cell.ClearContents(); } catch (Exception) { restored = false; } }
            throw new InvalidOperationException(restored ? "IMAGE関数を挿入できませんでした。今回の追加を取り消しました。"
                : "IMAGE関数の挿入と取り消しに失敗しました。対象シートを確認してください。", ex);
        }
        return $"{assignments.Count}件のIMAGE関数を挿入しました（{sheetName} {address}）。";
    }
}
