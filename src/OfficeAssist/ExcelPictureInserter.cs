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
    internal static string Insert(nint sheetHandle, Point screenPoint, IReadOnlyList<string> files) =>
        InsertCore(sheetHandle, screenPoint, files, false)!;

    internal static string? InsertInteractive(nint sheetHandle, Point screenPoint, IReadOnlyList<string> files,
        IReadOnlyList<string>? linkedFiles = null, bool allowLinkedImages = true) =>
        InsertCore(sheetHandle, screenPoint, files, true, linkedFiles, allowLinkedImages);

    static string? InsertCore(nint sheetHandle, Point screenPoint, IReadOnlyList<string> files, bool interactive,
        IReadOnlyList<string>? linkedFiles = null, bool allowLinkedImages = true)
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
        if (interactive && (int)areas.Count != 1) throw new InvalidOperationException("連続したセル範囲を選択してください。");
        string selectedAddress = (string)selected.Address;
        InsertOptions options;
        if (interactive)
        {
            using var dialog = new InsertOptionsForm(selectedAddress, files.Count, allowLinkedImages);
            if (dialog.ShowDialog() != DialogResult.OK || dialog.Options is null) return null;
            options = dialog.Options;
            if (!Native.IsWindow(sheetHandle) || !Native.IsWindowEnabled(Native.GetAncestor(sheetHandle, 2)))
                throw new InvalidOperationException("対象のExcelが閉じられたか、ダイアログを表示しています。");
            dynamic currentSheet = com.Own(window.ActiveSheet);
            dynamic currentSelection = com.Own(window.RangeSelection);
            if ((string)currentSheet.Name != sheetName || (string)currentSelection.Address != selectedAddress)
                throw new InvalidOperationException("確認中にシートまたは選択範囲が変わりました。もう一度ドロップしてください。");
            if ((bool)book.ReadOnly || (bool)sheet.ProtectContents || (bool)sheet.ProtectDrawingObjects)
                throw new InvalidOperationException("読み取り専用ブックや保護されたシートには挿入できません。");
        }
        else options = new InsertOptions(false, false, PictureKind.Embedded, []);
        dynamic destination = options.UseSelection ? selected : anchor;
        float left = (float)destination.Left, top = (float)destination.Top;
        string address = (string)destination.Address;
        if (options.Kind == PictureKind.ImageFormula)
            return InsertFormulas(com, sheet, destination, options, files.Count, sheetName, address);
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
            foreach (var file in insertFiles)
            {
                dynamic picture = com.Own(shapes.AddPicture(Path.GetFullPath(file), options.Kind == PictureKind.Linked ? -1 : 0,
                    options.Kind == PictureKind.Linked ? 0 : -1, left, top, -1, -1));
                added.Add(picture);
                float width = (float)picture.Width, height = (float)picture.Height;
                float ratio = interactive ? 1 : Math.Min(1, 480 / Math.Max(width, height));
                if (options.FitSelection)
                    ratio = Math.Min((float)selected.Width / width, ((float)selected.Height / files.Count) / height);
                if (!float.IsFinite(ratio) || ratio <= 0) throw new InvalidOperationException("選択範囲のサイズが不正です。");
                picture.LockAspectRatio = -1;
                picture.Width = width * ratio;
                picture.Placement = 3; // xlFreeFloating: 行列サイズ変更で縦横比を壊さない。
                top += options.FitSelection ? (float)selected.Height / files.Count : height * ratio + 12;
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
        return $"{files.Count}枚を挿入しました（{sheetName} {address}）。";
    }

    static string InsertFormulas(ComScope com, dynamic sheet, dynamic destination, InsertOptions options, int count, string sheetName, string address)
    {
        if (options.Urls.Length != count) throw new InvalidOperationException("画像URLの件数が一致しません。");
        dynamic cells = com.Own(destination.Cells);
        dynamic rows = com.Own(destination.Rows);
        dynamic columns = com.Own(destination.Columns);
        int rowCount = (int)rows.Count, columnCount = (int)columns.Count;
        if (options.UseSelection && (long)rowCount * columnCount < count)
            throw new InvalidOperationException("選択範囲のセル数が画像数より少ないため挿入できません。");
        var targets = new List<object>();
        for (int i = 0; i < count; i++)
        {
            dynamic cell = com.Own(cells.Item(options.UseSelection ? i / columnCount + 1 : i + 1,
                options.UseSelection ? i % columnCount + 1 : 1));
            if ((bool)cell.MergeCells || cell.Value2 is not null || (bool)cell.HasFormula)
                throw new InvalidOperationException("結合セルや値の入ったセルにはIMAGE関数を挿入できません。空のセルを選択してください。");
            targets.Add(cell);
        }
        var written = new List<object>();
        try
        {
            for (int i = 0; i < count; i++)
            {
                dynamic cell = targets[i];
                string url = options.Urls[i].Replace("\"", "\"\"");
                cell.Formula = "=IMAGE(\"" + url + "\",\"\"," + (options.FitSelection ? "0" : "2") + ")";
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
        return $"{count}件のIMAGE関数を挿入しました（{sheetName} {address}）。";
    }
}
