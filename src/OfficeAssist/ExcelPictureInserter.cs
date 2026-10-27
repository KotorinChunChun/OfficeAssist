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
    internal static string Insert(nint sheetHandle, Point screenPoint, IReadOnlyList<string> files)
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
        float left = (float)anchor.Left, top = (float)anchor.Top;
        string address = (string)anchor.Address;
        string sheetName = (string)sheet.Name;
        dynamic shapes = com.Own(sheet.Shapes);
        var added = new List<object>();
        try
        {
            foreach (var file in files)
            {
                dynamic picture = com.Own(shapes.AddPicture(Path.GetFullPath(file), 0, -1, left, top, -1, -1));
                added.Add(picture);
                var fit = ImageFiles.Fit((float)picture.Width, (float)picture.Height);
                picture.LockAspectRatio = -1;
                picture.Width = fit.Width;
                picture.Placement = 3; // xlFreeFloating: 行列サイズ変更で縦横比を壊さない。
                top += fit.Height + 12;
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
}
