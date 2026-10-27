using OfficeAssist;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Diagnostics;

internal static class Program
{
    static int count;
    static void Check(bool value, string name)
    {
        if (!value) throw new Exception("FAIL: " + name);
        Console.WriteLine("PASS: " + name);
        count++;
    }
    static void Reject(Action action, string name)
    {
        bool rejected = false;
        try { action(); }
        catch (Exception ex) when (ex is InvalidDataException or InvalidOperationException) { rejected = true; }
        Check(rejected, name);
    }
    [STAThread]
    static int Main(string[] args)
    {
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        var output = Path.GetFullPath(args.LastOrDefault() ?? "dev/history/v0.1.0/logs/unit");
        Directory.CreateDirectory(output);
        try
        {
            var image = Path.Combine(output, "01-blue.png");
            using (var bitmap = new Bitmap(240, 120))
            {
                using var graphics = Graphics.FromImage(bitmap);
                graphics.Clear(Color.CornflowerBlue);
                graphics.DrawString("OfficeAssist 01", SystemFonts.DefaultFont, Brushes.White, 20, 40);
                bitmap.Save(image, ImageFormat.Png);
            }
            var image2 = Path.Combine(output, "02-orange.png");
            using (var bitmap = new Bitmap(120, 240))
            {
                using var graphics = Graphics.FromImage(bitmap);
                graphics.Clear(Color.Coral);
                graphics.DrawString("OfficeAssist 02", SystemFonts.DefaultFont, Brushes.White, 10, 40);
                bitmap.Save(image2, ImageFormat.Png);
            }
            Check(ImageFiles.Accepts([image, image2]), "複数画像を許可");
            Check(ImageFiles.Accepts(["UPPER.JPG", "image.tiff"]), "拡張子の大文字とTIFFを許可");
            Check(!ImageFiles.Accepts([]) && !ImageFiles.Accepts(null), "空入力を拒否");
            Check(!ImageFiles.Accepts([image, "book.xlsx"]) && !ImageFiles.Accepts(["folder"]), "非画像混在とフォルダーを拒否");
            Check(!ImageFiles.Accepts(Enumerable.Repeat(image, 51).ToArray()), "件数制限");
            ImageFiles.Validate([image, image2]);
            Check(true, "実画像をデコード検証");
            var invalid = Path.Combine(output, "invalid.png");
            File.WriteAllText(invalid, "これは画像ではありません。");
            Reject(() => ImageFiles.Validate([image, invalid]), "拡張子偽装を拒否");
            Reject(() => ImageFiles.Validate([Path.Combine(output, "missing.jpg")]), "存在しない画像を拒否");
            Check(ImageFiles.Fit(960, 480) == new SizeF(480, 240), "横長画像の縦横比と上限");
            Check(ImageFiles.Fit(100, 200) == new SizeF(100, 200), "小さい画像を拡大しない");
            Check(ImageFiles.Fit(240, 960) == new SizeF(120, 480), "縦長画像の縦横比と上限");
            Reject(() => ImageFiles.Fit(float.NaN, 1), "不正な寸法を拒否");
            if (args.Contains("--excel") || args.Contains("--fixture")) Excel(output, image, image2, invalid, args.Contains("--fixture"));
            Console.WriteLine($"PASS: {count}件");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }

    static void Excel(string output, string image, string image2, string invalid, bool fixture)
    {
        var existing = Process.GetProcessesByName("EXCEL").Select(p => { using (p) return p.Id; }).ToHashSet();
        using var com = new ComScope();
        dynamic app = com.Own(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!);
        Native.GetWindowThreadProcessId((nint)(int)app.Hwnd, out uint pid);
        if (existing.Contains((int)pid)) throw new InvalidOperationException("既存のExcelを保護するため実機試験を中止します。");
        try
        {
            app.Visible = true;
            app.DisplayAlerts = false;
            dynamic books = com.Own(app.Workbooks);
            dynamic book = com.Own(books.Add());
            dynamic sheet = com.Own(book.ActiveSheet);
            sheet.Name = "OfficeAssist検証";
            dynamic window = com.Own(app.ActiveWindow);
            window.WindowState = -4143;
            var area = Screen.PrimaryScreen!.WorkingArea;
            // Excelの座標はポイント。現在の画面から換算し右半分へ配置する。
            using var graphics = Graphics.FromHwnd((nint)(int)app.Hwnd);
            float factor = 72 / graphics.DpiX;
            app.Left = (area.Left + area.Width / 2) * factor;
            app.Top = area.Top * factor;
            app.Width = area.Width / 2 * factor;
            app.Height = area.Height * factor;
            window.Zoom = 100;
            dynamic cell = com.Own(sheet.Range["C5"]);
            cell.Value2 = "画像ドロップのテスト";
            cell.Select();
            Application.DoEvents();
            Thread.Sleep(700);
            nint handle = 0;
            Native.EnumChildWindows((nint)(int)app.Hwnd, (child, _) =>
            {
                if (Native.ClassName(child) == "EXCEL7" && Native.IsWindowVisible(child)) { handle = child; return false; }
                return true;
            }, 0);
            Check(handle != 0, "専用ExcelのEXCEL7を取得");
            Point point = FindCellPoint(window, handle, "$C$5");
            if (fixture)
            {
                var stop = Path.Combine(output, "stop-fixture");
                if (File.Exists(stop)) throw new InvalidOperationException("新しい実行識別子を指定してください。");
                File.WriteAllText(Path.Combine(output, "fixture.json"), System.Text.Json.JsonSerializer.Serialize(new { pid, hwnd = (long)handle, point.X, point.Y }));
                Console.WriteLine("FIXTURE READY " + output);
                var end = DateTime.UtcNow.AddMinutes(25);
                while (!File.Exists(stop) && DateTime.UtcNow < end) { Application.DoEvents(); Thread.Sleep(100); }
                dynamic shapes = com.Own(sheet.Shapes);
                File.WriteAllText(Path.Combine(output, "fixture-result.txt"), $"画像数: {shapes.Count}");
                book.SaveAs(Path.Combine(output, "OfficeAssist-UI検証.xlsx"), 51);
                book.Close(false);
                return;
            }
            Console.WriteLine(ExcelPictureInserter.Insert(handle, point, [image, image2]));
            dynamic inserted = com.Own(sheet.Shapes);
            Check((int)inserted.Count == 2, "実Excelへ2枚挿入");
            dynamic first = com.Own(inserted.Item(1));
            dynamic second = com.Own(inserted.Item(2));
            Check(Math.Abs((double)first.Left - (double)cell.Left) < 1 && Math.Abs((double)first.Top - (double)cell.Top) < 1, "ドロップしたC5の左上へ配置");
            Check((double)second.Top >= (double)first.Top + (double)first.Height + 11, "複数画像が縦に重ならず並ぶ");
            Check((int)first.Type == 13 && Math.Abs((double)first.Width / (double)first.Height - 2) < 0.01, "非リンクの画像と縦横比");
            Reject(() => ExcelPictureInserter.Insert(handle, point, [image, invalid]), "一括検証で壊れた画像を拒否");
            Check((int)inserted.Count == 2, "検証失敗時に既存画像を変更しない");
            sheet.Protect();
            Reject(() => ExcelPictureInserter.Insert(handle, point, [image]), "保護シートを拒否");
            Check((int)inserted.Count == 2, "保護シートの画像を保持");
            sheet.Unprotect();
            dynamic other = com.Own(books.Add());
            dynamic otherSheet = com.Own(other.ActiveSheet);
            dynamic otherShapes = com.Own(otherSheet.Shapes);
            window.Activate();
            Application.DoEvents();
            ExcelPictureInserter.Insert(handle, point, [image]);
            Check((int)inserted.Count == 3 && (int)otherShapes.Count == 0, "複数ブックで対象ウィンドウだけへ挿入");
            other.Close(false);
            book.SaveAs(Path.Combine(output, "OfficeAssist-COM検証.xlsx"), 51);
            book.Close(false);
        }
        finally { app.Quit(); }
    }

    static Point FindCellPoint(dynamic window, nint handle, string address)
    {
        Native.GetWindowRect(handle, out var rect);
        // PointsToScreenPixelsはこの環境の表示倍率でセル位置と一致しないため、
        // Excel自身のヒットテストで試験対象セルの画面座標を特定する。
        for (int y = rect.Top + 10; y < Math.Min(rect.Bottom, rect.Top + 400); y += 8)
        for (int x = rect.Left + 10; x < Math.Min(rect.Right, rect.Left + 500); x += 16)
        {
            object? value = window.RangeFromPoint(x, y);
            if (value is null) continue;
            try { if ((string)((dynamic)value).Address == address) return new Point(x, y); }
            catch (Exception ex) when (ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            finally { if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        }
        throw new Exception("試験対象セルの画面座標を取得できません。");
    }
}
