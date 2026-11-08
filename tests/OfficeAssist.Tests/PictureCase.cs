using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeAssist;

internal static partial class Program
{
    sealed record PictureCase(string Workbook, string Sheet, string[] Images, string Destination, string Selection, string? DropCell);
    readonly record struct PictureBounds(double Left, double Top, double Width, double Height)
    {
        internal bool Near(PictureBounds other) => Math.Abs(Left - other.Left) <= 0.05 && Math.Abs(Top - other.Top) <= 0.05
            && Math.Abs(Width - other.Width) <= 0.05 && Math.Abs(Height - other.Height) <= 0.05;
    }
    sealed record PictureShape(string Name, int Type, PictureBounds Bounds, int Placement, int LockAspectRatio);

    static int VerifyPictureCase(string[] args, string output)
    {
        if (args is not ["--excel", "--picture-case", var manifest, _])
            throw new ArgumentException("--excel --picture-case ケース.json 出力フォルダー の順で指定してください。");
        manifest = Path.GetFullPath(manifest);
        var test = JsonSerializer.Deserialize<PictureCase>(File.ReadAllText(manifest)) ?? throw new InvalidDataException("試験ケースが空です。");
        if (string.IsNullOrWhiteSpace(test.Workbook) || string.IsNullOrWhiteSpace(test.Sheet)
            || string.IsNullOrWhiteSpace(test.Selection) || test.Images is null || test.Images.Any(string.IsNullOrWhiteSpace))
            throw new ArgumentException("Workbook、Sheet、Selection、Imagesを指定してください。");
        string InputPath(string path) => Path.GetFullPath(path, Path.GetDirectoryName(manifest)!);
        string source = InputPath(test.Workbook);
        if (!Path.GetExtension(source).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("複製元は試験用の.xlsxを指定してください。");
        string selectionAddress = ExcelImportEngine.NormalizeAddress(test.Selection);
        string dropCell = ExcelImportEngine.NormalizeAddress(test.DropCell ?? "A1");
        if (dropCell.Contains(':') || dropCell.Contains(',')) throw new ArgumentException("DropCellは1セルを指定してください。");
        var destination = test.Destination switch
        {
            "drop" => Destination.DropPosition,
            "selection-position" => Destination.SelectionPosition,
            "selection-range" => Destination.SelectionRange,
            _ => throw new ArgumentException("Destinationはdrop、selection-position、selection-rangeのいずれかです。")
        };
        if (destination == Destination.DropPosition && test.DropCell is null) throw new ArgumentException("dropにはDropCellが必要です。");
        var images = test.Images.Select(InputPath).ToArray();
        ImageFiles.Validate(images);
        string resultPath = Path.Combine(output, "picture-result.xlsx"), reportPath = Path.Combine(output, "picture-report.json");
        if (File.Exists(resultPath) || File.Exists(reportPath)) throw new IOException("別の出力フォルダーを指定してください。");
        string sourceHash = PictureFileHash(source);
        string working = Path.Combine(output, $"picture-working-{Guid.NewGuid():N}.xlsx");
        File.Copy(source, working);
        try
        {
            object report = VerifyPictureWorkbook(test, selectionAddress, dropCell, images, destination, working, source, sourceHash);
            File.Move(working, resultPath);
            using var result = new FileStream(reportPath, FileMode.CreateNew);
            JsonSerializer.Serialize(result, report, new JsonSerializerOptions { WriteIndented = true });
            Console.WriteLine($"PASS: {count}件。結果: {reportPath}");
            return 0;
        }
        finally { if (File.Exists(working)) File.Delete(working); }
    }

    static object VerifyPictureWorkbook(PictureCase test, string selectionAddress, string dropCell, string[] images,
        Destination destination, string working, string source, string sourceHash)
    {
        var existing = Process.GetProcessesByName("EXCEL").Select(p => { using (p) return p.Id; }).ToHashSet();
        using var com = new ComScope();
        dynamic app = com.Own(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!);
        Native.GetWindowThreadProcessId((nint)(int)app.Hwnd, out uint pid);
        if (pid == 0 || existing.Contains((int)pid)) throw new InvalidOperationException("試験専用のExcelを確保できません。");
        try
        {
            app.Visible = false;
            app.DisplayAlerts = false;
            app.EnableEvents = false;
            app.AutomationSecurity = 3;
            dynamic books = com.Own(app.Workbooks);
            PictureBounds[] original = ReadNaturalPictureSizes(com, books, images);
            dynamic book = com.Own(books.Open(working, UpdateLinks: 0, ReadOnly: false, AddToMru: false));
            PictureShape[] inserted;
            PictureShape[] after;
            PictureBounds[] expected;
            string? warning;
            try
            {
                dynamic sheet = com.Own(book.Worksheets.Item(test.Sheet));
                sheet.Activate();
                dynamic window = com.Own(app.ActiveWindow);
                dynamic selected = com.Own(sheet.Range[selectionAddress]), anchor = com.Own(sheet.Range[dropCell]);
                selected.Select();
                PictureShape[] before = ReadPictureShapes(com, sheet);
                expected = ExpectedPictureBounds(com, selected, anchor, destination, original);
                var options = new InsertOptions(destination, PictureKind.Embedded, []);
                ExcelPictureInserter.Result insertion = ExcelPictureInserter.InsertResolved(com, window, sheet, anchor, selected, images, options);
                warning = insertion.Warning;
                after = ReadPictureShapes(com, sheet);
                Check(after.Length == before.Length + images.Length && after.Take(before.Length).SequenceEqual(before), "既存の図形を保持し指定画像だけを追加");
                inserted = after.Skip(before.Length).ToArray();
                for (int i = 0; i < inserted.Length; i++)
                    Check(inserted[i].Bounds.Near(expected[i]) && inserted[i].Type == 13 && inserted[i].Placement == 3 && inserted[i].LockAspectRatio == -1,
                        $"画像{i + 1}の位置・サイズ・埋め込み・縦横比固定を照合");
                book.Save();
            }
            finally { book.Close(false); }

            dynamic saved = com.Own(books.Open(working, UpdateLinks: 0, ReadOnly: true, AddToMru: false));
            try
            {
                dynamic sheet = com.Own(saved.Worksheets.Item(test.Sheet));
                PictureShape[] reopened = ReadPictureShapes(com, sheet);
                Check(reopened.Length == after.Length && reopened.Zip(after).All(pair =>
                    pair.First.Bounds.Near(pair.Second.Bounds) && pair.First with { Bounds = pair.Second.Bounds } == pair.Second),
                    "保存した複製ブックを開き直して画像配置を照合");
            }
            finally { saved.Close(false); }
            Check(PictureFileHash(source) == sourceHash, "複製元のSHA-256が一致");
            return new { verified = true, pid, test.Sheet, test.Destination, selection = selectionAddress, test.DropCell, sourceSha256 = sourceHash,
                resultSha256 = PictureFileHash(working), warning, pictures = inserted.Select((picture, i) => new { image = Path.GetFileName(images[i]), expected = expected[i], actual = picture }) };
        }
        finally { app.Quit(); }
    }

    static PictureBounds[] ReadNaturalPictureSizes(ComScope com, dynamic books, string[] images)
    {
        dynamic scratch = com.Own(books.Add(-4167));
        try
        {
            dynamic sheet = com.Own(scratch.Worksheets.Item(1)), shapes = com.Own(sheet.Shapes);
            return images.Select(path =>
            {
                dynamic picture = com.Own(shapes.AddPicture(path, 0, -1, 0, 0, -1, -1));
                return new PictureBounds(0, 0, (double)picture.Width, (double)picture.Height);
            }).ToArray();
        }
        finally { scratch.Close(false); }
    }

    static PictureBounds[] ExpectedPictureBounds(ComScope com, dynamic selected, dynamic anchor, Destination mode, PictureBounds[] originals)
    {
        dynamic areas = com.Own(selected.Areas);
        int areaCount = mode == Destination.DropPosition ? 1 : (int)areas.Count;
        var result = new PictureBounds[originals.Length];
        int index = 0;
        for (int a = 1; a <= areaCount && index < originals.Length; a++)
        {
            dynamic range = mode == Destination.DropPosition ? anchor : com.Own(areas.Item(a));
            int count = a == areaCount ? originals.Length - index : 1;
            double top = (double)range.Top, rowHeight = (double)range.Height / count;
            for (int i = 0; i < count; i++, index++)
            {
                var original = originals[index];
                double scale = mode == Destination.SelectionRange ? Math.Min((double)range.Width / original.Width, rowHeight / original.Height) : 1;
                result[index] = new((double)range.Left, top, original.Width * scale, original.Height * scale);
                top += mode == Destination.SelectionRange ? rowHeight : original.Height + 12;
            }
        }
        return result;
    }

    static PictureShape[] ReadPictureShapes(ComScope com, dynamic sheet)
    {
        dynamic shapes = com.Own(sheet.Shapes);
        return Enumerable.Range(1, (int)shapes.Count).Select(i =>
        {
            dynamic shape = com.Own(shapes.Item(i));
            return new PictureShape((string)shape.Name, (int)shape.Type, new((double)shape.Left, (double)shape.Top, (double)shape.Width, (double)shape.Height),
                (int)shape.Placement, (int)shape.LockAspectRatio);
        }).ToArray();
    }

    static string PictureFileHash(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}
