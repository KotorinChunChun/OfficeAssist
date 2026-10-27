namespace OfficeAssist;

internal static class ImageFiles
{
    internal const int MaximumCount = 50;
    internal const long MaximumBytes = 50L * 1024 * 1024;
    static readonly HashSet<string> extensions = new(StringComparer.OrdinalIgnoreCase) { ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff" };

    // DragOverでは画像デコードやファイルI/Oを繰り返さない。
    internal static bool Accepts(IReadOnlyList<string>? paths) => paths is { Count: > 0 and <= MaximumCount }
        && paths.All(path => !string.IsNullOrWhiteSpace(path) && extensions.Contains(Path.GetExtension(path)));

    internal static void Validate(IReadOnlyList<string> paths)
    {
        if (!Accepts(paths)) throw new InvalidDataException("PNG・JPEG・BMP・GIF・TIFFを1～50件指定してください。");
        foreach (var path in paths)
        {
            var file = new FileInfo(path);
            if (!file.Exists || file.Length is <= 0 or > MaximumBytes)
                throw new InvalidDataException($"読み込めない画像、または50MiBを超える画像です: {Path.GetFileName(path)}");
            try
            {
                using var stream = File.OpenRead(path);
                using var image = Image.FromStream(stream, false, true);
                if ((long)image.Width * image.Height > 100_000_000)
                    throw new InvalidDataException($"画像は100メガピクセルまでです: {file.Name}");
            }
            catch (Exception ex) when (ex is ArgumentException or OutOfMemoryException)
            {
                throw new InvalidDataException($"画像として読み込めません: {file.Name}", ex);
            }
        }
    }

    internal static SizeF Fit(float width, float height)
    {
        if (width <= 0 || height <= 0 || !float.IsFinite(width) || !float.IsFinite(height))
            throw new InvalidDataException("画像の寸法が不正です。");
        float ratio = Math.Min(1, 480 / Math.Max(width, height));
        return new SizeF(width * ratio, height * ratio);
    }
}
