namespace OfficeAssist;

// Copy応答前に受信画像を所有する。送信元の一時ファイルの寿命へ依存させない。
internal sealed class ReceivedImageBatch : IDisposable
{
    readonly string directory;
    readonly List<string> staged = [];
    internal IReadOnlyList<string> Originals { get; }
    internal IReadOnlyList<string> Images => staged.AsReadOnly();

    ReceivedImageBatch(string directory, string[] originals)
    { this.directory = directory; Originals = Array.AsReadOnly(originals); }

    internal static ReceivedImageBatch Capture(IReadOnlyList<string> files, string? root = null)
    {
        if (!ImageFiles.Accepts(files)) throw new InvalidDataException("対応する画像ファイルを指定してください。");
        root ??= Path.Combine(Path.GetTempPath(), "OfficeAssist", "received-images");
        var batch = new ReceivedImageBatch(Path.Combine(Path.GetFullPath(root), Guid.NewGuid().ToString("N")), files.ToArray());
        try
        {
            Directory.CreateDirectory(batch.directory);
            for (int i = 0; i < batch.Originals.Count; i++)
            {
                string original = batch.Originals[i];
                // 読取中の削除・書換えを許可せず、開いた実ファイルの長さを検証する。
                using var source = new FileStream(original, FileMode.Open, FileAccess.Read, FileShare.Read);
                if (source.Length is <= 0 or > ImageFiles.MaximumBytes) throw new InvalidDataException("画像は1ファイル50MiBまでです。");
                string copy = Path.Combine(batch.directory, i.ToString("D2") + Path.GetExtension(original));
                using var destination = new FileStream(copy, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                batch.staged.Add(copy);
                source.CopyTo(destination);
            }
            ImageFiles.Validate(batch.staged);
            return batch;
        }
        catch { batch.Dispose(); throw; }
    }

    public void Dispose()
    {
        // 自分が生成したファイルだけ。元画像・他セッション・親フォルダーを削除しない。
        foreach (string path in staged)
        {
            try { File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
        try { Directory.Delete(directory, false); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
