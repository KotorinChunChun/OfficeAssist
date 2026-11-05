namespace OfficeAssist;

internal sealed class ReceivedTabularBatch : IDisposable
{
    readonly Dictionary<string, FileStream> streams = new(StringComparer.OrdinalIgnoreCase);
    internal IReadOnlyList<string> Inputs { get; }
    internal bool HasFolders { get; }
    ReceivedTabularBatch(string[] inputs) { Inputs = inputs; HasFolders = inputs.Any(Directory.Exists); }
    internal static ReceivedTabularBatch Capture(string[] inputs)
    {
        if (!TabularInputDiscovery.Accepts(inputs)) throw new InvalidDataException("CSV・TSVまたはフォルダを指定してください。");
        var batch = new ReceivedTabularBatch(inputs.Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).ToArray());
        try
        {
            long totalBytes = 0;
            foreach (string path in batch.Inputs)
            {
                if (Directory.Exists(path)) { using var entries = Directory.EnumerateFileSystemEntries(path).GetEnumerator(); entries.MoveNext(); continue; }
                batch.streams[path] = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                if (batch.streams[path].Length is <= 0 or > 128 * 1024 * 1024) throw new InvalidDataException("CSV・TSVは空でない128MiB以下のファイルを指定してください。");
                if ((totalBytes += batch.streams[path].Length) > 512 * 1024 * 1024) throw new InvalidDataException("入力の合計容量は512MiBまでです。");
            }
            return batch;
        }
        catch { batch.Dispose(); throw; }
    }
    internal IReadOnlyList<TabularData> Read(IReadOnlyList<InputCandidate> selected, CancellationToken token)
    {
        if (selected.Count is 0 or > 50) throw new InvalidDataException("対象は1～50ファイルにしてください。");
        var opened = new List<FileStream>(); var tables = new List<TabularData>(); long totalBytes = 0, totalCells = 0;
        try
        {
            foreach (var file in selected)
            {
                token.ThrowIfCancellationRequested(); FileStream stream;
                if (!streams.TryGetValue(file.Path, out stream!))
                {
                    stream = new FileStream(file.Path, FileMode.Open, FileAccess.Read, FileShare.Read); opened.Add(stream);
                    if (stream.Length != file.Length || File.GetLastWriteTimeUtc(file.Path) != file.Modified) throw new IOException($"{file.Path}: 検索後にファイルが変わりました。対象を確認し直してください。");
                }
                if (stream.Length is <= 0 or > 128 * 1024 * 1024 || (totalBytes += stream.Length) > 512 * 1024 * 1024) throw new InvalidDataException("入力容量の上限（1ファイル128MiB・合計512MiB）を超えます。");
                var table = TabularFiles.ReadStream(stream, file.Path, token, file.Encoding); tables.Add(table);
                if ((totalCells += (long)table.RowCount * table.ColumnCount) > TabularFiles.MaximumCells) throw new InvalidDataException("入力は合計200万セルまでです。");
            }
            return tables;
        }
        catch { foreach (var table in tables) table.Rows.Dispose(); throw; }
        finally { foreach (var stream in opened) stream.Dispose(); }
    }
    public void Dispose() { foreach (var stream in streams.Values) stream.Dispose(); streams.Clear(); }
}
