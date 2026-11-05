using System.Collections;
using System.Text;

namespace OfficeAssist;

// 行データの実体はディスクに保持し、画面・変換は同じ読み取り専用ビューを参照する。
internal sealed class TabularRows : IReadOnlyList<string[]>, IDisposable
{
    readonly Func<int, string[]> read;
    readonly Action? release;
    internal int Length { get; }
    public int Count => Length;
    internal int Width { get; }
    public string[] this[int index] => index >= 0 && index < Length ? read(index) : throw new ArgumentOutOfRangeException(nameof(index));
    internal TabularRows(int count, int width, Func<int, string[]> read, Action? release = null)
    { Length = count; Width = width; this.read = read; this.release = release; }
    internal static TabularRows Memory(string[][] rows) => new(rows.Length, rows.Length == 0 ? 0 : rows.Max(r => r.Length), i => rows[i]);
    public IEnumerator<string[]> GetEnumerator() { for (int i = 0; i < Length; i++) yield return this[i]; }
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    public void Dispose() => release?.Invoke();

    internal static TabularRows Store(IEnumerable<string[]> rows, CancellationToken token)
    {
        string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAssist", "table-cache");
        Directory.CreateDirectory(directory);
        // DeleteOnCloseによりキャンセル・正常終了・プロセス終了で自分のファイルだけを回収する。
        var stream = new FileStream(Path.Combine(directory, Guid.NewGuid().ToString("N") + ".bin"), FileMode.CreateNew,
            FileAccess.ReadWrite, FileShare.None, 65536, FileOptions.DeleteOnClose | FileOptions.RandomAccess);
        try
        {
            var offsets = new List<long>(); int width = 0;
            using (var writer = new BinaryWriter(stream, Encoding.UTF8, true))
                foreach (var row in rows)
                {
                    token.ThrowIfCancellationRequested(); offsets.Add(stream.Position); width = Math.Max(width, row.Length);
                    if ((long)offsets.Count * width > TabularFiles.MaximumCells) throw new InvalidDataException("展開は合計200万セルまでです。");
                    writer.Write(row.Length); foreach (string value in row) writer.Write(value);
                    if (stream.Position > 2L * 1024 * 1024 * 1024) throw new InvalidDataException("一時データは2GiBまでです。");
                }
            var reader = new BinaryReader(stream, Encoding.UTF8, true);
            int cached = -1; string[] last = [];
            return new(offsets.Count, width, index =>
            {
                lock (stream)
                {
                    if (cached == index) return last;
                    stream.Position = offsets[index]; var row = new string[reader.ReadInt32()];
                    for (int c = 0; c < row.Length; c++) row[c] = reader.ReadString();
                    cached = index; last = row; return row;
                }
            }, () => { lock (stream) { reader.Dispose(); stream.Dispose(); } });
        }
        catch { stream.Dispose(); throw; }
    }
}
