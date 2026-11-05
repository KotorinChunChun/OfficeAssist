using System.Text;

namespace OfficeAssist;

internal sealed record TabularData(string Name, TabularRows Rows)
{
    internal TabularData(string name, string[][] rows) : this(name, TabularRows.Memory(rows)) { }
    internal int RowCount => Rows.Length;
    internal int ColumnCount => Rows.Width;
    internal Func<int, string>? Origin { get; init; }
    internal string? SourceSha256 { get; init; }
    internal string? SourcePath { get; init; }
    internal string? DetectedEncoding { get; init; }
    internal object?[,] ExcelValues(ColumnType[] types, bool header, int start = 0, int? count = null)
    {
        if (types.Length != ColumnCount || types.Any(t => !Enum.IsDefined(t))) throw new InvalidDataException("列形式の指定が不正です。");
        int length = count ?? RowCount - start;
        var values = new object?[length, ColumnCount];
        for (int r = start; r < start + length; r++)
        {
            var row = Rows[r];
            for (int c = 0; c < row.Length; c++)
                try { values[r - start, c] = ColumnConversion.Convert(row[c], types[c], header && r == 0); }
                catch (InvalidDataException ex) { throw new InvalidDataException($"{Origin?.Invoke(r) ?? $"{Name}: 行{r + 1}"}・列{c + 1} — {ex.Message}", ex); }
        }
        return values;
    }
}

internal static class TabularFiles
{
    internal const int MaximumCells = 2_000_000;
    internal static bool Accepts(IReadOnlyList<string>? files) => files is { Count: > 0 and <= 50 }
        && files.All(f => !string.IsNullOrWhiteSpace(f) && Path.GetExtension(f).ToLowerInvariant() is ".csv" or ".tsv");

    internal static IReadOnlyList<TabularData> Read(IReadOnlyList<string> files, CancellationToken token = default, TableEncoding encoding = TableEncoding.Auto, IReadOnlyDictionary<string, TableEncoding>? encodings = null)
    {
        if (!Accepts(files)) throw new InvalidDataException("CSV・TSVファイルだけを指定してください。");
        var tables = new List<TabularData>();
        long totalCells = 0;
        long totalBytes = 0;
        try { foreach (string file in files)
        {
            using var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (stream.Length is <= 0 or > 128 * 1024 * 1024) throw new InvalidDataException("CSV・TSVは１ファイル128MiBまでの空でないファイルを指定してください。");
            totalBytes += stream.Length;
            if (totalBytes > 512 * 1024 * 1024) throw new InvalidDataException("１回のファイル容量は合計512MiBまでです。");
            var table = ReadStream(stream, file, token, encodings is not null && encodings.TryGetValue(Path.GetFullPath(file), out var selected) ? selected : encoding);
            totalCells += (long)table.RowCount * table.ColumnCount;
            tables.Add(table);
            if (totalCells > MaximumCells) throw new InvalidDataException("１回の展開は合計200万セルまでです。");
        } }
        catch { foreach (var table in tables) table.Rows.Dispose(); throw; }
        return tables;
    }

    internal static TabularData ReadStream(FileStream stream, string file, CancellationToken token, TableEncoding choice = TableEncoding.Auto)
    {
        stream.Position = 0; token.ThrowIfCancellationRequested();
        string digest = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream)).ToLowerInvariant();
        stream.Position = 0; Encoding encoding = choice == TableEncoding.Auto ? DetectEncoding(stream, token) : TabularEncoding.Explicit(stream, choice); stream.Position = 0;
        using var reader = new StreamReader(stream, encoding, false, 65536, true);
        var positions = new List<(int Start, int End)>();
        TabularRows rows;
        try { rows = TabularRows.Store(ParseRecords(reader, Path.GetExtension(file).Equals(".tsv", StringComparison.OrdinalIgnoreCase) ? '\t' : ',', token,
            (_, first, last) => positions.Add((first, last))), token); }
        catch (DecoderFallbackException ex) { throw new InvalidDataException($"{Path.GetFileName(file)}: 選択した文字コードで読み込めません。文字コードを変更してください。", ex); }
        return new(Path.GetFileName(file), rows) { SourcePath = Path.GetFullPath(file), DetectedEncoding = encoding.WebName, SourceSha256 = digest, Origin = r => $"{Path.GetFullPath(file)}: 行{r + 1}（物理{positions[r].Start}～{positions[r].End}行）" };
    }

    static Encoding DetectEncoding(Stream stream, CancellationToken token)
    {
        byte[] prefix = new byte[4]; int n = stream.Read(prefix);
        if (n >= 4 && (prefix.AsSpan().SequenceEqual(new byte[] {255,254,0,0}) || prefix.AsSpan().SequenceEqual(new byte[] {0,0,254,255}))) throw new InvalidDataException("UTF-32は対象外です。");
        if (n >= 2 && prefix[0] == 255 && prefix[1] == 254) return new UnicodeEncoding(false, true, true);
        if (n >= 2 && prefix[0] == 254 && prefix[1] == 255) return new UnicodeEncoding(true, true, true);
        if (n >= 3 && prefix[0] == 239 && prefix[1] == 187 && prefix[2] == 191) return new UTF8Encoding(true, true);
        stream.Position = 0;
        try { using var reader = new StreamReader(stream, new UTF8Encoding(false, true), false, 65536, true); char[] buffer = new char[32768]; while (reader.Read(buffer) != 0) token.ThrowIfCancellationRequested(); return new UTF8Encoding(false, true); }
        catch (DecoderFallbackException) { Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); return Encoding.GetEncoding(932, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback); }
    }

    internal static string Decode(byte[] bytes)
    {
        if (bytes.AsSpan().StartsWith(new byte[] { 0xFF, 0xFE, 0, 0 }) || bytes.AsSpan().StartsWith(new byte[] { 0, 0, 0xFE, 0xFF }))
            throw new InvalidDataException("UTF-32は対象外です。UTF-8で保存してください。");
        if (bytes.AsSpan().StartsWith(new byte[] { 0xEF, 0xBB, 0xBF })) return new UTF8Encoding(false, true).GetString(bytes, 3, bytes.Length - 3);
        if (bytes.AsSpan().StartsWith(new byte[] { 0xFF, 0xFE })) return new UnicodeEncoding(false, false, true).GetString(bytes, 2, bytes.Length - 2);
        if (bytes.AsSpan().StartsWith(new byte[] { 0xFE, 0xFF })) return new UnicodeEncoding(true, false, true).GetString(bytes, 2, bytes.Length - 2);
        try { return new UTF8Encoding(false, true).GetString(bytes); }
        catch (DecoderFallbackException)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return Encoding.GetEncoding(932, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback).GetString(bytes);
        }
    }

    internal static string[][] Parse(string text, char separator)
    {
        if (text.Length == 0 || text.Contains('\0')) throw new InvalidDataException("空データまたはバイナリデータは展開できません。");
        return ParseRecords(new StringReader(text), separator).ToArray();
    }
    internal static IEnumerable<string[]> ParseRecords(TextReader reader, char separator, CancellationToken token = default, Action<int, int, int>? position = null)
    {
        var fields = new List<string>();
        var field = new StringBuilder();
        bool quoted = false, closed = false;
        int cells = 0, maxColumns = 0;
        void EndField()
        {
            if (++cells > MaximumCells || fields.Count >= 16384) throw new InvalidDataException("列数またはセル数の上限を超えています。");
            fields.Add(field.ToString()); field.Clear(); closed = false;
        }
        int rowCount = 0; bool pending = false;
        int physical = 1, recordStart = 1; char previous = '\0';
        for (int item; (item = reader.Read()) >= 0;)
        {
            char ch = (char)item; pending = true;
            int lineBefore = physical;
            if (ch == '\r' || ch == '\n' && previous != '\r') physical++;
            previous = ch;
            if ((cells & 1023) == 0) token.ThrowIfCancellationRequested();
            if (ch == '\0') throw new InvalidDataException("バイナリデータは展開できません。");
            if (quoted)
            {
                if (ch == '"')
                {
                    if (reader.Peek() == '"') { field.Append('"'); reader.Read(); }
                    else { quoted = false; closed = true; }
                }
                else field.Append(ch);
            }
            else if (ch == separator) EndField();
            else if (ch is '\r' or '\n')
            {
                EndField(); maxColumns = Math.Max(maxColumns, fields.Count);
                if ((long)++rowCount * maxColumns > MaximumCells) throw new InvalidDataException("１回の展開は200万セルまでです。");
                position?.Invoke(rowCount, recordStart, lineBefore);
                yield return fields.ToArray(); fields.Clear(); pending = false; recordStart = physical;
                if (ch == '\r' && reader.Peek() == '\n') { reader.Read(); previous = '\n'; }
            }
            else if (closed) throw new InvalidDataException("閉じ引用符の後に不正な文字があります。");
            else if (ch == '"')
            {
                if (field.Length != 0) throw new InvalidDataException("引用符は項目の先頭から指定してください。");
                quoted = true;
            }
            else field.Append(ch);
            if (field.Length > 32767) throw new InvalidDataException("１セルの文字数は32767文字までです。");
        }
        if (quoted) throw new InvalidDataException("引用符が閉じられていません。");
        if (pending) { EndField(); if ((long)(rowCount + 1) * Math.Max(maxColumns, fields.Count) > MaximumCells) throw new InvalidDataException("１回の展開は200万セルまでです。"); position?.Invoke(rowCount + 1, recordStart, physical); yield return fields.ToArray(); }
    }
}
