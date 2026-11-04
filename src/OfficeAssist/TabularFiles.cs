using System.Text;

namespace OfficeAssist;

internal sealed record TabularData(string Name, string[][] Rows)
{
    internal int RowCount => Rows.Length;
    internal int ColumnCount => Rows.Max(r => r.Length);
    internal object?[,] ExcelValues()
    {
        var values = new object?[RowCount, ColumnCount];
        for (int r = 0; r < RowCount; r++)
            for (int c = 0; c < Rows[r].Length; c++)
                if (Rows[r][c].Length > 0) values[r, c] = "'" + Rows[r][c];
        return values;
    }
}

internal static class TabularFiles
{
    internal const int MaximumCells = 200_000;
    internal static bool Accepts(IReadOnlyList<string>? files) => files is { Count: > 0 and <= 50 }
        && files.All(f => !string.IsNullOrWhiteSpace(f) && Path.GetExtension(f).ToLowerInvariant() is ".csv" or ".tsv");

    internal static IReadOnlyList<TabularData> Read(IReadOnlyList<string> files)
    {
        if (!Accepts(files)) throw new InvalidDataException("CSV・TSVファイルだけを指定してください。");
        var tables = new List<TabularData>();
        long totalCells = 0;
        long totalBytes = 0;
        foreach (string file in files)
        {
            using var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (stream.Length is <= 0 or > 20 * 1024 * 1024) throw new InvalidDataException("CSV・TSVは１ファイル20MiBまでの空でないファイルを指定してください。");
            totalBytes += stream.Length;
            if (totalBytes > 64 * 1024 * 1024) throw new InvalidDataException("１回のファイル容量は合計64MiBまでです。");
            var bytes = new byte[(int)stream.Length];
            stream.ReadExactly(bytes);
            string content = Decode(bytes);
            var rows = Parse(content, Path.GetExtension(file).Equals(".tsv", StringComparison.OrdinalIgnoreCase) ? '\t' : ',');
            var table = new TabularData(Path.GetFileName(file), rows);
            totalCells += (long)table.RowCount * table.ColumnCount;
            if (totalCells > MaximumCells) throw new InvalidDataException("１回の展開は合計20万セルまでです。");
            tables.Add(table);
        }
        return tables;
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
        var rows = new List<string[]>();
        var fields = new List<string>();
        var field = new StringBuilder();
        bool quoted = false, closed = false;
        int cells = 0, maxColumns = 0;
        void EndField()
        {
            if (++cells > MaximumCells || fields.Count >= 16384) throw new InvalidDataException("列数またはセル数の上限を超えています。");
            fields.Add(field.ToString()); field.Clear(); closed = false;
        }
        void EndRow()
        {
            EndField(); maxColumns = Math.Max(maxColumns, fields.Count);
            rows.Add(fields.ToArray()); fields.Clear();
            if ((long)rows.Count * maxColumns > MaximumCells) throw new InvalidDataException("１回の展開は20万セルまでです。");
        }
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (quoted)
            {
                if (ch == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                    else { quoted = false; closed = true; }
                }
                else field.Append(ch);
            }
            else if (ch == separator) EndField();
            else if (ch is '\r' or '\n')
            {
                EndRow(); if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
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
        if (field.Length > 0 || fields.Count > 0 || closed || text[^1] is not ('\r' or '\n')) EndRow();
        return rows.ToArray();
    }
}
