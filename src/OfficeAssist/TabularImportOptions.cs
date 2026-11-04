using System.Globalization;

namespace OfficeAssist;

internal enum ImportDestination { NewBook, NewSheet, Selection }
internal enum ColumnType { TextFormat, TextQuote, Date, General, Number }
internal sealed record TabularImportOptions(ImportDestination Destination, ColumnType[][] Types, bool Header)
{
    internal static TabularImportOptions Infer(IReadOnlyList<TabularData> tables, ImportDestination destination = ImportDestination.Selection, bool? header = null)
    {
        bool hasHeader = header ?? tables.Any(ColumnConversion.LooksLikeHeader);
        return new(destination, tables.Select(t => Enumerable.Range(0, t.ColumnCount)
            .Select(c => ColumnConversion.Infer(t.Rows.Skip(hasHeader ? 1 : 0).Select(r => c < r.Length ? r[c] : ""))).ToArray()).ToArray(), hasHeader);
    }
}

internal static class ColumnConversion
{
    internal static readonly string[] Labels = ["文字列（書式設定）", "文字列（シングルクォート）", "日付", "標準", "数値"];
    static readonly string[] DateFormats = ["yyyy/M/d", "yyyy-MM-dd", "yyyy/M/d H:mm", "yyyy/M/d H:mm:ss", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd'T'HH:mm:ss"];

    internal static bool TryDate(string value, out DateTime date) => DateTime.TryParseExact(value, DateFormats,
        CultureInfo.InvariantCulture, DateTimeStyles.None, out date) && date.Year >= 1900 && date.Year <= 9999;
    internal static bool TryNumber(string value, out double number) => double.TryParse(value,
        NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number) && double.IsFinite(number);
    static bool PreserveAsText(string value)
    {
        string s = value.TrimStart('+', '-');
        return value != value.Trim() || (s.Length > 1 && s[0] == '0' && char.IsDigit(s[1])) || s.Count(char.IsDigit) > 15;
    }
    internal static ColumnType Infer(IEnumerable<string> items)
    {
        var values = items.Where(v => v.Length != 0).ToArray();
        if (values.Length == 0) return ColumnType.General;
        if (values.All(v => TryDate(v, out _))) return ColumnType.Date;
        if (values.All(v => !PreserveAsText(v) && TryNumber(v, out _))) return ColumnType.Number;
        return ColumnType.TextFormat;
    }
    internal static bool LooksLikeHeader(TabularData table)
    {
        if (table.RowCount < 2) return false;
        return Enumerable.Range(0, table.ColumnCount).Any(c =>
        {
            string first = c < table.Rows[0].Length ? table.Rows[0][c] : "";
            var type = Infer(table.Rows.Skip(1).Select(r => c < r.Length ? r[c] : ""));
            return first.Length != 0 && type is ColumnType.Number or ColumnType.Date && !TryNumber(first, out _) && !TryDate(first, out _);
        });
    }
    internal static object? Convert(string value, ColumnType type, bool header)
    {
        if (value.Length == 0) return null;
        // 元データ自身が引用符で始まる場合だけ、Excelの制御文字解釈をエスケープする。
        if (header || type == ColumnType.TextFormat) return value.StartsWith('\'') ? "'" + value : value;
        if (type == ColumnType.TextQuote) return "'" + value;
        if (type == ColumnType.Date && TryDate(value, out var date)) return date.ToOADate();
        if (type == ColumnType.Number && TryNumber(value, out var number)) return number;
        if (type == ColumnType.General)
        {
            if (TryDate(value, out date)) return date.ToOADate();
            if (!PreserveAsText(value) && TryNumber(value, out number)) return number;
            return value.StartsWith('\'') ? "'" + value : value;
        }
        throw new InvalidDataException(type == ColumnType.Date ? "日付へ変換できない値があります。" : "数値へ変換できない値があります。");
    }
    internal static string Format(ColumnType type) => type switch
    {
        ColumnType.TextFormat => "@", ColumnType.Date => "yyyy/mm/dd hh:mm:ss", ColumnType.Number => "0.###############", _ => "General"
    };
}
