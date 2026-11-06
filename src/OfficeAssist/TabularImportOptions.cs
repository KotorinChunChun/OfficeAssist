using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeAssist;

internal enum ImportDestination { NewBook, NewSheet, Selection, AppendSameSheet }
internal enum ColumnType { TextFormat, TextQuote, Date, General, Number, DateWeekday, DateTime, Time, GroupedNumber, Percent, Yen, Duration }
internal sealed record TabularImportOptions(ImportDestination Destination, ColumnType[][] Types, bool Header)
{
    internal int[][]? DecimalPlaces { get; init; }
    internal string? BookId { get; init; }
    internal string? SheetName { get; init; }
    internal string? CellAddress { get; init; }
    internal bool Merged { get; init; }
    internal CancellationToken Cancellation { get; init; }
    internal Action<string>? Progress { get; init; }
    internal Action<IReadOnlyList<ExcelImportEngine.Write>>? Completed { get; init; }
    internal int Decimals(int file, int column)
    {
        if (DecimalPlaces is null) return 0;
        if (DecimalPlaces.Length != Types.Length || DecimalPlaces[file].Length != Types[file].Length || DecimalPlaces[file][column] is < 0 or > 15)
            throw new InvalidDataException("小数点以下桁数は0～15桁で指定してください。");
        return DecimalPlaces[file][column];
    }
    internal static TabularImportOptions Infer(IReadOnlyList<TabularData> tables, ImportDestination destination = ImportDestination.Selection, bool? header = null, CancellationToken token = default)
    {
        bool hasHeader = header ?? tables.Any(t => ColumnConversion.LooksLikeHeader(t, token));
        var inferred = tables.Select(t => ColumnInference.Scan(t, hasHeader, token)).ToArray();
        return new(destination, inferred.Select(i => i.Types).ToArray(), hasHeader) { DecimalPlaces = inferred.Select(i => i.Decimals).ToArray() };
    }
}

internal static class ColumnConversion
{
    internal static readonly string[] Labels = ["文字列（書式設定）", "文字列（シングルクォート）", "日付", "標準", "数値", "日付と曜日", "日付と時刻", "時刻", "桁区切り数値", "パーセント", "円", "経過時間"];
    static readonly string[] DateFormats = ["yyyy/M/d", "yyyy-MM-dd", "yyyy/M/d H:mm", "yyyy/M/d H:mm:ss", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd'T'HH:mm:ss"];

    internal static bool TryDate(string value, out DateTime date) => DateTime.TryParseExact(value, DateFormats,
        CultureInfo.InvariantCulture, DateTimeStyles.None, out date) && date.Year >= 1900 && date.Year <= 9999;
    internal static bool TryNumber(string value, out double number) => double.TryParse(value,
        NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number) && double.IsFinite(number);
    internal static bool IsNumeric(ColumnType type) => type is ColumnType.Number or ColumnType.GroupedNumber or ColumnType.Percent or ColumnType.Yen;
    internal static bool IsDate(ColumnType type) => type is ColumnType.Date or ColumnType.DateWeekday or ColumnType.DateTime;
    internal static bool TryYen(string value, out double number) => TryNumber(value.Replace("¥", "").Replace("￥", "").Replace("\\", ""), out number);
    internal static bool TryPercent(string value, out double number)
    {
        bool percent = value.EndsWith('%') || value.EndsWith('％');
        if (!TryNumber(percent ? value[..^1] : value, out number)) return false;
        if (percent) number /= 100;
        return double.IsFinite(number);
    }
    internal static bool TryDuration(string value, out double days)
    {
        days = 0;
        var match = Regex.Match(value, @"^(?<sign>-?)(?<hours>\d+):(?<minutes>[0-5]\d):(?<seconds>[0-5]\d(?:\.\d+)?)$", RegexOptions.CultureInvariant);
        if (!match.Success || !TryNumber(match.Groups["hours"].Value, out double hours) || !TryNumber(match.Groups["seconds"].Value, out double seconds)) return false;
        days = (hours * 3600 + int.Parse(match.Groups["minutes"].Value, CultureInfo.InvariantCulture) * 60 + seconds) / 86400;
        if (match.Groups["sign"].Value == "-") days = -days;
        return double.IsFinite(days);
    }
    internal static int InferDecimals(IEnumerable<string> items, ColumnType type)
    {
        if (!IsNumeric(type)) return 0;
        int places = 0;
        foreach (string value in items.Where(v => v.Length != 0))
        {
            double number;
            bool valid = type == ColumnType.Percent ? TryPercent(value, out number) : type == ColumnType.Yen ? TryYen(value, out number) : TryNumber(value, out number);
            if (!valid) continue;
            if (type == ColumnType.Percent) number *= 100;
            places = Math.Max(places, CountDecimals(number));
        }
        return Math.Min(places, 15);
    }
    internal static int CountDecimals(double number)
    {
        if (!double.IsFinite(number)) return 0;
        string text = number.ToString("0.###############", CultureInfo.InvariantCulture); int point = text.IndexOf('.');
        return point < 0 ? 0 : Math.Min(text.Length - point - 1, 15);
    }
    internal static bool PreserveAsText(string value)
    {
        string s = value.TrimStart('+', '-');
        return value != value.Trim() || (s.Length > 1 && s[0] == '0' && char.IsDigit(s[1])) || s.Count(char.IsDigit) > 15;
    }
    internal static ColumnType Infer(IEnumerable<string> items)
    {
        var values = items.Where(v => v.Length != 0).ToArray();
        if (values.Length == 0) return ColumnType.General;
        if (values.All(v => TryDate(v, out _))) return values.Any(v => TryDate(v, out var date) && date.TimeOfDay != TimeSpan.Zero) ? ColumnType.DateTime : ColumnType.Date;
        if (values.All(v => TryDuration(v, out _))) return values.Any(v => TryDuration(v, out var days) && (days < 0 || days >= 1)) ? ColumnType.Duration : ColumnType.Time;
        if (values.All(v => TryPercent(v, out _)) && values.Any(v => v.EndsWith('%') || v.EndsWith('％'))) return ColumnType.Percent;
        if (values.All(v => TryYen(v, out _)) && values.Any(v => v.Contains('¥') || v.Contains('￥') || v.Contains('\\'))) return ColumnType.Yen;
        if (values.All(v => !PreserveAsText(v) && TryNumber(v, out _))) return values.Any(v => v.Contains(',')) ? ColumnType.GroupedNumber : ColumnType.Number;
        return ColumnType.TextFormat;
    }
    internal static bool LooksLikeHeader(TabularData table, CancellationToken token = default)
    {
        if (table.RowCount < 2) return false;
        var types = ColumnInference.Scan(table, true, token).Types;
        return Enumerable.Range(0, table.ColumnCount).Any(c =>
        {
            string first = c < table.Rows[0].Length ? table.Rows[0][c] : "";
            var type = types[c];
            if (first.Length == 0 || type is ColumnType.General or ColumnType.TextFormat or ColumnType.TextQuote) return false;
            try { Convert(first, type, false); return false; } catch (InvalidDataException) { return true; }
        });
    }
    internal static object? Convert(string value, ColumnType type, bool header)
    {
        if (TryConvert(value, type, header, out var result)) return result;
        throw new InvalidDataException($"{Labels[(int)type]}へ変換できない値があります。");
    }
    // 検査で不正値ごとに例外を生成しない。実書込みもこの判定を通る。
    internal static bool TryConvert(string value, ColumnType type, bool header, out object? result, bool materialize = true)
    {
        result = null;
        if (value.Length == 0) return true;
        // 元データ自身が引用符で始まる場合だけ、Excelの制御文字解釈をエスケープする。
        if (header || type == ColumnType.TextFormat) { if (materialize) result = value.StartsWith('\'') ? "'" + value : value; return true; }
        if (type == ColumnType.TextQuote) { if (materialize) result = "'" + value; return true; }
        if (IsDate(type) && TryDate(value, out var date)) { if (materialize) result = date.ToOADate(); return true; }
        if (type == ColumnType.Time)
        {
            if (TryDuration(value, out var time) && time >= 0 && time < 1) { if (materialize) result = time; return true; }
            if (TryDate(value, out date)) { if (materialize) result = date.ToOADate(); return true; }
        }
        if (type == ColumnType.Duration && TryDuration(value, out var duration)) { if (materialize) result = duration; return true; }
        if (type is ColumnType.Number or ColumnType.GroupedNumber && TryNumber(value, out var number)) { if (materialize) result = number; return true; }
        if (type == ColumnType.Percent && TryPercent(value, out var percent)) { if (materialize) result = percent; return true; }
        if (type == ColumnType.Yen && TryYen(value, out var yen)) { if (materialize) result = yen; return true; }
        if (type == ColumnType.General)
        {
            if (!materialize) return true;
            if (TryDate(value, out date)) { result = date.ToOADate(); return true; }
            if (!PreserveAsText(value) && TryNumber(value, out number)) { result = number; return true; }
            result = value.StartsWith('\'') ? "'" + value : value; return true;
        }
        return false;
    }
    internal static string Format(ColumnType type, int decimals = 0)
    {
        if (decimals is < 0 or > 15) throw new ArgumentOutOfRangeException(nameof(decimals));
        string fraction = decimals == 0 ? "" : "." + new string('0', decimals);
        return type switch
        {
            ColumnType.TextFormat => "@", ColumnType.Date => "yyyy/mm/dd", ColumnType.DateWeekday => "[$-ja-JP]yyyy/mm/dd (ddd)",
            ColumnType.DateTime => "yyyy/mm/dd hh:mm:ss", ColumnType.Time => "hh:mm:ss", ColumnType.Duration => "[h]:mm:ss",
            ColumnType.Number => "0" + fraction, ColumnType.GroupedNumber => "#,##0" + fraction,
            ColumnType.Percent => "0" + fraction + "%", ColumnType.Yen => "\"¥\"#,##0" + fraction, _ => "General"
        };
    }
}
