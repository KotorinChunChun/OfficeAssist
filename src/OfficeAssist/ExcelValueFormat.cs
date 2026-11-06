namespace OfficeAssist;

// 実書込みとプレビューで、Excelへの書式設定の方法を共通化する。
internal static class ExcelValueFormat
{
    internal static string Code(ColumnType type, int decimals) => ColumnConversion.Format(type, decimals);
    internal static string CellCode(ColumnType type, int decimals, object? value) => value is string && type is not (ColumnType.TextFormat or ColumnType.TextQuote) ? "@" : Code(type, decimals);
    internal static void Apply(dynamic app, dynamic range, string code)
    {
        if (code == "General") range.NumberFormatLocal = (string)app.International[26];
        else range.NumberFormat = code;
    }
    internal static object? Value(string raw, ColumnType type, bool header, bool date1904)
    {
        object? value = ColumnConversion.Convert(raw, type, header);
        return AdjustSerial(value, type, date1904);
    }
    internal static object? AdjustSerial(object? value, ColumnType type, bool date1904)
    {
        if (date1904 && value is double number && (ColumnConversion.IsDate(type) || type == ColumnType.Time && number >= 1)) return number - 1462;
        return value;
    }
}
