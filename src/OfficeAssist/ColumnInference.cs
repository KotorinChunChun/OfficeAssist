namespace OfficeAssist;

internal static class ColumnInference
{
    sealed class Stats
    {
        bool any, date = true, time = true, number = true, percent = true, yen = true;
        bool hasDateTime, duration, hasPercent, hasYen, grouped;
        int decimals, percentDecimals, yenDecimals;
        internal void Add(string value)
        {
            if (value.Length == 0) return;
            any = true;
            bool d = ColumnConversion.TryDate(value, out var dt); date &= d; hasDateTime |= d && dt.TimeOfDay != TimeSpan.Zero;
            bool t = ColumnConversion.TryDuration(value, out var days); time &= t; duration |= t && (days < 0 || days >= 1);
            number &= !ColumnConversion.PreserveAsText(value) && ColumnConversion.TryNumber(value, out _);
            percent &= ColumnConversion.TryPercent(value, out _); hasPercent |= value.EndsWith('%') || value.EndsWith('％');
            yen &= ColumnConversion.TryYen(value, out _); hasYen |= value.Contains('¥') || value.Contains('￥') || value.Contains('\\'); grouped |= value.Contains(',');
            decimals = Math.Max(decimals, ColumnConversion.InferDecimals([value], ColumnType.Number));
            percentDecimals = Math.Max(percentDecimals, ColumnConversion.InferDecimals([value], ColumnType.Percent));
            yenDecimals = Math.Max(yenDecimals, ColumnConversion.InferDecimals([value], ColumnType.Yen));
        }
        internal ColumnType Type => !any ? ColumnType.General : date ? hasDateTime ? ColumnType.DateTime : ColumnType.Date
            : time ? duration ? ColumnType.Duration : ColumnType.Time : percent && hasPercent ? ColumnType.Percent
            : yen && hasYen ? ColumnType.Yen : number ? grouped ? ColumnType.GroupedNumber : ColumnType.Number : ColumnType.TextFormat;
        internal int Decimals => Type == ColumnType.Percent ? percentDecimals : Type == ColumnType.Yen ? yenDecimals : ColumnConversion.IsNumeric(Type) ? decimals : 0;
    }
    internal static (ColumnType[] Types, int[] Decimals) Scan(TabularData table, bool header)
    {
        var stats = Enumerable.Range(0, table.ColumnCount).Select(_ => new Stats()).ToArray();
        for (int r = header ? 1 : 0; r < table.RowCount; r++)
        { var row = table.Rows[r]; for (int c = 0; c < row.Length; c++) stats[c].Add(row[c]); }
        return (stats.Select(s => s.Type).ToArray(), stats.Select(s => s.Decimals).ToArray());
    }
}
