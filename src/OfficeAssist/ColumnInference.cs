namespace OfficeAssist;

internal static class ColumnInference
{
    sealed class Stats
    {
        bool any, date = true, time = true, number = true, percent = true, yen = true;
        bool hasDateTime, duration, hasPercent, hasYen, grouped;
        int decimals, percentDecimals, yenDecimals;
        HashSet<string>? seen = new(StringComparer.Ordinal);
        internal bool Final => !date && !time && !number && !percent && !yen;
        internal void Add(string value)
        {
            if (value.Length == 0) return;
            if (Final) return;
            if (seen is not null) { if (!seen.Add(value)) return; if (seen.Count > 256) seen = null; }
            any = true;
            if (date) { bool d = ColumnConversion.TryDate(value, out var dt); date = d; hasDateTime |= d && dt.TimeOfDay != TimeSpan.Zero; }
            if (time) { bool t = ColumnConversion.TryDuration(value, out var days); time = t; duration |= t && (days < 0 || days >= 1); }
            if (!number && !percent && !yen) return;
            bool parsed = ColumnConversion.TryNumber(value, out double numeric);
            bool percentSign = value.EndsWith('%') || value.EndsWith('％'), yenSign = value.Contains('¥') || value.Contains('￥') || value.Contains('\\');
            hasPercent |= percentSign; hasYen |= yenSign; grouped |= value.Contains(',');
            if (number) { number = !ColumnConversion.PreserveAsText(value) && parsed; if (number) decimals = Math.Max(decimals, ColumnConversion.CountDecimals(numeric)); }
            if (percent) { double p = numeric; percent = percentSign ? ColumnConversion.TryPercent(value, out p) : parsed; if (percent) percentDecimals = Math.Max(percentDecimals, ColumnConversion.CountDecimals(p * 100)); }
            if (yen) { double y = numeric; yen = yenSign ? ColumnConversion.TryYen(value, out y) : parsed; if (yen) yenDecimals = Math.Max(yenDecimals, ColumnConversion.CountDecimals(y)); }
        }
        internal ColumnType Type => !any ? ColumnType.General : date ? hasDateTime ? ColumnType.DateTime : ColumnType.Date
            : time ? duration ? ColumnType.Duration : ColumnType.Time : percent && hasPercent ? ColumnType.Percent
            : yen && hasYen ? ColumnType.Yen : number ? grouped ? ColumnType.GroupedNumber : ColumnType.Number : ColumnType.TextFormat;
        internal int Decimals => Type == ColumnType.Percent ? percentDecimals : Type == ColumnType.Yen ? yenDecimals : ColumnConversion.IsNumeric(Type) ? decimals : 0;
    }
    internal static (ColumnType[] Types, int[] Decimals) Scan(TabularData table, bool header, CancellationToken token = default)
    {
        token.ThrowIfCancellationRequested();
        var stats = Enumerable.Range(0, table.ColumnCount).Select(_ => new Stats()).ToArray();
        for (int r = header ? 1 : 0; r < table.RowCount; r++)
        { if ((r & 255) == 0) token.ThrowIfCancellationRequested(); var row = table.Rows[r]; for (int c = 0; c < row.Length; c++) stats[c].Add(row[c]); if (stats.All(s => s.Final)) break; }
        return (stats.Select(s => s.Type).ToArray(), stats.Select(s => s.Decimals).ToArray());
    }
}
