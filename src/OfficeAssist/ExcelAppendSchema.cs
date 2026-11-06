namespace OfficeAssist;

internal static class ExcelAppendSchema
{
    internal static void ValidateName(string name)
    {
        if (name.Length is 0 or > 31 || name.IndexOfAny("[]:*?/\\".ToCharArray()) >= 0 || name.StartsWith('\'') || name.EndsWith('\''))
            throw new InvalidDataException("同名シートとして使用できない名前です。");
    }
    internal static (int NextRow, string[] Headers) Read(ComScope com, dynamic sheet, bool header)
    {
        dynamic cells = com.Own(sheet.Cells);
        object? lastObject = cells.Find("*", After: com.Own(cells.Item(1, 1)), LookIn: -4123, LookAt: 2, SearchOrder: 1, SearchDirection: 2, MatchCase: false, MatchByte: false, SearchFormat: false);
        if (lastObject is null) return (1, []);
        dynamic last = com.Own(lastObject);
        object? rightObject = cells.Find("*", After: com.Own(cells.Item(1, 1)), LookIn: -4123, LookAt: 2, SearchOrder: 2, SearchDirection: 2, MatchCase: false, MatchByte: false, SearchFormat: false);
        dynamic right = com.Own(rightObject!); int width = (int)right.Column;
        if (!header) return ((int)last.Row + 1, new string[width]);
        dynamic range = com.Own(sheet.Range["A1"].Resize[1, width]); object raw = range.Value2;
        return ((int)last.Row + 1, Enumerable.Range(0, width).Select(c => raw is Array a ? Convert.ToString(a.GetValue(1, c + 1)) ?? "" : Convert.ToString(raw) ?? "").ToArray());
    }
    internal static int[] Map(TabularData source, string[] headers, bool header)
    {
        if (!header)
        { if (headers.Length != source.ColumnCount) throw new InvalidDataException($"{source.Name}: 既存シートと列数が異なります。"); return Enumerable.Range(0, source.ColumnCount).ToArray(); }
        var existing = TabularTransform.HeaderKeys(headers); var input = TabularTransform.HeaderKeys(source.Rows[0]);
        if (input.Length != source.ColumnCount || input.Any(key => !existing.Contains(key))) throw new InvalidDataException($"{source.Name}: 既存シートにない列があります。");
        return existing.Select(key => Array.IndexOf(input, key)).ToArray();
    }
}
