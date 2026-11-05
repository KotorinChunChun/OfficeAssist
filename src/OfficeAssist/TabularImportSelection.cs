namespace OfficeAssist;

internal sealed class TabularImportSelection
{
    readonly IReadOnlyList<TabularData> tables;
    readonly bool[][] edited;
    readonly bool[][] decimalEdited;
    internal ColumnType[][] Types { get; private set; }
    internal int[][] DecimalPlaces { get; private set; }
    internal bool Header { get; private set; }
    internal bool ApplyToAll { get; private set; }
    internal TabularImportSelection(IReadOnlyList<TabularData> tables, CancellationToken token = default, bool? header = null)
    {
        this.tables = tables;
        var inferred = TabularImportOptions.Infer(tables, header: header, token: token);
        Header = inferred.Header; Types = inferred.Types;
        DecimalPlaces = inferred.DecimalPlaces!;
        edited = tables.Select(t => new bool[t.ColumnCount]).ToArray();
        decimalEdited = tables.Select(t => new bool[t.ColumnCount]).ToArray();
    }
    internal void SetHeader(bool value)
    {
        Header = value;
        var inferred = TabularImportOptions.Infer(tables, header: value);
        for (int f = 0; f < Types.Length; f++)
            for (int c = 0; c < Types[f].Length; c++)
            {
                if (!edited[f][c]) Types[f][c] = inferred.Types[f][c];
                if (!decimalEdited[f][c]) DecimalPlaces[f][c] = ColumnConversion.InferDecimals(tables[f].Rows.Skip(Header ? 1 : 0).Select(r => c < r.Length ? r[c] : ""), Types[f][c]);
            }
    }
    internal void SetType(int file, int column, ColumnType type)
    {
        if (!Enum.IsDefined(type) || file < 0 || file >= Types.Length || column < 0 || column >= Types[file].Length)
            throw new ArgumentOutOfRangeException();
        Types[file][column] = type; edited[file][column] = true;
        if (!decimalEdited[file][column]) DecimalPlaces[file][column] = ColumnConversion.InferDecimals(tables[file].Rows.Skip(Header ? 1 : 0).Select(r => column < r.Length ? r[column] : ""), type);
        if (ApplyToAll)
            for (int f = 0; f < Types.Length; f++)
                if (column < Types[f].Length) { Types[f][column] = type; edited[f][column] = true; DecimalPlaces[f][column] = DecimalPlaces[file][column]; }
    }
    internal void SetDecimals(int file, int column, int decimals)
    {
        if (decimals is < 0 or > 15) throw new ArgumentOutOfRangeException(nameof(decimals));
        DecimalPlaces[file][column] = decimals; decimalEdited[file][column] = true;
        if (ApplyToAll)
            for (int f = 0; f < Types.Length; f++)
                if (column < Types[f].Length) { DecimalPlaces[f][column] = decimals; decimalEdited[f][column] = true; }
    }
    internal void SetApplyToAll(bool value, int selectedFile)
    {
        ApplyToAll = value;
        if (value)
            for (int c = 0; c < Types[selectedFile].Length; c++) { SetType(selectedFile, c, Types[selectedFile][c]); SetDecimals(selectedFile, c, DecimalPlaces[selectedFile][c]); }
    }
    internal TabularImportOptions CreateOptions(ImportDestination destination) => new(destination, Types.Select(t => t.ToArray()).ToArray(), Header) { DecimalPlaces = DecimalPlaces.Select(d => d.ToArray()).ToArray() };
}
