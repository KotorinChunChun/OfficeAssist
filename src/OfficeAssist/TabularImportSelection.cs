namespace OfficeAssist;

internal sealed class TabularImportSelection
{
    readonly IReadOnlyList<TabularData> tables;
    readonly bool[][] edited;
    internal ColumnType[][] Types { get; private set; }
    internal bool Header { get; private set; }
    internal bool ApplyToAll { get; private set; }
    internal TabularImportSelection(IReadOnlyList<TabularData> tables)
    {
        this.tables = tables;
        var inferred = TabularImportOptions.Infer(tables);
        Header = inferred.Header; Types = inferred.Types;
        edited = tables.Select(t => new bool[t.ColumnCount]).ToArray();
    }
    internal void SetHeader(bool value)
    {
        Header = value;
        var inferred = TabularImportOptions.Infer(tables, header: value);
        for (int f = 0; f < Types.Length; f++)
            for (int c = 0; c < Types[f].Length; c++)
                if (!edited[f][c]) Types[f][c] = inferred.Types[f][c];
    }
    internal void SetType(int file, int column, ColumnType type)
    {
        if (!Enum.IsDefined(type) || file < 0 || file >= Types.Length || column < 0 || column >= Types[file].Length)
            throw new ArgumentOutOfRangeException();
        Types[file][column] = type; edited[file][column] = true;
        if (ApplyToAll)
            for (int f = 0; f < Types.Length; f++)
                if (column < Types[f].Length) { Types[f][column] = type; edited[f][column] = true; }
    }
    internal void SetApplyToAll(bool value, int selectedFile)
    {
        ApplyToAll = value;
        if (value)
            for (int c = 0; c < Types[selectedFile].Length; c++) SetType(selectedFile, c, Types[selectedFile][c]);
    }
    internal TabularImportOptions CreateOptions(ImportDestination destination) => new(destination, Types.Select(t => t.ToArray()).ToArray(), Header);
}
