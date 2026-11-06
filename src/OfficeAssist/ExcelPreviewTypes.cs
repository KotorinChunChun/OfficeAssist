namespace OfficeAssist;

internal sealed record ExcelPreviewStyle(double? WidthPoints = null, double? WidthCharacters = null,
    string? FontName = null, double? FontSize = null, bool Bold = false, bool Italic = false,
    bool Wrap = false, bool Shrink = false, double? RowHeight = null);
internal sealed record ExcelPreviewEnvironment(bool Date1904 = false, bool UseSystemSeparators = true,
    string DecimalSeparator = ".", string ThousandsSeparator = ",", string? NormalFont = null, double? NormalFontSize = null);
internal sealed record ExcelPreviewInput(int Row, int Column, string Raw, ColumnType Type, int Decimals, ExcelPreviewStyle? Style = null);
internal sealed record ExcelPreviewCell(int Row, int Column, string Raw, string Display, object? StoredValue,
    string FormatCode, string NumberFormat, string Prefix, string? Error = null, string? Warning = null);
internal interface IExcelFormatPreview : IDisposable
{
    Task<IReadOnlyList<ExcelPreviewCell>> RenderAsync(IReadOnlyList<ExcelPreviewInput> cells, ExcelPreviewEnvironment environment, CancellationToken token);
}
