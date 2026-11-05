namespace OfficeAssist;

internal static class ExcelAddress
{
    internal static string Display(dynamic range) => (string)range.Address[false, false, 1];
}
