namespace OfficeAssist;

internal enum ExcelTableDropZone { Grid, Bottom, Other }

internal static class ExcelTableDropSite
{
    internal static ExcelTableDropZone Zone(Point point, Rectangle sheet, int bottomBand)
    {
        if (sheet.IsEmpty) return ExcelTableDropZone.Other;
        if (point.Y >= sheet.Bottom - bottomBand) return ExcelTableDropZone.Bottom;
        return sheet.Contains(point) ? ExcelTableDropZone.Grid : ExcelTableDropZone.Other;
    }

    internal static ExcelTableDropZone Locate(Point point, nint sheetHandle, Rectangle bounds)
    {
        if (sheetHandle == 0) return ExcelTableDropZone.Other;
        int dpi = (int)Native.GetDpiForWindow(sheetHandle);
        int fallback = Math.Min(bounds.Height / 2, 32 * Math.Max(96, dpi) / 96);
        int band = 0;
        Native.EnumChildWindows(sheetHandle, (child, _) =>
        {
            if (!Native.IsWindowVisible(child) || !Native.GetWindowRect(child, out var rect)) return true;
            string kind = Native.ClassName(child);
            var control = rect.Bounds;
            if (control.Top >= bounds.Bottom - 2 * fallback && control.Top < bounds.Bottom
                && (kind == "XLCTL" || kind == "NUIScrollbar" && control.Width > control.Height))
                band = Math.Max(band, bounds.Bottom - control.Top);
            return true;
        }, 0);
        return Zone(point, bounds, band == 0 ? fallback : band);
    }

    internal static ImportDestination? PreferredDestination(ExcelTableDropZone zone, bool hasCell) => zone switch
    {
        ExcelTableDropZone.Bottom => ImportDestination.NewSheet,
        ExcelTableDropZone.Grid when hasCell => null,
        _ => ImportDestination.NewBook
    };
}
