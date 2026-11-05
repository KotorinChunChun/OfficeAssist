using System.Runtime.InteropServices;

namespace OfficeAssist;

internal static class ExcelDropData
{
    internal static string[]? Read(IDataObject? data, DragDropEffects allowed)
    {
        if ((allowed & DragDropEffects.Copy) == 0 || data is null) return null;
        try
        {
            if (!data.GetDataPresent(DataFormats.FileDrop, false)
                || data.GetData(DataFormats.FileDrop, false) is not string[] paths
                || (!ImageFiles.Accepts(paths) && !TabularInputDiscovery.Accepts(paths))) return null;
            return (string[])paths.Clone();
        }
        catch (Exception ex) when (ex is ExternalException or ArgumentException or InvalidOperationException or NotSupportedException)
        { return null; }
    }
}
