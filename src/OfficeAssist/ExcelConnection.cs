using System.Runtime.InteropServices;

namespace OfficeAssist;

internal static class ExcelConnection
{
    [DllImport("user32.dll")] static extern bool EnumWindows(Native.EnumWindowProc callback, nint parameter);
    internal static IReadOnlyList<(int Pid, nint Handle)> Windows()
    {
        var result = new List<(int, nint)>(); var seen = new HashSet<uint>();
        EnumWindows((root, _) =>
        {
            if (Native.ClassName(root) != "XLMAIN") return true;
            Native.GetWindowThreadProcessId(root, out uint pid);
            if (seen.Contains(pid)) return true;
            Native.EnumChildWindows(root, (child, _) => { if (Native.ClassName(child) == "EXCEL7") { result.Add(((int)pid, child)); seen.Add(pid); return false; } return true; }, 0);
            return true;
        }, 0);
        return result;
    }
    internal static object Attach(ComScope com, nint handle)
    {
        var iid = new Guid("00020400-0000-0000-C000-000000000046");
        Marshal.ThrowExceptionForHR(Native.AccessibleObjectFromWindow(handle, 0xFFFFFFF0, ref iid, out object window));
        return com.Own(window);
    }
}
