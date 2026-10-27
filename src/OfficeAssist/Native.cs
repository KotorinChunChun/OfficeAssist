using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeAssist;

internal static class Native
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct Rect { public int Left, Top, Right, Bottom; public readonly Rectangle Bounds => Rectangle.FromLTRB(Left, Top, Right, Bottom); }
    internal delegate bool EnumWindowProc(nint hwnd, nint parameter);
    [DllImport("user32.dll")] internal static extern short GetAsyncKeyState(int key);
    [DllImport("user32.dll")] internal static extern bool GetCursorPos(out Point point);
    [DllImport("user32.dll")] internal static extern nint WindowFromPoint(Point point);
    [DllImport("user32.dll")] internal static extern nint GetAncestor(nint hwnd, uint flags);
    [DllImport("user32.dll")] internal static extern nint GetParent(nint hwnd);
    [DllImport("user32.dll")] internal static extern bool GetWindowRect(nint hwnd, out Rect rect);
    [DllImport("user32.dll")] internal static extern bool IsWindowVisible(nint hwnd);
    [DllImport("user32.dll")] internal static extern bool IsWindowEnabled(nint hwnd);
    [DllImport("user32.dll")] internal static extern bool IsWindow(nint hwnd);
    [DllImport("user32.dll")] internal static extern bool EnumChildWindows(nint hwnd, EnumWindowProc callback, nint parameter);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(nint hwnd, StringBuilder name, int count);
    [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(nint hwnd, out uint pid);
    [DllImport("oleacc.dll")] internal static extern int AccessibleObjectFromWindow(nint hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object result);

    internal static string ClassName(nint hwnd)
    {
        var name = new StringBuilder(256);
        GetClassName(hwnd, name, name.Capacity);
        return name.ToString();
    }

    internal static bool IsExplorer(nint hwnd)
    {
        GetWindowThreadProcessId(hwnd, out uint pid);
        try { using var process = Process.GetProcessById((int)pid); return process.ProcessName.Equals("explorer", StringComparison.OrdinalIgnoreCase); }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception) { return false; }
    }

    internal static nint ExcelSheetAt(Point point)
    {
        var hwnd = WindowFromPoint(point);
        for (int depth = 0; hwnd != 0 && depth < 12; depth++, hwnd = GetParent(hwnd))
        {
            if (ClassName(hwnd) != "EXCEL7") continue;
            var root = GetAncestor(hwnd, 2);
            return ClassName(root) == "XLMAIN" && IsWindowEnabled(root) && IsWindowVisible(hwnd) ? hwnd : 0;
        }
        return 0;
    }
}
