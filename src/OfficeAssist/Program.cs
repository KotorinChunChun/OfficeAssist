using System.Security.Principal;

namespace OfficeAssist;

internal static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        ApplicationConfiguration.Initialize();
        using var mutex = new Mutex(true, @"Local\OfficeAssist." + WindowsIdentity.GetCurrent().User!.Value, out bool first);
        if (!first)
        {
            return;
        }
        try
        {
            using var context = new TrayContext();
            Application.Run(context);
        }
        finally { mutex.ReleaseMutex(); }
    }
}
