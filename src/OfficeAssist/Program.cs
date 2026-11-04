using System.Security.Principal;

namespace OfficeAssist;

internal static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        ApplicationConfiguration.Initialize();
        if (args.Contains("--import-table")) return TabularImportCli.Run(args);
        using var mutex = new Mutex(true, @"Local\OfficeAssist." + WindowsIdentity.GetCurrent().User!.Value, out bool first);
        if (!first)
        {
            return 0;
        }
        try
        {
            using var context = new TrayContext();
            Application.Run(context);
        }
        finally { mutex.ReleaseMutex(); }
        return 0;
    }
}
