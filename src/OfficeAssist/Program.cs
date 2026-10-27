using KotorinToys.Integration;
using System.Security.Principal;

namespace OfficeAssist;

internal static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        ApplicationConfiguration.Initialize();
        bool silent = args.Contains("--startup");
        using var mutex = new Mutex(true, @"Local\OfficeAssist." + WindowsIdentity.GetCurrent().User!.Value, out bool first);
        if (!first)
        {
            if (!silent && ManagementEndpoint.ShowExisting("office-assist") is { } error)
                MessageBox.Show(error, WindowTitle.Create(), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        try
        {
            using var context = new TrayContext(silent);
            Application.Run(context);
        }
        finally { mutex.ReleaseMutex(); }
    }
}
