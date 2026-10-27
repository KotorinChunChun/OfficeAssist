namespace OfficeAssist;

internal static class WindowTitle
{
    public static string Create(string? screen = null)
    {
        var version = typeof(WindowTitle).Assembly.GetName().Version!;
        var title = $"OfficeAssist - ver {version.Major}.{version.Minor}.{version.Build}";
        return string.IsNullOrWhiteSpace(screen) ? title : $"{screen} - {title}";
    }
}

internal static class AppIcon
{
    public static Icon Create()
    {
        using var stream = typeof(AppIcon).Assembly.GetManifestResourceStream("OfficeAssist.AppIcon.ico")!;
        return new Icon(stream, 32, 32);
    }
}

internal class BrandedForm : Form
{
    readonly Icon ownedIcon = AppIcon.Create();
    public BrandedForm() => Icon = ownedIcon;
    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (disposing) ownedIcon.Dispose();
    }
}
