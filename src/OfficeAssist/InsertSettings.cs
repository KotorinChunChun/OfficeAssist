using System.Text.Json;

namespace OfficeAssist;

internal enum DialogModifier { Shift, Ctrl, Alt }
internal sealed record DropDialogBehavior(DialogModifier Modifier = DialogModifier.Shift, bool ImagesNormally = false, bool TablesNormally = true)
{
    internal bool Show(bool table, int keyState) => (table ? TablesNormally : ImagesNormally) != InsertSettings.IsPressed(Modifier, keyState);
}

internal static class InsertSettings
{
    internal static InsertOptions Default => new(Destination.SelectionRange, PictureKind.Embedded, []);
    static string PathName => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAssist", "insert-settings.json");
    sealed record StoredSettings(Destination Destination, PictureKind Kind, DialogModifier Modifier = DialogModifier.Shift,
        bool ImagesNormally = false, bool TablesNormally = true);

    internal static InsertOptions Load(string? path = null)
    {
        var saved = Read(path);
        return new(saved.Destination, saved.Kind, []);
    }

    internal static DialogModifier LoadModifier(string? path = null) => Read(path).Modifier;
    internal static DropDialogBehavior LoadBehavior(string? path = null) { var saved = Read(path); return new(saved.Modifier, saved.ImagesNormally, saved.TablesNormally); }

    static StoredSettings Read(string? path = null)
    {
        path ??= PathName;
        if (!File.Exists(path)) return new(Destination.SelectionRange, PictureKind.Embedded);
        try
        {
            var saved = JsonSerializer.Deserialize<StoredSettings>(File.ReadAllText(path));
            if (saved is null || !Enum.IsDefined(saved.Destination) || !Enum.IsDefined(saved.Kind) || !Enum.IsDefined(saved.Modifier)) throw new InvalidDataException();
            return saved;
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or InvalidDataException)
        { throw new InvalidOperationException("挿入設定を読み込めません。設定ファイルを確認してください。", ex); }
    }

    internal static bool IsPressed(DialogModifier modifier, int keyState) => (keyState & (modifier switch
    {
        DialogModifier.Shift => 4, DialogModifier.Ctrl => 8, DialogModifier.Alt => 32,
        _ => throw new ArgumentOutOfRangeException(nameof(modifier))
    })) != 0;

    internal static void Save(InsertOptions options, DialogModifier? modifier = null, string? path = null, bool? imagesNormally = null, bool? tablesNormally = null)
    {
        path ??= PathName;
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            var previous = Read(path);
            File.WriteAllText(temporary, JsonSerializer.Serialize(new StoredSettings(options.Destination, options.Kind, modifier ?? previous.Modifier,
                imagesNormally ?? previous.ImagesNormally, tablesNormally ?? previous.TablesNormally)));
            File.Move(temporary, path, true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
