using System.Text.Json;

namespace OfficeAssist;

internal static class InsertSettings
{
    internal static InsertOptions Default => new(Destination.SelectionRange, PictureKind.Embedded, []);
    static string PathName => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAssist", "insert-settings.json");
    sealed record StoredSettings(Destination Destination, PictureKind Kind);

    internal static InsertOptions Load()
    {
        if (!File.Exists(PathName)) return Default;
        try
        {
            var saved = JsonSerializer.Deserialize<StoredSettings>(File.ReadAllText(PathName));
            if (saved is null || !Enum.IsDefined(saved.Destination) || !Enum.IsDefined(saved.Kind)) throw new InvalidDataException();
            return new(saved.Destination, saved.Kind, []);
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or InvalidDataException)
        { throw new InvalidOperationException("挿入設定を読み込めません。設定ファイルを確認してください。", ex); }
    }

    internal static void Save(InsertOptions options)
    {
        string path = PathName;
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.WriteAllText(temporary, JsonSerializer.Serialize(new StoredSettings(options.Destination, options.Kind)));
            File.Move(temporary, path, true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
