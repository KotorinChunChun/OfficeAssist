using System.Text.Json;

namespace OfficeAssist;

internal sealed record SavedTableColumns(string Name, string[] Headers, ColumnType?[] Types, int?[] Decimals)
{
    internal bool Explicit => Types.Any(t => t.HasValue) || Decimals.Any(d => d.HasValue);
}
internal sealed record TabularDropProfile(ImportDestination Destination = ImportDestination.Selection, bool? Header = null,
    TableTransformOptions? Rows = null, PreviewMode Preview = PreviewMode.First10, bool ApplyToAll = false,
    TableEncoding Encoding = TableEncoding.Auto, SavedTableColumns[]? Columns = null, bool MixedEncoding = false)
{
    // 文字コードの手動指定は、その取り込みの入力ファイルだけに適用する。
    internal TabularDropProfile ForNewDrop(ImportDestination? destination) =>
        this with { Destination = destination ?? Destination, Encoding = TableEncoding.Auto, MixedEncoding = false };

    internal PreparedTabularDrop Prepare(IReadOnlyList<TabularData> source, CancellationToken token = default, bool strict = true)
    {
        if (strict && MixedEncoding) throw new InvalidDataException("前回の文字コードがファイルごとに異なります。取り込み画面で指定してください。");
        var rules = Rows ?? new();
        bool header = Header ?? TabularImportOptions.Infer(TabularTransform.Prepare(source, rules with { Merge = false, MergeByName = false }, false, token).Tables, token: token).Header;
        var built = TabularTransform.Prepare(source, rules with { MergeByName = rules.Merge && header && rules.MergeByName }, header, token);
        var selection = new TabularImportSelection(built.Tables, token, header); ApplyColumns(built.Tables, selection, strict);
        if (rules.Merge) { selection.SetType(0, 0, ColumnType.TextFormat); selection.SetDecimals(0, 0, 0); }
        return new(built.Tables, selection, selection.CreateOptions(Destination) with { Merged = rules.Merge }, built.Notices);
    }
    internal void ApplyColumns(IReadOnlyList<TabularData> tables, TabularImportSelection selection, bool strict)
    {
        var profiles = Columns ?? [];
        for (int f = 0; f < tables.Count; f++)
        {
            var table = tables[f];
            string[] headers = selection.Header ? Enumerable.Range(0, table.ColumnCount).Select(c => c < table.Rows[0].Length ? TabularTransform.Key(table.Rows[0][c]) : "").ToArray() : [];
            var matching = profiles.Where(p => p.Types.Length == table.ColumnCount && p.Headers.SequenceEqual(headers)).ToArray();
            var named = matching.Where(p => p.Name == table.Name).ToArray();
            SavedTableColumns? chosen = named.Length == 1 ? named[0] : matching.Length == 1 ? matching[0]
                : matching.Length > 1 && matching.Select(p => JsonSerializer.Serialize(new { p.Types, p.Decimals })).Distinct().Count() == 1 ? matching[0] : null;
            if (chosen is null)
            {
                if (strict && profiles.Any(p => p.Explicit)) throw new InvalidDataException("保存した列設定と列構成が異なるか、複数の設定が候補になります。取り込み画面で確認してください。");
                continue;
            }
            var infer = new List<int>();
            for (int c = 0; c < table.ColumnCount; c++)
            {
                if (chosen.Types[c] is ColumnType type)
                {
                    if (type != selection.Types[f][c] && chosen.Decimals[c] is null && ColumnConversion.IsNumeric(type)) infer.Add(c);
                    selection.SetType(f, c, type, inferDecimals: false);
                }
                if (chosen.Decimals[c] is int digits) selection.SetDecimals(f, c, digits);
            }
            selection.InferRestoredDecimals(f, infer);
        }
    }
}
internal sealed record PreparedTabularDrop(IReadOnlyList<TabularData> Tables, TabularImportSelection Selection, TabularImportOptions Options, IReadOnlyList<string> Notices);

internal static class TabularDropSettings
{
    static string PathName => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAssist", "table-drop-settings.json");
    internal static TabularDropProfile Load(string? path = null)
    {
        path ??= PathName; if (!File.Exists(path)) return new();
        try { var profile = JsonSerializer.Deserialize<TabularDropProfile>(File.ReadAllText(path)) ?? throw new InvalidDataException(); Validate(profile); return profile; }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or InvalidDataException)
        { throw new InvalidOperationException("CSV・TSVの保存設定を読み込めません。設定ファイルを確認してください。", ex); }
    }
    internal static void Save(TabularDropProfile profile, string? path = null)
    {
        Validate(profile); path ??= PathName; Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try { File.WriteAllText(temporary, JsonSerializer.Serialize(profile)); File.Move(temporary, path, true); }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
    static void Validate(TabularDropProfile profile)
    {
        var rows = profile.Rows ?? new();
        if (!Enum.IsDefined(profile.Destination) || !Enum.IsDefined(profile.Preview) || !Enum.IsDefined(profile.Encoding)
            || rows.SkipFirst < 0 || rows.SkipLast < 0 || rows.Merge && profile.Destination == ImportDestination.AppendSameSheet
            || profile.Columns?.Length > 50) throw new InvalidDataException("CSV・TSVの保存設定が不正です。");
        foreach (var table in profile.Columns ?? [])
        {
            if (table is null || table.Name is null || table.Headers is null || table.Types is null || table.Decimals is null
                || table.Types.Length is < 1 or > 16384 || table.Decimals.Length != table.Types.Length
                || table.Headers.Length != 0 && table.Headers.Length != table.Types.Length || table.Headers.Any(h => h is null)
                || table.Types.Any(t => t.HasValue && !Enum.IsDefined(t.Value)) || table.Decimals.Any(d => d is < 0 or > 15)) throw new InvalidDataException("保存した列設定が不正です。");
        }
    }
}
