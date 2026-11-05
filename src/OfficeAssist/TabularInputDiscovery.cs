using System.IO.Enumeration;

namespace OfficeAssist;

internal sealed record InputCandidate(string Path, long Length, DateTime Modified)
{
    internal TableEncoding Encoding { get; init; } = TableEncoding.Auto;
}
internal sealed record DiscoveryResult(IReadOnlyList<InputCandidate> Files, IReadOnlyList<string> Errors, int Duplicates);

internal static class TabularInputDiscovery
{
    internal static bool Accepts(IReadOnlyList<string>? inputs) => inputs is { Count: > 0 and <= 50 }
        && inputs.All(p => !string.IsNullOrWhiteSpace(p) && (Directory.Exists(p) || TabularFiles.Accepts([p])));
    internal static DiscoveryResult Find(IReadOnlyList<string> inputs, CancellationToken token = default, IProgress<int>? progress = null)
    {
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase); var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var files = new List<InputCandidate>(); var errors = new List<string>(); int duplicates = 0;
        void Add(string path)
        {
            if (!TabularFiles.Accepts([path])) return;
            path = System.IO.Path.GetFullPath(path);
            if (!paths.Add(path)) { duplicates++; return; }
            try { var info = new FileInfo(path); files.Add(new(path, info.Length, info.LastWriteTimeUtc)); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { errors.Add(path + ": " + ex.Message); }
            if (files.Count > 100000) throw new InvalidDataException("検索候補が10万件を超えました。探索元を絞り直してください。");
            if (files.Count % 100 == 0) progress?.Report(files.Count);
        }
        foreach (string input in inputs)
        {
            token.ThrowIfCancellationRequested(); string root = System.IO.Path.GetFullPath(input);
            if (!Directory.Exists(root)) { Add(root); continue; }
            var stack = new Stack<string>(); stack.Push(root);
            while (stack.TryPop(out string? directory))
            {
                token.ThrowIfCancellationRequested(); if (!visited.Add(directory)) { duplicates++; continue; }
                try
                {
                    if ((File.GetAttributes(directory) & FileAttributes.ReparsePoint) != 0) { errors.Add(directory + ": リンクフォルダは探索しません。"); continue; }
                    foreach (string entry in Directory.EnumerateFileSystemEntries(directory))
                    {
                        token.ThrowIfCancellationRequested();
                        try { var attributes = File.GetAttributes(entry); if ((attributes & FileAttributes.Directory) != 0) stack.Push(entry); else Add(entry); }
                        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { errors.Add(entry + ": " + ex.Message); }
                    }
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { errors.Add(directory + ": " + ex.Message); }
            }
        }
        progress?.Report(files.Count);
        return new(files.OrderBy(f => f.Path, StringComparer.OrdinalIgnoreCase).ToArray(), errors, duplicates);
    }
    internal static bool Matches(string path, string pattern)
    {
        if (string.IsNullOrEmpty(pattern)) return true;
        if (pattern.IndexOfAny(['*', '?']) < 0) pattern = "*" + pattern + "*";
        return FileSystemName.MatchesSimpleExpression(pattern, path, true);
    }
}

internal sealed class InputSelection
{
    internal List<InputCandidate> Files { get; }
    readonly HashSet<string> selected;
    internal InputSelection(IEnumerable<InputCandidate> files) { Files = files.ToList(); selected = Files.Select(f => f.Path).ToHashSet(StringComparer.OrdinalIgnoreCase); }
    internal bool IsChecked(string path) => selected.Contains(path);
    internal void Check(string path, bool value) { if (value) selected.Add(path); else selected.Remove(path); }
    internal int[] Visible(string pattern) => Enumerable.Range(0, Files.Count).Where(i => TabularInputDiscovery.Matches(Files[i].Path, pattern)).ToArray();
    internal InputCandidate[] Selected => Files.Where(f => selected.Contains(f.Path)).ToArray();
    internal void SetAll(IEnumerable<int> indexes, bool value) { foreach (int i in indexes) Check(Files[i].Path, value); }
    internal void Move(int index, int delta) { int next = index + delta; if (next < 0 || next >= Files.Count) return; (Files[index], Files[next]) = (Files[next], Files[index]); }
    internal void Sort() => Files.Sort((a, b) => StringComparer.OrdinalIgnoreCase.Compare(a.Path, b.Path));
    internal void SetEncoding(IEnumerable<int> indexes, TableEncoding encoding)
    { if (!Enum.IsDefined(encoding)) throw new InvalidDataException("文字コードが不正です。"); foreach (int index in indexes) Files[index] = Files[index] with { Encoding = encoding }; }
}
