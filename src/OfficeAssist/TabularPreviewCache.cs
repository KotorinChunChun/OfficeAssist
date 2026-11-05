namespace OfficeAssist;

// UIのCellValueNeededではI/Oせず、要求された行だけを有界キューから読む。
internal sealed class TabularPreviewCache : IDisposable
{
    readonly TabularData table;
    readonly Action loaded;
    readonly object gate = new();
    readonly Dictionary<int, string[]> cache = new();
    readonly Queue<int> order = new(), pending = new();
    readonly HashSet<int> requested = new();
    bool running, disposed;
    long characters;
    internal TabularPreviewCache(TabularData table, Action loaded) { this.table = table; this.loaded = loaded; }
    internal string[]? Get(int row)
    {
        lock (gate)
        {
            if (disposed) return null;
            if (cache.TryGetValue(row, out var values)) return values;
            if (requested.Count < 128 && requested.Add(row)) pending.Enqueue(row);
            if (!running && pending.Count > 0) { running = true; _ = Task.Run(Read); }
            return null;
        }
    }
    void Read()
    {
        while (true)
        {
            int row;
            lock (gate) { if (disposed || !pending.TryDequeue(out row)) { running = false; return; } }
            string[] values;
            try { values = table.Rows[row]; } catch (ObjectDisposedException) { return; } catch (IOException) { return; }
            lock (gate)
            {
                if (disposed) return;
                requested.Remove(row); cache[row] = values; order.Enqueue(row); characters += values.Sum(v => (long)v.Length);
                while ((characters > 2_000_000 || cache.Count > 256) && cache.Count > 1 && order.TryDequeue(out int expired))
                    if (cache.Remove(expired, out var old)) characters -= old.Sum(v => (long)v.Length);
            }
            loaded();
        }
    }
    public void Dispose() { lock (gate) { disposed = true; cache.Clear(); pending.Clear(); requested.Clear(); } }
}
