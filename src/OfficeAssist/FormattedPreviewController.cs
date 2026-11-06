namespace OfficeAssist;

internal sealed class FormattedPreviewController : IDisposable
{
    readonly Control owner;
    readonly IExcelFormatPreview renderer;
    readonly Action changed;
    readonly System.Windows.Forms.Timer timer = new() { Interval = 16 };
    readonly Dictionary<(int Row, int Column), ExcelPreviewCell> cache = new();
    readonly Queue<(int Row, int Column)> order = new();
    readonly HashSet<(int Row, int Column)> pending = new();
    TabularPreviewCache? raw;
    ColumnType[] types = []; int[] decimals = [];
    Func<IReadOnlyList<(int Row, int Column)>, ExcelPreviewSnapshot>? context;
    CancellationTokenSource cancellation = new();
    int generation; bool inFlight, active, disposed;
    long characters;
    int notificationQueued;
    internal bool Busy => active && (inFlight || pending.Count > 0);
    internal string? Failure { get; private set; }
    internal int WarningCount => cache.Values.Count(c => c.Warning is not null);
    internal IEnumerable<ExcelPreviewCell> Errors => cache.Values.Where(c => c.Error is not null);
    internal FormattedPreviewController(Control owner, IExcelFormatPreview renderer, Action changed)
    { this.owner = owner; this.renderer = renderer; this.changed = changed; timer.Tick += async (_, _) => await Flush(); }
    internal void Reset(TabularData table, ColumnType[] columnTypes, int[] decimalPlaces, Func<IReadOnlyList<(int Row, int Column)>, ExcelPreviewSnapshot> capture)
    {
        Suspend(); active = true; types = columnTypes.ToArray(); decimals = decimalPlaces.ToArray(); context = capture;
        raw = new(table, () => Notify()); timer.Start();
    }
    internal void Suspend()
    {
        generation++; active = false; timer.Stop(); cancellation.Cancel(); cancellation.Dispose(); cancellation = new();
        raw?.Dispose(); raw = null; cache.Clear(); order.Clear(); pending.Clear(); characters = 0; Failure = null;
    }
    internal ExcelPreviewCell? Get(int row, int column)
    {
        if (!active || disposed || column < 0 || column >= types.Length) return null;
        if (cache.TryGetValue((row, column), out var cell)) return cell;
        if (pending.Count < 256 && pending.Add((row, column))) { raw?.Get(row); Notify(); }
        return null;
    }
    void Notify()
    {
        if (!owner.IsHandleCreated || owner.IsDisposed || disposed) return;
        if (Interlocked.Exchange(ref notificationQueued, 1) != 0) return;
        try { owner.BeginInvoke(async () => { Interlocked.Exchange(ref notificationQueued, 0); if (!disposed && !owner.IsDisposed) { changed(); await Flush(); } }); }
        catch (InvalidOperationException) { Interlocked.Exchange(ref notificationQueued, 0); }
    }
    async Task Flush()
    {
        if (disposed || !active || inFlight || context is null || pending.Count == 0) return;
        int revision = generation; var token = cancellation.Token;
        var addresses = new List<(int Row, int Column)>(); var values = new List<string>();
        foreach (var position in pending.ToArray())
        {
            var row = raw?.Get(position.Row); if (row is null) continue;
            addresses.Add(position); values.Add(position.Column < row.Length ? row[position.Column] : "");
            if (addresses.Count == 128) break;
        }
        if (addresses.Count == 0) return;
        inFlight = true;
        try
        {
            var snapshot = context(addresses);
            var inputs = addresses.Select((p, i) => new ExcelPreviewInput(p.Row, p.Column, values[i], types[p.Column], decimals[p.Column], snapshot.Styles[i])).ToArray();
            var cells = await renderer.RenderAsync(inputs, snapshot.Environment, token);
            if (revision != generation || disposed) return;
            foreach (var cell in cells)
            {
                var key = (cell.Row, cell.Column); pending.Remove(key); cache[key] = cell; order.Enqueue(key);
                characters += cell.Raw.Length + cell.Display.Length;
                if (cell.Error is not null) Failure ??= cell.Error;
            }
            while ((cache.Count > 2000 || characters > 2_000_000) && cache.Count > 1 && order.TryDequeue(out var expired))
                if (cache.Remove(expired, out var old)) characters -= old.Raw.Length + old.Display.Length;
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            if (revision == generation && !disposed)
            {
                Failure = ex.Message;
                for (int i = 0; i < addresses.Count; i++) { var p = addresses[i]; pending.Remove(p); cache[p] = new(p.Row, p.Column, values[i], "確認エラー", null, ExcelValueFormat.Code(types[p.Column], decimals[p.Column]), "", "", ex.Message); }
            }
        }
        finally { inFlight = false; if (!disposed && revision == generation) changed(); }
    }
    public void Dispose()
    {
        if (disposed) return; Suspend(); disposed = true; timer.Dispose(); cancellation.Dispose(); renderer.Dispose();
    }
}
