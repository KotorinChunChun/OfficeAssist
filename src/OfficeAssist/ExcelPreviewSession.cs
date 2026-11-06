namespace OfficeAssist;

internal sealed record ExcelPreviewMetrics(int Cells, int Cached, int Literal, int ExcelCells, int StyleGroups, int WidthSearches);

// ExcelのTextは数値・日付に使用する。共通の書式と属性を範囲へまとめて設定する。
internal sealed class ExcelPreviewSession(object application, object workbook, string defaultFont, double defaultFontSize,
    bool cacheResults, bool batchStyles, bool literalText)
{
    sealed record Key(ExcelPreviewEnvironment Environment, ExcelPreviewStyle Style, string Raw, ColumnType Type, int Decimals);
    sealed record Group(ExcelPreviewStyle Style, string Code, bool Quote, int Separate);
    sealed record Pending(int Index, ExcelPreviewInput Input, object? Value, Key Key);
    readonly Dictionary<Key, ExcelPreviewCell> cache = new();
    readonly Queue<Key> cacheOrder = new();
    readonly Dictionary<(string Font, double Size, double Points), double> widths = new();
    long characters;
    ExcelPreviewEnvironment? priorEnvironment;
    internal ExcelPreviewMetrics Metrics { get; private set; } = new(0, 0, 0, 0, 0, 0);
    internal int CacheCount => cache.Count;

    // @で格納する文字列はExcelの数値・日付変換や丸めを受けない。
    // 自前で数値書式を模倣せず、入力文字列と引用符のエスケープだけを戻す。
    internal static bool TryLiteral(ExcelPreviewInput input, out ExcelPreviewCell cell)
    {
        // Excel.Textは非常に長い文字列を省略する。長文は必ずExcel自身へ問い合わせる。
        if (input.Raw.Length > 1024) { cell = null!; return false; }
        string code = ExcelValueFormat.Code(input.Type, input.Decimals);
        if (ColumnConversion.TryConvert(input.Raw, input.Type, false, out var value) && ExcelValueFormat.CellCode(input.Type, input.Decimals, value) == "@")
        {
            string? text = value as string; bool quote = text?.StartsWith('\'') == true; if (quote) text = text![1..];
            cell = new(input.Row, input.Column, input.Raw, text ?? "", text, code, "@", quote ? "'" : ""); return true;
        }
        cell = null!; return false;
    }
    internal IReadOnlyList<ExcelPreviewCell> Render(IReadOnlyList<ExcelPreviewInput> inputs, ExcelPreviewEnvironment environment, CancellationToken token)
    {
        using var com = new ComScope(); dynamic app = application, book = workbook;
        var output = new ExcelPreviewCell[inputs.Count]; var pending = new List<Pending>();
        var duplicates = new Dictionary<Key, List<int>>(); int hits = 0, literals = 0, searches = 0;
        for (int i = 0; i < inputs.Count; i++)
        {
            token.ThrowIfCancellationRequested(); var input = inputs[i]; var key = new Key(environment, input.Style ?? new(), input.Raw, input.Type, input.Decimals);
            if (cacheResults && cache.TryGetValue(key, out var found)) { output[i] = found with { Row = input.Row, Column = input.Column }; hits++; continue; }
            if (literalText && TryLiteral(input, out var text)) { output[i] = text; literals++; continue; }
            if (duplicates.TryGetValue(key, out var positions)) { positions.Add(i); continue; }
            duplicates[key] = [i]; string code = ExcelValueFormat.Code(input.Type, input.Decimals);
            if (!ColumnConversion.TryConvert(input.Raw, input.Type, false, out var value))
            {
                output[i] = new(input.Row, input.Column, input.Raw, "変換エラー", null, code, code, "", $"{ColumnConversion.Labels[(int)input.Type]}へ変換できない値があります。");
                continue;
            }
            pending.Add(new(i, input, ExcelValueFormat.AdjustSerial(value, input.Type, environment.Date1904), key));
        }
        if (pending.Count == 0)
        {
            foreach (var pair in duplicates) foreach (int i in pair.Value.Skip(1)) output[i] = output[pair.Value[0]] with { Row = inputs[i].Row, Column = inputs[i].Column };
            Metrics = new(inputs.Count, hits, literals, 0, 0, 0); return output;
        }
        dynamic sheet = com.Own(com.Own(book.Worksheets).Item(1)), first = com.Own(sheet.Range["A1"]), column = com.Own(first.EntireColumn);
        string normalName = environment.NormalFont ?? defaultFont; double normalSize = environment.NormalFontSize ?? defaultFontSize;
        if (environment != priorEnvironment)
        {
            app.UseSystemSeparators = environment.UseSystemSeparators;
            if (!environment.UseSystemSeparators) { app.DecimalSeparator = environment.DecimalSeparator; app.ThousandsSeparator = environment.ThousandsSeparator; }
            book.Date1904 = environment.Date1904;
            dynamic normalFont = com.Own(com.Own(com.Own(book.Styles).Item("Normal")).Font);
            normalFont.Name = normalName; normalFont.Size = normalSize; priorEnvironment = environment;
        }
        var groups = pending.GroupBy(p => new Group(p.Key.Style, ExcelValueFormat.CellCode(p.Input.Type, p.Input.Decimals, p.Value),
            p.Value is string value && value.StartsWith('\''), batchStyles ? 0 : p.Index)).ToArray();
        double standardWidth = pending.Count == 0 ? 0 : (double)sheet.StandardWidth;
        double standardHeight = pending.Count == 0 ? 0 : (double)sheet.StandardHeight;
        foreach (var grouping in groups)
        {
            token.ThrowIfCancellationRequested(); var items = grouping.ToArray(); var style = grouping.Key.Style;
            try
            {
                dynamic range = com.Own(first.Resize[items.Length, 1]); range.Clear();
                dynamic font = com.Own(range.Font); font.Name = style.FontName ?? normalName; font.Size = style.FontSize ?? normalSize;
                font.Bold = style.Bold; font.Italic = style.Italic; range.WrapText = style.Wrap; range.ShrinkToFit = style.Shrink;
                range.RowHeight = style.RowHeight ?? standardHeight;
                if (style.WidthPoints is double points)
                {
                    var widthKey = (normalName, normalSize, points);
                    if (!widths.TryGetValue(widthKey, out double width))
                    {
                        searches++; double low = 0, high = 255;
                        for (int n = 0; n < 16; n++) { column.ColumnWidth = (low + high) / 2; double actual = (double)column.Width; if (Math.Abs(actual - points) < 0.2) break; if (actual < points) low = (low + high) / 2; else high = (low + high) / 2; }
                        width = (double)column.ColumnWidth; if (widths.Count >= 256) widths.Clear(); widths[widthKey] = width;
                    }
                    column.ColumnWidth = width;
                }
                else column.ColumnWidth = style.WidthCharacters ?? standardWidth;
                ExcelValueFormat.Apply(app, range, grouping.Key.Code);
                var values = new object?[items.Length, 1]; for (int i = 0; i < items.Length; i++) values[i, 0] = items[i].Value;
                range.Value2 = values;
                if (range.HasFormula is not bool formula || formula) throw new InvalidOperationException("プレビューに意図しない数式があります。");
                bool stringsOnly = items.All(p => p.Value is null or string);
                object? stored = stringsOnly ? null : range.Value2; var array = stored as Array; string numberFormat = (string)range.NumberFormat;
                string prefix = (string)range.PrefixCharacter;
                dynamic cells = com.Own(range.Cells);
                for (int i = 0; i < items.Length; i++)
                {
                    token.ThrowIfCancellationRequested(); var input = items[i].Input; dynamic cell = com.Own(cells.Item(i + 1, 1)); string display = (string)cell.Text;
                    string? warning = display.Length > 0 && display.All(c => c == '#') ? "Excelではこの列幅・値・書式の組合せが#表示になります。列幅、負の時刻などを確認してください。" : null;
                    object? value = stringsOnly ? items[i].Value is string text && text.StartsWith('\'') ? text[1..] : items[i].Value : array?.GetValue(i + 1, 1) ?? (array is null ? stored : null);
                    output[items[i].Index] = new(input.Row, input.Column, input.Raw, display, value, ExcelValueFormat.Code(input.Type, input.Decimals), numberFormat, prefix, Warning: warning);
                }
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            { foreach (var p in items) { string code = ExcelValueFormat.Code(p.Input.Type, p.Input.Decimals); output[p.Index] = new(p.Input.Row, p.Input.Column, p.Input.Raw, "変換エラー", null, code, code, "", ex.Message); } }
        }
        foreach (var pair in duplicates)
        {
            var result = output[pair.Value[0]];
            foreach (int i in pair.Value.Skip(1)) output[i] = result with { Row = inputs[i].Row, Column = inputs[i].Column };
            if (cacheResults && result.Error is null)
            {
                long size = result.Raw.Length + result.Display.Length + ((result.StoredValue as string)?.Length ?? 0);
                if (size > 2_000_000) continue; cache[pair.Key] = result; cacheOrder.Enqueue(pair.Key); characters += size;
                while ((cache.Count > 4096 || characters > 2_000_000) && cacheOrder.TryDequeue(out var expired))
                    if (cache.Remove(expired, out var old)) characters -= old.Raw.Length + old.Display.Length + ((old.StoredValue as string)?.Length ?? 0);
            }
        }
        Metrics = new(inputs.Count, hits, literals, pending.Count, groups.Length, searches); return output;
    }
}
