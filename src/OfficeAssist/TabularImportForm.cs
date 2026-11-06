namespace OfficeAssist;

// 表内の編集用コンボを使わず、候補の背景と文字色を明示して描画する。
internal sealed class ImportComboBox : ComboBox
{
    internal HashSet<int> DisabledItems { get; } = new();
    int previous = -1;
    internal ImportComboBox()
    {
        DropDownStyle = ComboBoxStyle.DropDownList; DrawMode = DrawMode.OwnerDrawFixed;
        BackColor = SystemColors.Window; ForeColor = SystemColors.WindowText;
        IntegralHeight = true; MaxDropDownItems = 8; ItemHeight = Font.Height + 6;
    }
    protected override void OnFontChanged(EventArgs e) { base.OnFontChanged(e); ItemHeight = Font.Height + 6; }
    protected override void OnSelectedIndexChanged(EventArgs e)
    { if (DisabledItems.Contains(SelectedIndex)) { SelectedIndex = previous; return; } previous = SelectedIndex; base.OnSelectedIndexChanged(e); }
    protected override void OnDrawItem(DrawItemEventArgs e)
    {
        bool selected = (e.State & DrawItemState.Selected) != 0;
        using var brush = new SolidBrush(selected ? SystemColors.Highlight : SystemColors.Window);
        e.Graphics.FillRectangle(brush, e.Bounds);
        string text = e.Index >= 0 && e.Index < Items.Count ? GetItemText(Items[e.Index]) ?? "" : Text;
        TextRenderer.DrawText(e.Graphics, text, Font, Rectangle.Inflate(e.Bounds, -3, 0), DisabledItems.Contains(e.Index) ? SystemColors.GrayText : selected ? SystemColors.HighlightText : SystemColors.WindowText,
            TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
        e.DrawFocusRectangle();
    }
}

internal sealed partial class TabularImportForm : BrandedForm
{
    const int ColumnWidth = 240, RowHeaderWidth = 48;
    readonly ImportComboBox destination = new() { Name = "ImportDestination", Width = 210 };
    readonly ImportComboBox files = new() { Name = "ImportFile", Width = 360, AccessibleName = "ファイル名を選択" };
    readonly CheckBox applyToAll = new() { Name = "ApplyToAllFiles", Text = "すべてのファイルに適用", AutoSize = true };
    readonly CheckBox header = new() { Name = "ImportHeader", Text = "先頭行を見出しとして扱う", AutoSize = true };
    readonly CheckBox merge = new() { Name = "MergeFiles", Text = "すべてのファイルを結合", AutoSize = true };
    readonly CheckBox dropEmpty = new() { Name = "DropEmptyRows", Text = "空欄行を除外", AutoSize = true };
    readonly NumericUpDown skipFirst = new() { Name = "SkipFirstRows", Width = 80, Maximum = int.MaxValue };
    readonly NumericUpDown skipLast = new() { Name = "SkipLastRows", Width = 80, Maximum = int.MaxValue };
    readonly ImportComboBox mergeBy = new() { Name = "MergeBy", Width = 100 };
    readonly ImportComboBox previewMode = new() { Name = "PreviewMode", Width = 210 };
    readonly ImportComboBox encoding = new() { Name = "ImportEncoding", Width = 190 };
    readonly Label mergeByLabel = LabelFor("結合基準");
    readonly Panel settingsViewport = new() { Dock = DockStyle.Fill, BackColor = SystemColors.Window };
    readonly Panel settings = new() { Height = 144 };
    readonly HScrollBar settingsScroll = new() { Dock = DockStyle.Bottom, SmallChange = 60 };
    readonly DataGridView preview = new() { Name = "ImportDataPreview", Dock = DockStyle.Fill, ReadOnly = true, VirtualMode = true,
        AllowUserToAddRows = false, AllowUserToDeleteRows = false, AllowUserToOrderColumns = false,
        AllowUserToResizeColumns = false, AllowUserToResizeRows = false, RowHeadersWidth = RowHeaderWidth,
        BackgroundColor = SystemColors.Window, AutoGenerateColumns = false, RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing };
    readonly ToolTip tips = new();
    IReadOnlyList<TabularData> sources;
    IReadOnlyList<TabularData> tables;
    TabularImportSelection selection;
    readonly Func<TabularImportOptions, ExcelPictureInserter.Result>? execute;
    readonly Action<IWin32Window, string, MessageBoxIcon> showMessage;
    readonly ExcelTargetEditor? targets;
    readonly Func<IReadOnlyList<TabularData>?>? reselect;
    readonly bool usePreviewCache;
    readonly Func<string?, int>? readEncoding;
    readonly Func<string?, TableEncoding, IReadOnlyList<TabularData>?>? reencode;
    TabularPreviewCache? previewCache;
    readonly Button import = new() { Text = "取り込む", AutoSize = true };
    readonly Button cancel = new() { Text = "キャンセル", AutoSize = true };
    readonly Button chooseFiles = new() { Text = "対象ファイル…", AutoSize = true };
    readonly Label status = new() { Dock = DockStyle.Fill, AutoEllipsis = true };
    readonly FlowLayoutPanel fileBar = Bar();
    readonly Dictionary<string, (ColumnType? Type, int? Decimals)> edits = new();
    string currentKey = "file";
    bool currentByName, updating, working, pendingClose, headerEdited;
    int generation, tasks;
    CancellationTokenSource? rebuildCancellation, writeCancellation;
    int[] previewRows = [];
    string[] notices = [];
    bool valid = true;
    int SelectedFile => Math.Clamp(files.SelectedIndex, 0, Math.Max(0, tables.Count - 1));
    internal IReadOnlyList<TabularData> Tables => tables;
    internal bool IsReady => valid && tasks == 0;
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    internal TabularImportOptions? Options { get; private set; }
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    internal ExcelPictureInserter.Result? Result { get; private set; }
    internal TabularImportForm(IReadOnlyList<TabularData> tables, Func<TabularImportOptions, ExcelPictureInserter.Result>? execute = null,
        Action<IWin32Window, string, MessageBoxIcon>? showMessage = null, ExcelTargetEditor? targets = null,
        Func<IReadOnlyList<TabularData>?>? reselect = null, bool defer = false,
        Func<string?, int>? readEncoding = null, Func<string?, TableEncoding, IReadOnlyList<TabularData>?>? reencode = null,
        IExcelFormatPreview? formatPreview = null,
        Func<IReadOnlyList<TabularData>, TabularImportOptions, int, IReadOnlyList<(int Row, int Column)>, ExcelPreviewSnapshot>? previewContext = null)
    {
        if (tables.Count == 0) throw new ArgumentException("ファイルがありません。", nameof(tables));
        sources = tables; usePreviewCache = defer; this.tables = defer ? new[] { new TabularData("読込中", [[""]]) } : tables;
        selection = new(this.tables); this.targets = targets; this.reselect = reselect;
        this.execute = execute; this.showMessage = showMessage ?? ((owner, message, icon) => MessageBox.Show(owner, message, Text, MessageBoxButtons.OK, icon));
        this.readEncoding = readEncoding; this.reencode = reencode;
        Text = WindowTitle.Create("CSV・TSVの取り込み"); StartPosition = FormStartPosition.CenterParent;
        ClientSize = new(1180, 760); MinimumSize = new(860, 600); MinimizeBox = false; MaximizeBox = true; ShowInTaskbar = false; TopMost = true; Font = new("Yu Gothic UI", 10);
        destination.Items.AddRange(["新しいブック", "新しいシート", "指定セル", "同名のシートに追記"]); destination.SelectedIndex = (int)ImportDestination.Selection;
        mergeBy.Items.AddRange(["列順", "列名"]); mergeBy.SelectedIndex = 0;
        previewMode.Items.AddRange(["しない", "各ファイル先頭10行", "すべて"]); previewMode.SelectedIndex = 1;
        encoding.Items.AddRange(TabularEncoding.Labels.Concat(new[] { "ファイル別（混在）" }).ToArray()); encoding.SelectedIndex = 0; encoding.Enabled = reencode is not null;
        header.Checked = selection.Header;
        var top = Bar(); top.Controls.AddRange([LabelFor("展開先"), destination]); if (targets is not null) top.Controls.Add(targets);
        var rows = Bar(); rows.Controls.AddRange([LabelFor("先頭から除外"), skipFirst, LabelFor("行　末尾から除外"), skipLast, LabelFor("行"), dropEmpty, LabelFor("文字コード"), encoding]);
        var modeBar = Bar(); modeBar.Controls.AddRange([merge, header, mergeByLabel, mergeBy, LabelFor("プレビュー"), previewMode, chooseFiles, validationDetails]);
        fileBar.Controls.AddRange([LabelFor("ファイル名を選択"), files, applyToAll]);
        tips.SetToolTip(applyToAll, "他のファイルの同じ列番号へ列形式と小数桁を適用します。行除外は全ファイル共通です。");
        settingsViewport.Controls.Add(settings); settingsViewport.Controls.Add(settingsScroll);
        preview.RowTemplate.Height = 27; preview.DefaultCellStyle.BackColor = SystemColors.Window; preview.DefaultCellStyle.ForeColor = SystemColors.WindowText;
        preview.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight; preview.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;
        preview.CellValueNeeded += (_, e) => { if (e.RowIndex >= 0 && e.RowIndex < previewRows.Length && e.ColumnIndex >= 0) { if (formattedPreview is not null) { e.Value = valid ? formattedPreview.Get(previewRows[e.RowIndex], e.ColumnIndex)?.Display ?? "…" : ""; } else { var row = usePreviewCache ? previewCache?.Get(previewRows[e.RowIndex]) : this.tables[SelectedFile].Rows[previewRows[e.RowIndex]]; e.Value = row is null ? "" : e.ColumnIndex < row.Length ? row[e.ColumnIndex] : ""; } } };
        preview.CellToolTipTextNeeded += (_, e) => { if (e.RowIndex >= 0 && e.RowIndex < previewRows.Length) e.ToolTipText = PreviewToolTip(previewRows[e.RowIndex], e.ColumnIndex); };
        preview.RowPostPaint += (_, e) => { if (e.RowIndex < previewRows.Length) TextRenderer.DrawText(e.Graphics, (previewRows[e.RowIndex] + 1).ToString(), Font,
            new Rectangle(e.RowBounds.Left, e.RowBounds.Top, preview.RowHeadersWidth - 4, e.RowBounds.Height), SystemColors.ControlText, TextFormatFlags.Right | TextFormatFlags.VerticalCenter); };
        preview.Scroll += (_, _) => { if (preview.HorizontalScrollingOffset <= settingsScroll.Maximum) settingsScroll.Value = preview.HorizontalScrollingOffset; PopulateSettings(); };
        settingsScroll.Scroll += (_, _) => { if (preview.Visible) try { preview.HorizontalScrollingOffset = settingsScroll.Value; } catch (ArgumentOutOfRangeException) { } PopulateSettings(); };
        settingsViewport.SizeChanged += (_, _) => PopulateSettings(); files.SelectedIndexChanged += (_, _) => { if (!updating) Populate(); };
        header.CheckedChanged += async (_, _) => { if (!updating) { headerEdited = true; await Rebuild(); } };
        skipFirst.ValueChanged += async (_, _) => await Rebuild(autoHeader: !headerEdited); skipLast.ValueChanged += async (_, _) => await Rebuild(autoHeader: !headerEdited); dropEmpty.CheckedChanged += async (_, _) => await Rebuild(autoHeader: !headerEdited);
        merge.CheckedChanged += async (_, _) => { UpdateModes(); await Rebuild(); }; mergeBy.SelectedIndexChanged += async (_, _) => { if (!updating) await Rebuild(); };
        previewMode.SelectedIndexChanged += (_, _) => Populate();
        encoding.SelectedIndexChanged += async (_, _) =>
        {
            if (updating || working || !valid || encoding.SelectedIndex is < 0 or > 4 || this.reencode is null) return;
            CancelPreviewWork(); await validationWork;
            try
            {
                string? file = merge.Checked || applyToAll.Checked ? null : this.tables[SelectedFile].SourcePath;
                var decoded = this.reencode(file, (TableEncoding)encoding.SelectedIndex);
                if (decoded is not null) { sources = decoded; edits.Clear(); await Rebuild(autoHeader: !headerEdited); }
            }
            catch (Exception ex) { this.showMessage(this, ex.Message, MessageBoxIcon.Warning); }
            finally { SyncEncoding(); RefreshFormatting(); }
        };
        destination.SelectedIndexChanged += (_, _) => { UpdateModes(); RefreshFormatting(); };
        applyToAll.CheckedChanged += (_, _) => { if (!updating && valid) { selection.SetApplyToAll(applyToAll.Checked, SelectedFile); if (applyToAll.Checked) for (int c = 0; c < tables[SelectedFile].ColumnCount; c++) { Remember(SelectedFile, c, true); Remember(SelectedFile, c, false); } PopulateSettings(); SyncEncoding(); RefreshFormatting(); } };
        chooseFiles.Enabled = reselect is not null; chooseFiles.Click += async (_, _) => { CancelPreviewWork(); await validationWork; try { var selected = this.reselect?.Invoke(); if (selected is null) return; sources = selected; edits.Clear(); await Rebuild(); } catch (Exception ex) { this.showMessage(this, ex.Message, MessageBoxIcon.Warning); } finally { RefreshFormatting(); } };
        import.Click += (_, _) => TryImport(); cancel.Click += (_, _) => { CancelPreviewWork(); if (working) writeCancellation?.Cancel(); else if (tasks > 0) { pendingClose = true; rebuildCancellation?.Cancel(); } else { DialogResult = DialogResult.Cancel; Close(); } };
        var buttons = Bar(); buttons.FlowDirection = FlowDirection.RightToLeft; buttons.Controls.AddRange([cancel, import]);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new(12), ColumnCount = 1, RowCount = 8 };
        foreach (int height in new[] { 76, 42, 76, 40, 164, 0, 38, 42 }) layout.RowStyles.Add(height == 0 ? new(SizeType.Percent, 100) : new(SizeType.Absolute, height));
        Control[] controls = [top, rows, modeBar, fileBar, settingsViewport, preview, status, buttons]; for (int i = 0; i < controls.Length; i++) layout.Controls.Add(controls[i], 0, i); Controls.Add(layout);
        FormClosing += (_, e) => { if (working || tasks > 0) { CancelPreviewWork(); writeCancellation?.Cancel(); rebuildCancellation?.Cancel(); pendingClose = true; e.Cancel = true; } };
        Shown += async (_, _) => { BringToFront(); Activate(); if (defer) await Rebuild(autoHeader: true); };
        InitializeFormatting(formatPreview ?? (defer ? new ExcelFormatPreview() : null), previewContext);
        PopulateFiles(); UpdateModes(); Populate(); if (defer) { valid = false; import.Enabled = false; }
    }
    static Label LabelFor(string text) => new() { Text = text, AutoSize = true, Padding = new(0, 5, 0, 0) };
    static FlowLayoutPanel Bar() => new() { Dock = DockStyle.Fill, WrapContents = true, AutoScroll = true };
    string EditKey(int f, int c, string mode, bool byName)
    {
        var table = tables[f]; string file = mode == "merged" ? "結合" : table.Origin?.Invoke(0)?.Split(": 行")[0] ?? table.Name;
        string column = byName && selection.Header && c < table.Rows[0].Length ? TabularTransform.Key(table.Rows[0][c]) : c.ToString();
        return mode + "|" + (byName ? "name" : "order") + "|" + file + "|" + column;
    }
    void Remember(int f, int c, bool type)
    {
        string key = EditKey(f, c, currentKey, currentByName); edits.TryGetValue(key, out var prior);
        edits[key] = type ? (selection.Types[f][c], prior.Decimals) : (prior.Type, selection.DecimalPlaces[f][c]);
        if (selection.ApplyToAll) for (int i = 0; i < tables.Count; i++) if (c < tables[i].ColumnCount) { string other = EditKey(i, c, currentKey, currentByName); edits.TryGetValue(other, out var previous); edits[other] = type ? (selection.Types[i][c], previous.Decimals) : (previous.Type, selection.DecimalPlaces[i][c]); }
    }
    async Task Rebuild(bool autoHeader = false)
    {
        if (updating || working) return;
        CancelPreviewWork();
        int revision = ++generation; rebuildCancellation?.Cancel(); var cancellation = new CancellationTokenSource(); rebuildCancellation = cancellation; tasks++;
        valid = false; import.Enabled = chooseFiles.Enabled = encoding.Enabled = false; status.Text = "行加工・列形式を解析中";
        var rules = new TableTransformOptions((int)skipFirst.Value, (int)skipLast.Value, dropEmpty.Checked, merge.Checked, merge.Checked && mergeBy.SelectedIndex == 1);
        bool hasHeader = header.Checked; var input = sources;
        try
        {
            var built = await Task.Run(() =>
            {
                if (autoHeader) hasHeader = TabularImportOptions.Infer(TabularTransform.Prepare(input, rules with { Merge = false, MergeByName = false }, false, cancellation.Token).Tables, token: cancellation.Token).Header;
                var result = TabularTransform.Prepare(input, rules with { MergeByName = rules.Merge && hasHeader && rules.MergeByName }, hasHeader, cancellation.Token);
                var model = new TabularImportSelection(result.Tables, cancellation.Token, hasHeader);
                return (result, model);
            });
            if (revision != generation || cancellation.IsCancellationRequested) return;
            tables = built.result.Tables; selection = built.model; currentKey = rules.Merge ? "merged" : "file"; currentByName = rules.MergeByName && hasHeader;
            updating = true; header.Checked = hasHeader; updating = false;
            for (int f = 0; f < tables.Count; f++) for (int c = 0; c < tables[f].ColumnCount; c++) if (edits.TryGetValue(EditKey(f, c, currentKey, currentByName), out var edit)) { if (edit.Type.HasValue) selection.SetType(f, c, edit.Type.Value); if (edit.Decimals.HasValue) selection.SetDecimals(f, c, edit.Decimals.Value); }
            if (rules.Merge) { selection.SetType(0, 0, ColumnType.TextFormat); selection.SetDecimals(0, 0, 0); }
            else selection.SetApplyToAll(applyToAll.Checked, 0);
            notices = built.result.Notices.ToArray(); valid = true; PopulateFiles(); UpdateModes(); Populate();
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { if (revision == generation) { status.Text = ex.Message; previewRows = []; preview.RowCount = 0; } }
        finally { tasks--; cancellation.Dispose(); if (revision == generation) { rebuildCancellation = null; import.Enabled = valid && FormattingAllowsImport; chooseFiles.Enabled = reselect is not null; encoding.Enabled = reencode is not null; } if (pendingClose && tasks == 0) { DialogResult = DialogResult.Cancel; Close(); } }
    }
    void UpdateModes()
    {
        bool append = destination.SelectedIndex == (int)ImportDestination.AppendSameSheet;
        if (merge.Checked) destination.DisabledItems.Add((int)ImportDestination.AppendSameSheet); else destination.DisabledItems.Remove((int)ImportDestination.AppendSameSheet);
        merge.Enabled = !append && !working; mergeBy.Visible = mergeByLabel.Visible = merge.Checked && header.Checked;
        fileBar.Visible = !merge.Checked; targets?.SetMode((ImportDestination)destination.SelectedIndex);
        if (merge.Checked && append) { valid = false; import.Enabled = false; status.Text = "追記と結合は同時に指定できません。"; }
    }
    void PopulateFiles() { updating = true; int selected = files.SelectedIndex; files.Items.Clear(); files.Items.AddRange(tables.Select(t => (object)t.Name).ToArray()); if (tables.Count > 0) files.SelectedIndex = Math.Clamp(selected, 0, tables.Count - 1); updating = false; }
    internal bool TryImport()
    {
        if (working || !valid || tasks > 0 || !FormattingAllowsImport) return false;
        if (merge.Checked && destination.SelectedIndex == (int)ImportDestination.AppendSameSheet) { showMessage(this, "追記と結合は同時に指定できません。", MessageBoxIcon.Warning); return false; }
        working = true; writeCancellation = new(); import.Enabled = false; cancel.Text = "中止";
        var editors = Controls.Cast<Control>().ToArray();
        destination.Enabled = files.Enabled = header.Enabled = merge.Enabled = dropEmpty.Enabled = skipFirst.Enabled = skipLast.Enabled = previewMode.Enabled = mergeBy.Enabled = applyToAll.Enabled = chooseFiles.Enabled = settingsViewport.Enabled = encoding.Enabled = false;
        if (targets is not null) targets.Enabled = false;
        try
        {
            var options = selection.CreateOptions((ImportDestination)destination.SelectedIndex) with { Merged = merge.Checked, Cancellation = writeCancellation.Token, Progress = text => { status.Text = text; if (Visible) Application.DoEvents(); } };
            if (targets is not null) options = targets.Apply(options);
            ExcelImportEngine.Validate(tables, options);
            var result = execute?.Invoke(options);
            if (result?.Warning is not null) showMessage(this, result.Warning, MessageBoxIcon.Warning);
            Options = options; Result = result; working = false; DialogResult = DialogResult.OK; Close(); return true;
        }
        catch (Exception ex) { DialogResult = DialogResult.None; if (Visible) { BringToFront(); Activate(); } showMessage(this, ex.Message, MessageBoxIcon.Warning); return false; }
        finally
        {
            working = false; writeCancellation.Dispose(); writeCancellation = null;
            if (!IsDisposed) { import.Enabled = valid && FormattingAllowsImport; cancel.Text = "キャンセル"; destination.Enabled = files.Enabled = header.Enabled = dropEmpty.Enabled = skipFirst.Enabled = skipLast.Enabled = previewMode.Enabled = mergeBy.Enabled = applyToAll.Enabled = settingsViewport.Enabled = true; chooseFiles.Enabled = reselect is not null; encoding.Enabled = reencode is not null; if (targets is not null) targets.Enabled = true; UpdateModes(); }
        }
    }
    void Populate()
    {
        if (tables.Count == 0 || updating) return;
        preview.RowCount = 0; preview.Columns.Clear(); preview.HorizontalScrollingOffset = 0; settingsScroll.Value = 0;
        var table = tables[SelectedFile];
        previewCache?.Dispose(); previewCache = usePreviewCache && formattedPreview is null ? new(table, () => { if (!IsDisposed && IsHandleCreated) try { BeginInvoke(() => { if (!IsDisposed) preview.Invalidate(); }); } catch (InvalidOperationException) { } }) : null;
        for (int c = 0; c < table.ColumnCount; c++) { string title = selection.Header && c < table.Rows[0].Length ? table.Rows[0][c] : "列" + (c + 1); preview.Columns.Add(new DataGridViewTextBoxColumn { Name = "Data" + c, HeaderText = title, Width = ColumnWidth, SortMode = DataGridViewColumnSortMode.NotSortable }); }
        previewRows = TabularTransform.PreviewRows(table, selection.Header, (PreviewMode)Math.Max(0, previewMode.SelectedIndex)); preview.RowCount = previewRows.Length; preview.Visible = previewMode.SelectedIndex != 0;
        PopulateSettings(); SyncEncoding();
        if (valid) { basePreviewStatus = $"{tables.Count}表 / {tables.Sum(t => t.RowCount - (selection.Header ? 1 : 0))}データ行 / 表示 {previewRows.Length}行" + (table.DetectedEncoding is null ? "" : " / " + table.DetectedEncoding) + (notices.Length == 0 ? "" : " / " + string.Join("　", notices.Take(3))); status.Text = basePreviewStatus; tips.SetToolTip(status, string.Join(Environment.NewLine, notices)); RefreshFormatting(); }
    }
    void SyncEncoding()
    {
        if (readEncoding is null || tables.Count == 0) return;
        bool prior = updating; updating = true;
        encoding.SelectedIndex = readEncoding(merge.Checked || applyToAll.Checked ? null : tables[SelectedFile].SourcePath);
        updating = prior;
    }
    void PopulateSettings()
    {
        if (tables.Count == 0) return;
        settings.SuspendLayout();
        try
        {
            foreach (Control control in settings.Controls.Cast<Control>().ToArray()) control.Dispose();
            var table = tables[SelectedFile]; int offset = settingsScroll.Value;
            settings.Left = RowHeaderWidth; settings.Width = Math.Max(0, settingsViewport.ClientSize.Width - RowHeaderWidth);
            settingsScroll.LargeChange = Math.Max(1, settings.Width); settingsScroll.Maximum = Math.Max(0, table.ColumnCount * ColumnWidth - 1);
            int start = offset / ColumnWidth, end = Math.Min(table.ColumnCount, start + settings.Width / ColumnWidth + 2);
            for (int c = start; c < end; c++)
            {
                int column = c; int file = SelectedFile;
                string title = selection.Header && c < table.Rows[0].Length ? table.Rows[0][c] : "列" + (c + 1);
                var label = new Label { Text = $"{c + 1}: {title}", AutoEllipsis = true, Left = c * ColumnWidth - offset + 4, Top = 4, Width = ColumnWidth - 8, Height = 27 };
                var type = new ImportComboBox { Name = "ColumnType" + c, AccessibleName = $"列{c + 1}の形式", Left = label.Left, Top = 34, Width = ColumnWidth - 8, DropDownWidth = ColumnWidth + 40 };
                type.Items.AddRange(ColumnConversion.Labels); type.SelectedIndex = (int)selection.Types[file][c]; type.Enabled = !(merge.Checked && c == 0);
                var digitsLabel = new Label { Text = "小数点以下桁数", Left = label.Left, Top = 72, Width = 136, Height = 27 };
                var digits = new NumericUpDown { Name = "ColumnDecimals" + c, Left = label.Left + 144, Top = 68, Width = 80, Minimum = 0, Maximum = 15, Value = selection.DecimalPlaces[file][c], Enabled = ColumnConversion.IsNumeric(selection.Types[file][c]) };
                digitsLabel.Enabled = digits.Enabled;
                bool assigningDigits = false;
                var format = new TextBox { Name = "ColumnFormat" + c, AccessibleName = $"列{c + 1}の書式コード", Left = label.Left + 48, Top = 105, Width = ColumnWidth - 56, ReadOnly = true, Text = ExcelValueFormat.Code(selection.Types[file][c], selection.DecimalPlaces[file][c]) };
                var formatLabel = new Label { Text = "書式", Left = label.Left, Top = 108, Width = 44, Height = 25 };
                type.SelectedIndexChanged += (_, _) => { if (type.SelectedIndex < 0) return; var chosen = (ColumnType)type.SelectedIndex; selection.SetType(file, column, chosen); Remember(file, column, true); digits.Enabled = digitsLabel.Enabled = ColumnConversion.IsNumeric(chosen); assigningDigits = true; digits.Value = selection.DecimalPlaces[file][column]; assigningDigits = false; RefreshFormatting(); };
                digits.ValueChanged += (_, _) => { if (assigningDigits) return; selection.SetDecimals(file, column, (int)digits.Value); Remember(file, column, false); RefreshFormatting(); };
                tips.SetToolTip(label, title); tips.SetToolTip(digits, "0なら小数点を表示しません。保存する数値は変更しません。"); tips.SetToolTip(format, "Excelへ適用する書式コードです。セルごとの実際の書式と格納値はプレビューのツールチップで確認できます。"); settings.Controls.AddRange([label, type, digitsLabel, digits, formatLabel, format]);
            }
        }
        finally { settings.ResumeLayout(); }
    }
    protected override void Dispose(bool disposing) { if (disposing) { CancelPreviewWork(); rebuildCancellation?.Cancel(); formattedPreview?.Dispose(); previewCache?.Dispose(); tips.Dispose(); } base.Dispose(disposing); }
}
