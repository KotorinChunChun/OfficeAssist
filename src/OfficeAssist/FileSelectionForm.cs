namespace OfficeAssist;

internal sealed class FileSelectionForm : BrandedForm
{
    readonly TextBox search = new() { Name = "InputPattern", Width = 520 };
    readonly DataGridView grid = new() { Name = "InputFiles", Dock = DockStyle.Fill, VirtualMode = true, AutoGenerateColumns = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AllowUserToOrderColumns = false, RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false };
    readonly Label status = new() { AutoSize = true };
    readonly Button next = new() { Text = "次へ", AutoSize = true, Enabled = false };
    readonly ImportComboBox encoding = new() { Name = "InputEncoding", Width = 240 };
    readonly CheckBox encodingAll = new() { Text = "チェックした全ファイルに適用", AutoSize = true };
    bool syncingEncoding;
    readonly System.Windows.Forms.Timer debounce = new() { Interval = 200 };
    readonly CancellationTokenSource cancellation = new();
    readonly IReadOnlyList<string> inputs;
    readonly InputSelection? initial;
    InputSelection? selection;
    int[] visible = [];
    bool discovering;
    internal InputCandidate[] Selected { get; private set; } = [];
    internal InputSelection? Selection => selection;
    internal FileSelectionForm(IReadOnlyList<string> inputs, InputSelection? initial = null)
    {
        this.inputs = inputs; this.initial = initial;
        Text = WindowTitle.Create("ファイルの絞り込み"); ClientSize = new(1050, 650); MinimumSize = new(650, 400); ShowInTaskbar = false; MinimizeBox = false; TopMost = true;
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new(12), RowCount = 5, ColumnCount = 1 };
        foreach (int height in new[] { 42, 42, 42, 0, 70 }) layout.RowStyles.Add(height == 0 ? new(SizeType.Percent, 100) : new(SizeType.Absolute, height));
        var patternBar = new FlowLayoutPanel { Dock = DockStyle.Fill }; var clear = new Button { Text = "クリア", AutoSize = true }; clear.Click += (_, _) => search.Clear(); patternBar.Controls.AddRange([new Label { Text = "検索", AutoSize = true }, search, clear]);
        var operations = new FlowLayoutPanel { Dock = DockStyle.Fill };
        void Add(string label, Action action) { var button = new Button { Text = label, AutoSize = true }; button.Click += (_, _) => { if (selection is null) return; grid.EndEdit(); action(); RefreshRows(); }; operations.Controls.Add(button); }
        Add("すべて選択", () => selection!.SetAll(Enumerable.Range(0, selection.Files.Count), true));
        Add("すべて解除", () => selection!.SetAll(Enumerable.Range(0, selection.Files.Count), false));
        Add("表示分を選択", () => selection!.SetAll(visible, true)); Add("表示分を解除", () => selection!.SetAll(visible, false));
        Add("上へ", () => { if (grid.CurrentCell is not null) selection!.Move(visible[grid.CurrentCell.RowIndex], -1); });
        Add("下へ", () => { if (grid.CurrentCell is not null) selection!.Move(visible[grid.CurrentCell.RowIndex], 1); }); Add("パス順", () => selection!.Sort());
        grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Checked", Width = 40, SortMode = DataGridViewColumnSortMode.NotSortable });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Order", HeaderText = "順", Width = 65, ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Encoding", HeaderText = "文字コード", Width = 170, ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "FullPath", HeaderText = "フルパス", Width = 1400, ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable });
        grid.CellValueNeeded += (_, e) => { if (selection is null || e.RowIndex < 0 || e.RowIndex >= visible.Length) return; var file = selection.Files[visible[e.RowIndex]]; e.Value = e.ColumnIndex == 0 ? selection.IsChecked(file.Path) : e.ColumnIndex == 1 ? visible[e.RowIndex] + 1 : e.ColumnIndex == 2 ? TabularEncoding.Labels[(int)file.Encoding] : file.Path; };
        grid.CellValuePushed += (_, e) => { if (selection is not null && e.RowIndex < visible.Length && e.ColumnIndex == 0) { selection.Check(selection.Files[visible[e.RowIndex]].Path, e.Value is true); UpdateStatus(); } };
        grid.CurrentCellDirtyStateChanged += (_, _) => { if (grid.IsCurrentCellDirty) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
        grid.CellToolTipTextNeeded += (_, e) => { if (selection is not null && e.RowIndex >= 0 && e.RowIndex < visible.Length) e.ToolTipText = selection.Files[visible[e.RowIndex]].Path; };
        encoding.Items.AddRange(TabularEncoding.Labels); encoding.SelectedIndex = 0;
        var encodingBar = new FlowLayoutPanel { Dock = DockStyle.Fill }; encodingBar.Controls.AddRange([new Label { Text = "文字コード", AutoSize = true }, encoding, encodingAll]);
        grid.SelectionChanged += (_, _) => SyncEncoding();
        encoding.SelectedIndexChanged += (_, _) =>
        {
            if (syncingEncoding || selection is null || encoding.SelectedIndex < 0 || grid.CurrentCell is null || grid.CurrentCell.RowIndex >= visible.Length) return;
            var indexes = encodingAll.Checked ? Enumerable.Range(0, selection.Files.Count).Where(i => selection.IsChecked(selection.Files[i].Path)) : [visible[grid.CurrentCell.RowIndex]];
            selection.SetEncoding(indexes, (TableEncoding)encoding.SelectedIndex); grid.Invalidate();
        };
        encodingAll.CheckedChanged += (_, _) =>
        {
            if (encodingAll.Checked && selection is not null && encoding.SelectedIndex >= 0)
            { selection.SetEncoding(Enumerable.Range(0, selection.Files.Count).Where(i => selection.IsChecked(selection.Files[i].Path)), (TableEncoding)encoding.SelectedIndex); grid.Invalidate(); }
        };
        var bottom = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = true };
        var cancel = new Button { Text = "キャンセル", AutoSize = true }; cancel.Click += (_, _) => { cancellation.Cancel(); if (!discovering) { DialogResult = DialogResult.Cancel; Close(); } };
        next.Click += (_, _) => { grid.EndEdit(); if (errors.Length != 0 && MessageBox.Show(this, $"検索できなかった場所などが{errors.Length}件あります。検索結果の詳細を確認し、読めた対象だけで進めますか？", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return; Selected = selection!.Selected; DialogResult = DialogResult.OK; Close(); };
        var issues = new Button { Text = "検索結果の詳細", AutoSize = true }; issues.Click += (_, _) => { if (errors.Length != 0) MessageBox.Show(this, string.Join(Environment.NewLine, errors), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); };
        bottom.Controls.AddRange([cancel, next, issues, status]);
        layout.Controls.Add(patternBar, 0, 0); layout.Controls.Add(operations, 0, 1); layout.Controls.Add(encodingBar, 0, 2); layout.Controls.Add(grid, 0, 3); layout.Controls.Add(bottom, 0, 4); Controls.Add(layout);
        debounce.Tick += (_, _) => { debounce.Stop(); RefreshRows(); }; search.TextChanged += (_, _) => { debounce.Stop(); debounce.Start(); };
        Shown += async (_, _) => await Discover();
        FormClosing += (_, e) => { if (discovering) { cancellation.Cancel(); e.Cancel = true; } };
    }
    string[] errors = []; int duplicates;
    async Task Discover()
    {
        discovering = true; status.Text = "フォルダを検索中";
        try
        {
            if (initial is not null) selection = new(initial.Files) ;
            else { var progress = new Progress<int>(n => { if (!IsDisposed) status.Text = $"候補 {n}件"; }); var result = await Task.Run(() => TabularInputDiscovery.Find(inputs, cancellation.Token, progress)); selection = new(result.Files); errors = result.Errors.ToArray(); duplicates = result.Duplicates; }
            if (initial is not null) foreach (var file in initial.Files) selection!.Check(file.Path, initial.IsChecked(file.Path));
            RefreshRows();
        }
        catch (OperationCanceledException) { discovering = false; DialogResult = DialogResult.Cancel; Close(); }
        catch (Exception ex) { status.Text = ex.Message; errors = [ex.Message]; }
        finally { discovering = false; UpdateStatus(); }
    }
    void RefreshRows() { if (selection is null) return; visible = selection.Visible(search.Text); grid.RowCount = visible.Length; grid.Invalidate(); UpdateStatus(); SyncEncoding(); }
    void SyncEncoding()
    {
        if (selection is null || grid.CurrentCell is null || grid.CurrentCell.RowIndex < 0 || grid.CurrentCell.RowIndex >= visible.Length) return;
        syncingEncoding = true; encoding.SelectedIndex = (int)selection.Files[visible[grid.CurrentCell.RowIndex]].Encoding; syncingEncoding = false;
    }
    void UpdateStatus()
    {
        if (selection is null) return; int count = selection.Selected.Length;
        int shown = visible.Count(i => selection.IsChecked(selection.Files[i].Path));
        status.Text = $"候補 {selection.Files.Count} / 表示 {visible.Length} / 選択 {count}（非表示 {count - shown}）/ 重複 {duplicates} / 検索問題 {errors.Length}";
        next.Enabled = !discovering && count is > 0 and <= 50;
        if (count > 50) status.Text += " / 対象は50ファイルまで";
    }
    protected override void Dispose(bool disposing) { if (disposing) { cancellation.Cancel(); cancellation.Dispose(); debounce.Dispose(); } base.Dispose(disposing); }
}
