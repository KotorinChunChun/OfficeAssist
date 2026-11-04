namespace OfficeAssist;

// 表内の編集用コンボを使わず、候補の背景と文字色を明示して描画する。
internal sealed class ImportComboBox : ComboBox
{
    internal ImportComboBox()
    {
        DropDownStyle = ComboBoxStyle.DropDownList; DrawMode = DrawMode.OwnerDrawFixed;
        BackColor = SystemColors.Window; ForeColor = SystemColors.WindowText;
        IntegralHeight = true; MaxDropDownItems = 8; ItemHeight = Font.Height + 6;
    }
    protected override void OnFontChanged(EventArgs e) { base.OnFontChanged(e); ItemHeight = Font.Height + 6; }
    protected override void OnDrawItem(DrawItemEventArgs e)
    {
        bool selected = (e.State & DrawItemState.Selected) != 0;
        using var brush = new SolidBrush(selected ? SystemColors.Highlight : SystemColors.Window);
        e.Graphics.FillRectangle(brush, e.Bounds);
        string text = e.Index >= 0 && e.Index < Items.Count ? GetItemText(Items[e.Index]) ?? "" : Text;
        TextRenderer.DrawText(e.Graphics, text, Font, Rectangle.Inflate(e.Bounds, -3, 0), selected ? SystemColors.HighlightText : SystemColors.WindowText,
            TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
        e.DrawFocusRectangle();
    }
}

internal sealed class TabularImportForm : BrandedForm
{
    const int ColumnWidth = 240, RowHeaderWidth = 48;
    readonly ImportComboBox destination = new() { Name = "ImportDestination", Width = 230 };
    readonly ImportComboBox files = new() { Name = "ImportFile", Width = 360, AccessibleName = "ファイル名を選択" };
    readonly CheckBox applyToAll = new() { Name = "ApplyToAllFiles", Text = "すべてのファイルに適用", AutoSize = true };
    readonly CheckBox header = new() { Text = "先頭行を見出しとして扱う", AutoSize = true };
    readonly Panel settingsViewport = new() { Dock = DockStyle.Fill, BackColor = SystemColors.Window };
    readonly Panel settings = new() { Height = 82 };
    readonly DataGridView preview = new() { Name = "ImportDataPreview", Dock = DockStyle.Fill, ReadOnly = true, VirtualMode = true,
        AllowUserToAddRows = false, AllowUserToDeleteRows = false, AllowUserToOrderColumns = false,
        AllowUserToResizeColumns = false, AllowUserToResizeRows = false, RowHeadersWidth = RowHeaderWidth,
        BackgroundColor = SystemColors.Window, AutoGenerateColumns = false, RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing };
    readonly ToolTip tips = new();
    readonly IReadOnlyList<TabularData> tables;
    readonly TabularImportSelection selection;
    int SelectedFile => files.SelectedIndex;
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    internal TabularImportOptions? Options { get; private set; }

    internal TabularImportForm(IReadOnlyList<TabularData> tables)
    {
        if (tables.Count == 0) throw new ArgumentException("ファイルがありません。", nameof(tables));
        this.tables = tables; selection = new(tables);
        Text = WindowTitle.Create("CSV・TSVの取り込み"); StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(1020, 600); MinimumSize = new Size(780, 440);
        MinimizeBox = false; MaximizeBox = true; ShowInTaskbar = false; TopMost = true; Font = new Font("Yu Gothic UI", 10);
        Shown += (_, _) => { BringToFront(); Activate(); };
        destination.Items.AddRange(["新しいブック", "新しいシート", "選択中のセル"]); destination.SelectedIndex = (int)ImportDestination.Selection;
        files.Items.AddRange(tables.Select(t => (object)t.Name).ToArray()); files.SelectedIndex = 0;
        header.Checked = selection.Header;
        static Label LabelFor(string text) => new() { Text = text, AutoSize = true, Padding = new Padding(0, 5, 0, 0) };
        var top = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, AutoScroll = true };
        top.Controls.AddRange([LabelFor("展開先"), destination, header]);
        var fileBar = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, AutoScroll = true };
        fileBar.Controls.AddRange([LabelFor("ファイル名を選択"), files, applyToAll]);
        tips.SetToolTip(applyToAll, "選択中のファイルの列形式を、他のファイルの同じ列番号へ適用します。存在しない列は変更しません。");
        tips.SetToolTip(header, "すべてのファイルの先頭行を見出しとして表示します。手動で変更した列形式は維持します。");
        settingsViewport.Controls.Add(settings);
        preview.RowTemplate.Height = 27;
        preview.DefaultCellStyle.BackColor = SystemColors.Window; preview.DefaultCellStyle.ForeColor = SystemColors.WindowText;
        preview.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight; preview.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;
        preview.CellValueNeeded += (_, e) =>
        {
            int row = e.RowIndex + (selection.Header ? 1 : 0);
            if (SelectedFile < 0 || row < 0 || row >= tables[SelectedFile].RowCount || e.ColumnIndex < 0) return;
            var values = tables[SelectedFile].Rows[row]; e.Value = e.ColumnIndex < values.Length ? values[e.ColumnIndex] : "";
        };
        preview.CellToolTipTextNeeded += (_, e) =>
        {
            if (SelectedFile < 0 || e.RowIndex < 0 || e.ColumnIndex < 0 || e.RowIndex + (selection.Header ? 1 : 0) >= tables[SelectedFile].RowCount) return;
            var values = tables[SelectedFile].Rows[e.RowIndex + (selection.Header ? 1 : 0)];
            e.ToolTipText = e.ColumnIndex < values.Length ? values[e.ColumnIndex] : "";
        };
        preview.RowPostPaint += (_, e) => TextRenderer.DrawText(e.Graphics, (e.RowIndex + (selection.Header ? 2 : 1)).ToString(), Font,
            new Rectangle(e.RowBounds.Left, e.RowBounds.Top, preview.RowHeadersWidth - 4, e.RowBounds.Height), SystemColors.ControlText,
            TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
        preview.Scroll += (_, _) => PopulateSettings();
        settingsViewport.SizeChanged += (_, _) => PopulateSettings();
        files.SelectedIndexChanged += (_, _) => Populate();
        header.CheckedChanged += (_, _) => { selection.SetHeader(header.Checked); Populate(); };
        applyToAll.CheckedChanged += (_, _) => { selection.SetApplyToAll(applyToAll.Checked, SelectedFile); Populate(); };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var cancel = new Button { Text = "キャンセル", AutoSize = true, DialogResult = DialogResult.Cancel };
        var import = new Button { Text = "取り込む", AutoSize = true };
        buttons.Controls.AddRange([cancel, import]);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 5 };
        foreach (int height in new[] { 44, 44, 88, 0, 42 }) layout.RowStyles.Add(height == 0 ? new RowStyle(SizeType.Percent, 100) : new RowStyle(SizeType.Absolute, height));
        layout.Controls.Add(top, 0, 0); layout.Controls.Add(fileBar, 0, 1); layout.Controls.Add(settingsViewport, 0, 2);
        layout.Controls.Add(preview, 0, 3); layout.Controls.Add(buttons, 0, 4); Controls.Add(layout); CancelButton = cancel;
        import.Click += (_, _) =>
        {
            var options = selection.CreateOptions((ImportDestination)destination.SelectedIndex);
            try { for (int t = 0; t < tables.Count; t++) tables[t].ExcelValues(options.Types[t], options.Header); }
            catch (Exception ex) when (ex is InvalidDataException or ArgumentException)
            { MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            Options = options; DialogResult = DialogResult.OK; Close();
        };
        Populate();
    }

    void Populate()
    {
        if (SelectedFile < 0) return;
        settings.SuspendLayout(); preview.SuspendLayout();
        try
        {
            preview.RowCount = 0; preview.Columns.Clear(); preview.HorizontalScrollingOffset = 0;
            var table = tables[SelectedFile];
            for (int c = 0; c < table.ColumnCount; c++)
            {
                string title = selection.Header && c < table.Rows[0].Length ? table.Rows[0][c] : "列" + (c + 1);
                preview.Columns.Add(new DataGridViewTextBoxColumn { Name = "Data" + c, HeaderText = title, Width = ColumnWidth, SortMode = DataGridViewColumnSortMode.NotSortable });
            }
            preview.RowCount = table.RowCount - (selection.Header ? 1 : 0);
            PopulateSettings();
        }
        finally { preview.ResumeLayout(); settings.ResumeLayout(); }
    }
    void PopulateSettings()
    {
        if (SelectedFile < 0) return;
        settings.SuspendLayout();
        try
        {
            foreach (Control control in settings.Controls.Cast<Control>().ToArray()) control.Dispose();
            var table = tables[SelectedFile];
            int offset = preview.HorizontalScrollingOffset;
            settings.Left = RowHeaderWidth; settings.Width = Math.Max(0, settingsViewport.ClientSize.Width - RowHeaderWidth);
            int start = offset / ColumnWidth, end = Math.Min(table.ColumnCount, start + settings.Width / ColumnWidth + 2);
            // 見えている列だけに設定コントロールを作り、幅の大きなファイルでもハンドルを使い切らない。
            for (int c = start; c < end; c++)
            {
                int columnIndex = c;
                string title = selection.Header && c < table.Rows[0].Length ? table.Rows[0][c] : "列" + (c + 1);
                var label = new Label { Text = $"{c + 1}: {title}", AutoEllipsis = true, Left = c * ColumnWidth - offset + 4, Top = 4, Width = ColumnWidth - 8, Height = 27 };
                var type = new ImportComboBox { Name = "ColumnType" + c, AccessibleName = $"列{c + 1}の形式", Left = label.Left, Top = 34, Width = ColumnWidth - 8, DropDownWidth = ColumnWidth + 40 };
                type.Items.AddRange(ColumnConversion.Labels); type.SelectedIndex = (int)selection.Types[SelectedFile][c];
                type.SelectedIndexChanged += (_, _) => { if (type.SelectedIndex >= 0) selection.SetType(SelectedFile, columnIndex, (ColumnType)type.SelectedIndex); };
                tips.SetToolTip(label, title); settings.Controls.Add(label); settings.Controls.Add(type);
            }
        }
        finally { settings.ResumeLayout(); }
    }
    protected override void Dispose(bool disposing) { if (disposing) tips.Dispose(); base.Dispose(disposing); }
}
