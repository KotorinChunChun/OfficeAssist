namespace OfficeAssist;

internal sealed class TabularImportForm : BrandedForm
{
    readonly ComboBox destination = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 230 };
    readonly CheckBox header = new() { Text = "先頭行を見出しとして扱う", AutoSize = true };
    readonly DataGridView columns = new() { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
        RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, EditMode = DataGridViewEditMode.EditOnEnter };
    readonly IReadOnlyList<TabularData> tables;
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    internal TabularImportOptions? Options { get; private set; }
    internal TabularImportForm(IReadOnlyList<TabularData> tables)
    {
        this.tables = tables;
        Text = WindowTitle.Create("CSV・TSVの取り込み");
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(870, 500); MinimizeBox = false; MaximizeBox = false; ShowInTaskbar = false;
        Font = new Font("Yu Gothic UI", 10);
        destination.Items.AddRange(["新しいブック", "新しいシート", "選択中のセル"]);
        destination.SelectedIndex = (int)ImportDestination.Selection;
        header.Checked = TabularImportOptions.Infer(tables).Header;
        var top = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, AutoScroll = true };
        top.Controls.Add(new Label { Text = "展開先", AutoSize = true, Padding = new Padding(0, 5, 0, 0) });
        top.Controls.Add(destination); top.Controls.Add(header);
        columns.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ファイル", ReadOnly = true, FillWeight = 24 });
        columns.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "列", ReadOnly = true, FillWeight = 18 });
        columns.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "値の例", ReadOnly = true, FillWeight = 26 });
        var typeColumn = new DataGridViewComboBoxColumn { HeaderText = "形式", FillWeight = 32 };
        typeColumn.Items.AddRange(ColumnConversion.Labels); columns.Columns.Add(typeColumn);
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var cancel = new Button { Text = "キャンセル", AutoSize = true, DialogResult = DialogResult.Cancel };
        var import = new Button { Text = "取り込む", AutoSize = true };
        buttons.Controls.Add(cancel); buttons.Controls.Add(import);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 3 };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        layout.Controls.Add(top, 0, 0); layout.Controls.Add(columns, 0, 1); layout.Controls.Add(buttons, 0, 2); Controls.Add(layout);
        CancelButton = cancel;
        void Populate()
        {
            columns.Rows.Clear();
            var inferred = TabularImportOptions.Infer(tables, header: header.Checked);
            for (int t = 0; t < tables.Count; t++)
                for (int c = 0; c < tables[t].ColumnCount; c++)
                {
                    string title = header.Checked && c < tables[t].Rows[0].Length ? tables[t].Rows[0][c] : "列" + (c + 1);
                    string example = string.Join(" / ", tables[t].Rows.Skip(header.Checked ? 1 : 0).Take(3).Select(r => c < r.Length ? r[c] : ""));
                    columns.Rows.Add(tables[t].Name, $"{c + 1}: {title}", example.Length > 100 ? example[..100] : example, ColumnConversion.Labels[(int)inferred.Types[t][c]]);
                }
        }
        header.CheckedChanged += (_, _) => Populate(); Populate();
        import.Click += (_, _) =>
        {
            columns.EndEdit();
            var types = tables.Select(t => new ColumnType[t.ColumnCount]).ToArray();
            int row = 0;
            for (int t = 0; t < types.Length; t++)
                for (int c = 0; c < types[t].Length; c++)
                    types[t][c] = (ColumnType)Array.IndexOf(ColumnConversion.Labels, columns.Rows[row++].Cells[3].Value?.ToString() ?? "");
            var options = new TabularImportOptions((ImportDestination)destination.SelectedIndex, types, header.Checked);
            try { for (int t = 0; t < tables.Count; t++) tables[t].ExcelValues(options.Types[t], options.Header); }
            catch (Exception ex) when (ex is InvalidDataException or ArgumentException)
            { MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            Options = options; DialogResult = DialogResult.OK; Close();
        };
    }
}
