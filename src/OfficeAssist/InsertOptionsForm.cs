namespace OfficeAssist;

internal enum PictureKind { Embedded, Linked, ImageFormula }
internal enum Destination { DropPosition, SelectionPosition, SelectionRange }
internal sealed record InsertOptions(Destination Destination, PictureKind Kind, string[] Urls);

internal sealed class InsertOptionsForm : BrandedForm
{
    readonly RadioButton drop = new() { Text = "ドロップ位置（原寸大）", AutoSize = true };
    readonly RadioButton selection = new() { Text = "選択位置（原寸大）", AutoSize = true };
    readonly RadioButton fit = new() { Text = "選択範囲", AutoSize = true };
    readonly RadioButton embedded = new() { Text = "埋め込み", AutoSize = true, Checked = true };
    readonly RadioButton linked = new() { Text = "リンクされた図", AutoSize = true };
    readonly RadioButton formula = new() { Text = "セル内 IMAGE 関数式", AutoSize = true };
    readonly TextBox urls = new() { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
    readonly Label urlLabel = new() { Text = "画像の HTTPS URL（1画像につき1行）", AutoSize = true };
    readonly ToolTip tips = new();
    readonly ComboBox modifier = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 130 };
    readonly CheckBox imagesNormally = new() { Name = "ImagesDialogNormally", Text = "画像の挿入画面を表示する", AutoSize = true };
    readonly CheckBox tablesNormally = new() { Name = "TablesDialogNormally", Text = "CSV・TSVの取り込み画面を表示する", AutoSize = true };
    readonly int fileCount;
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    internal InsertOptions? Options { get; private set; }
    internal DialogModifier Modifier => (DialogModifier)modifier.SelectedIndex;
    internal bool ImagesNormally => imagesNormally.Checked;
    internal bool TablesNormally => tablesNormally.Checked;

    internal InsertOptionsForm(string selectedAddress, int count, bool allowLinkedImages = true, InsertOptions? defaults = null,
        bool settingsMode = false, DialogModifier dialogModifier = DialogModifier.Shift, DropDialogBehavior? behavior = null)
    {
        fileCount = count;
        linked.Name = "LinkedPictureOption";
        linked.Enabled = allowLinkedImages;
        Text = WindowTitle.Create(settingsMode ? "Excelドロップ設定" : "画像の挿入");
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
        ClientSize = new Size(500, 395);
        Font = new Font("Yu Gothic UI", 10);
        TopMost = true;
        defaults ??= InsertSettings.Default;
        drop.Checked = defaults.Destination == Destination.DropPosition;
        selection.Checked = defaults.Destination == Destination.SelectionPosition;
        fit.Checked = defaults.Destination == Destination.SelectionRange;
        embedded.Checked = defaults.Kind == PictureKind.Embedded || (defaults.Kind == PictureKind.Linked && !allowLinkedImages);
        linked.Checked = defaults.Kind == PictureKind.Linked && allowLinkedImages;
        formula.Checked = defaults.Kind == PictureKind.ImageFormula;
        tips.SetToolTip(fit, selectedAddress);
        modifier.Items.AddRange(["Shift", "Ctrl", "Alt"]);
        modifier.SelectedIndex = (int)dialogModifier;
        imagesNormally.Checked = behavior?.ImagesNormally ?? false; tablesNormally.Checked = behavior?.TablesNormally ?? true;
        static GroupBox Group(string text, params Control[] controls)
        {
            var group = new GroupBox { Text = text, Dock = DockStyle.Fill, Padding = new Padding(8) };
            var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = true, AutoScroll = true };
            flow.Controls.AddRange(controls); group.Controls.Add(flow); return group;
        }
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 5 };
        foreach (int height in new[] { 95, 62, 26, 0, 42 })
            layout.RowStyles.Add(height == 0 ? new RowStyle(SizeType.Percent, 100) : new RowStyle(SizeType.Absolute, height));
        layout.Controls.Add(Group(settingsMode ? "画像の挿入先" : "挿入先", drop, selection, fit), 0, 0);
        layout.Controls.Add(Group(settingsMode ? "画像の挿入形式" : "挿入形式", embedded, linked, formula), 0, 1);
        if (settingsMode)
        {
            ClientSize = new Size(520, 415); layout.RowStyles[2].Height = 68;
            layout.Controls.Add(Group("表示・省略を切り替える修飾キー", modifier), 0, 2);
            layout.Controls.Add(Group("通常ドロップ時", imagesNormally, tablesNormally), 0, 3);
            tips.SetToolTip(modifier, "押しながらドロップすると、通常時とは逆の表示・省略になります。");
            tips.SetToolTip(imagesNormally, "OFFなら通常は省略し、修飾キーで表示します。ONなら通常は表示し、修飾キーで省略します。");
            tips.SetToolTip(tablesNormally, "省略時は前回の取り込み設定を使います。初回は指定セル・自動判定です。");
        }
        else { layout.Controls.Add(urlLabel, 0, 2); layout.Controls.Add(urls, 0, 3); }
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var cancel = new Button { Text = "キャンセル", AutoSize = true, DialogResult = DialogResult.Cancel };
        var apply = new Button { Text = settingsMode ? "保存" : "挿入", AutoSize = true };
        buttons.Controls.Add(cancel); buttons.Controls.Add(apply); layout.Controls.Add(buttons, 0, 4); Controls.Add(layout);
        AcceptButton = apply; CancelButton = cancel;
        tips.SetToolTip(urls, "HTTPSの画像ファイルを直接取得できるURLを画像と同じ順に指定します。認証が必要なURLはExcelで表示できません。");
        tips.SetToolTip(embedded, "受信画像をブックへ保存します。送信元の一時ファイルが削除されても表示できます。");
        tips.SetToolTip(linked, "Explorerからの画像のみ選択できます。元ファイルを移動・削除すると参照できなくなる場合があります。");
        void UpdateUrl() { urls.Enabled = formula.Checked; urlLabel.Enabled = formula.Checked; }
        formula.CheckedChanged += (_, _) => UpdateUrl(); UpdateUrl();
        apply.Click += (_, _) =>
        {
            var items = urls.Lines.Select(s => s.Trim()).Where(s => s.Length != 0).ToArray();
            if (!settingsMode && formula.Checked && (items.Length != fileCount || items.Any(s => !Uri.TryCreate(s, UriKind.Absolute, out var uri)
                || uri.Scheme != Uri.UriSchemeHttps || !string.IsNullOrEmpty(uri.UserInfo))))
            {
                MessageBox.Show(this, $"認証情報を含まないHTTPS画像URLを、画像の順に{fileCount}件指定してください。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Options = new InsertOptions(drop.Checked ? Destination.DropPosition : selection.Checked ? Destination.SelectionPosition : Destination.SelectionRange, formula.Checked ? PictureKind.ImageFormula : linked.Checked ? PictureKind.Linked : PictureKind.Embedded, items);
            DialogResult = DialogResult.OK; Close();
        };
    }
    protected override void Dispose(bool disposing) { if (disposing) tips.Dispose(); base.Dispose(disposing); }
}
