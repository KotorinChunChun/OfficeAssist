namespace OfficeAssist;

internal enum PictureKind { Embedded, Linked, ImageFormula }
internal sealed record InsertOptions(bool UseSelection, bool FitSelection, PictureKind Kind, string[] Urls);

internal sealed class InsertOptionsForm : BrandedForm
{
    readonly RadioButton drop = new() { Text = "ドロップ位置", AutoSize = true, Checked = true };
    readonly RadioButton selection = new() { Text = "選択範囲", AutoSize = true };
    readonly RadioButton original = new() { Text = "原寸大", AutoSize = true, Checked = true };
    readonly RadioButton fit = new() { Text = "選択範囲", AutoSize = true };
    readonly RadioButton embedded = new() { Text = "埋め込み", AutoSize = true, Checked = true };
    readonly RadioButton linked = new() { Text = "リンクされた図", AutoSize = true };
    readonly RadioButton formula = new() { Text = "セル内 IMAGE 関数式", AutoSize = true };
    readonly TextBox urls = new() { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
    readonly Label urlLabel = new() { Text = "画像の HTTPS URL（1画像につき1行）", AutoSize = true };
    readonly ToolTip tips = new();
    readonly int fileCount;
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    internal InsertOptions? Options { get; private set; }

    internal InsertOptionsForm(string selectedAddress, int count, bool allowLinkedImages = true)
    {
        fileCount = count;
        linked.Name = "LinkedPictureOption";
        linked.Enabled = allowLinkedImages;
        Text = WindowTitle.Create("画像の挿入");
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
        ClientSize = new Size(500, 395);
        Font = new Font("Yu Gothic UI", 10);
        TopMost = true;
        selection.Text = "選択範囲 " + selectedAddress;
        static GroupBox Group(string text, params Control[] controls)
        {
            var group = new GroupBox { Text = text, Dock = DockStyle.Fill, Padding = new Padding(8) };
            var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, AutoScroll = true };
            flow.Controls.AddRange(controls); group.Controls.Add(flow); return group;
        }
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 6 };
        foreach (int height in new[] { 62, 62, 62, 26, 0, 42 })
            layout.RowStyles.Add(height == 0 ? new RowStyle(SizeType.Percent, 100) : new RowStyle(SizeType.Absolute, height));
        layout.Controls.Add(Group("挿入先", drop, selection), 0, 0);
        layout.Controls.Add(Group("サイズ", original, fit), 0, 1);
        layout.Controls.Add(Group("挿入形式", embedded, linked, formula), 0, 2);
        layout.Controls.Add(urlLabel, 0, 3); layout.Controls.Add(urls, 0, 4);
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var cancel = new Button { Text = "キャンセル", AutoSize = true, DialogResult = DialogResult.Cancel };
        var apply = new Button { Text = "挿入", AutoSize = true };
        buttons.Controls.Add(cancel); buttons.Controls.Add(apply); layout.Controls.Add(buttons, 0, 5); Controls.Add(layout);
        AcceptButton = apply; CancelButton = cancel;
        tips.SetToolTip(urls, "HTTPSの画像ファイルを直接取得できるURLを画像と同じ順に指定します。認証が必要なURLはExcelで表示できません。");
        tips.SetToolTip(embedded, "受信画像をブックへ保存します。送信元の一時ファイルが削除されても表示できます。");
        tips.SetToolTip(linked, "Explorerからの画像のみ選択できます。元ファイルを移動・削除すると参照できなくなる場合があります。");
        void UpdateUrl() { urls.Enabled = formula.Checked; urlLabel.Enabled = formula.Checked; }
        formula.CheckedChanged += (_, _) => UpdateUrl(); UpdateUrl();
        apply.Click += (_, _) =>
        {
            var items = urls.Lines.Select(s => s.Trim()).Where(s => s.Length != 0).ToArray();
            if (formula.Checked && (items.Length != fileCount || items.Any(s => !Uri.TryCreate(s, UriKind.Absolute, out var uri)
                || uri.Scheme != Uri.UriSchemeHttps || !string.IsNullOrEmpty(uri.UserInfo))))
            {
                MessageBox.Show(this, $"認証情報を含まないHTTPS画像URLを、画像の順に{fileCount}件指定してください。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Options = new InsertOptions(selection.Checked, fit.Checked, formula.Checked ? PictureKind.ImageFormula : linked.Checked ? PictureKind.Linked : PictureKind.Embedded, items);
            DialogResult = DialogResult.OK; Close();
        };
    }
    protected override void Dispose(bool disposing) { if (disposing) tips.Dispose(); base.Dispose(disposing); }
}
