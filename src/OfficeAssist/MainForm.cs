namespace OfficeAssist;

internal sealed class MainForm : BrandedForm
{
    readonly CheckBox enabled = new() { Text = "Excelへの画像ドロップ", Checked = true, AutoSize = true };
    readonly Label result = new() { Text = "待機中", AutoSize = false, Dock = DockStyle.Fill };
    readonly ToolTip tips = new();
    public event Action<bool>? FeatureChanged;
    public event Action? ExitRequested;
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    internal bool FeatureEnabled { get => enabled.Checked; set => enabled.Checked = value; }
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    internal bool AllowExit { get; set; }

    public MainForm()
    {
        Text = WindowTitle.Create();
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(440, 190);
        MinimumSize = new Size(400, 220);
        Font = new Font("Yu Gothic UI", 10);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), RowCount = 3, ColumnCount = 1 };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
        var exit = new Button { Text = "終了", AutoSize = true, Anchor = AnchorStyles.Right };
        layout.Controls.Add(enabled, 0, 0);
        layout.Controls.Add(result, 0, 1);
        layout.Controls.Add(exit, 0, 2);
        Controls.Add(layout);
        tips.SetToolTip(enabled, "フォルダーからExcelのシートへ画像をドラッグ＆ドロップして挿入します。");
        tips.SetToolTip(exit, "OfficeAssistの常駐を終了します。");
        tips.SetToolTip(result, "直近の操作結果です。ウィンドウの×ボタンでは通知領域に常駐します。");
        enabled.CheckedChanged += (_, _) => FeatureChanged?.Invoke(enabled.Checked);
        exit.Click += (_, _) => ExitRequested?.Invoke();
        FormClosing += (_, e) =>
        {
            if (!AllowExit && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); }
        };
    }

    public void SetResult(string message) => result.Text = message;
    protected override void Dispose(bool disposing)
    {
        if (disposing) tips.Dispose();
        base.Dispose(disposing);
    }
}
