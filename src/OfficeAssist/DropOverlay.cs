namespace OfficeAssist;

internal sealed class DropOverlay : BrandedForm
{
    internal event Func<string[], Point, int, bool>? FilesDropped;
    internal event Action? Rejected;
    internal bool Receiving => pendingFiles is not null;
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    internal Rectangle ImageBounds { get; set; }
    string[]? pendingFiles;
    bool tables;

    internal DropOverlay()
    {
        Text = WindowTitle.Create("ファイルドロップ");
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        TopMost = true;
        BackColor = Color.White;
        Opacity = 0.015; // 完全透明ではヒットテストを通過するため微小な不透明度を保つ。
        AllowDrop = true;
        DragEnter += (_, e) =>
        {
            pendingFiles = ExcelDropData.Read(e.Data, e.AllowedEffect);
            if (pendingFiles is not null)
            {
                tables = !ImageFiles.Accepts(pendingFiles);
                e.Effect = CanDrop(e) ? DragDropEffects.Copy : DragDropEffects.None;
            }
            else { e.Effect = DragDropEffects.None; pendingFiles = null; Rejected?.Invoke(); }
        };
        DragOver += (_, e) => e.Effect = CanDrop(e) ? DragDropEffects.Copy : DragDropEffects.None;
        DragLeave += (_, _) => Reset();
        DragDrop += (_, e) =>
        {
            var files = pendingFiles;
            try
            {
                e.Effect = DragDropEffects.None;
                if (files is not null && CanDrop(e)
                    && FilesDropped?.Invoke(files, new Point(e.X, e.Y), e.KeyState) == true)
                    e.Effect = DragDropEffects.Copy;
            }
            finally { Reset(); }
        };
    }

    bool CanDrop(DragEventArgs e) => pendingFiles is not null && (e.AllowedEffect & DragDropEffects.Copy) != 0
        && (tables || ImageBounds.Contains(e.X, e.Y));

    internal void Reset()
    {
        pendingFiles = null;
        if (Visible) Hide();
    }

    protected override bool ShowWithoutActivation => true;
    protected override CreateParams CreateParams
    {
        get { var parameters = base.CreateParams; parameters.ExStyle |= 0x08000000 | 0x80; return parameters; }
    }
}
