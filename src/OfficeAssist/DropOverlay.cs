namespace OfficeAssist;

internal sealed class DropOverlay : BrandedForm
{
    internal event Action<string[], Point>? FilesDropped;
    internal event Action? Rejected;
    internal bool Receiving { get; private set; }
    string[]? pendingFiles;

    internal DropOverlay()
    {
        Text = WindowTitle.Create("画像ドロップ");
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        TopMost = true;
        BackColor = Color.White;
        Opacity = 0.015; // 完全透明ではヒットテストを通過するため微小な不透明度を保つ。
        AllowDrop = true;
        DragEnter += (_, e) =>
        {
            Receiving = true;
            pendingFiles = e.Data?.GetData(DataFormats.FileDrop) as string[];
            if (ImageFiles.Accepts(pendingFiles) && (e.AllowedEffect & DragDropEffects.Copy) != 0)
                e.Effect = DragDropEffects.Copy;
            else { e.Effect = DragDropEffects.None; pendingFiles = null; Rejected?.Invoke(); }
        };
        DragOver += (_, e) => e.Effect = pendingFiles is not null && (e.AllowedEffect & DragDropEffects.Copy) != 0
            ? DragDropEffects.Copy : DragDropEffects.None;
        DragLeave += (_, _) => { Receiving = false; pendingFiles = null; Hide(); };
        DragDrop += (_, e) =>
        {
            var files = pendingFiles;
            Receiving = false;
            pendingFiles = null;
            if (files is not null && (e.AllowedEffect & DragDropEffects.Copy) != 0)
            {
                e.Effect = DragDropEffects.Copy;
                FilesDropped?.Invoke(files, new Point(e.X, e.Y));
            }
            Hide();
        };
    }

    protected override bool ShowWithoutActivation => true;
    protected override CreateParams CreateParams
    {
        get { var parameters = base.CreateParams; parameters.ExStyle |= 0x08000000 | 0x80; return parameters; }
    }
}
