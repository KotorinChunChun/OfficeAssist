namespace OfficeAssist;

internal sealed class TableLoadForm<T> : BrandedForm
{
    readonly CancellationTokenSource cancellation = new();
    internal T? Result { get; private set; }
    internal Exception? Error { get; private set; }
    internal TableLoadForm(string title, Func<CancellationToken, T> load)
    {
        Text = WindowTitle.Create(title); ClientSize = new(520, 110); ShowInTaskbar = false; MinimizeBox = MaximizeBox = false; TopMost = true;
        var progress = new ProgressBar { Dock = DockStyle.Top, Style = ProgressBarStyle.Marquee, Height = 25 };
        var cancel = new Button { Text = "キャンセル", Dock = DockStyle.Bottom }; cancel.Click += (_, _) => cancellation.Cancel(); Controls.AddRange([progress, cancel]);
        bool working = true;
        FormClosing += (_, e) => { if (working) { cancellation.Cancel(); e.Cancel = true; } };
        Shown += async (_, _) =>
        {
            try { Result = await Task.Run(() => load(cancellation.Token)); DialogResult = DialogResult.OK; }
            catch (OperationCanceledException) { DialogResult = DialogResult.Cancel; }
            catch (Exception ex) { Error = ex; DialogResult = DialogResult.Abort; }
            finally { working = false; Close(); }
        };
    }
    protected override void Dispose(bool disposing) { if (disposing) cancellation.Dispose(); base.Dispose(disposing); }
}
