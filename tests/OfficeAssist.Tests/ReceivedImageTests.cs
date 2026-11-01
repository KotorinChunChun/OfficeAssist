using OfficeAssist;
using System.Reflection;

internal static partial class Program
{
    static void ReceivedImageTests(string output, string image, string invalid)
    {
        string root = Path.Combine(output, "received-tests");
        Directory.CreateDirectory(root);
        string temporary = Path.Combine(output, "temporary-source.png");
        File.Copy(image, temporary, true);
        var batch = ReceivedImageBatch.Capture([temporary], root);
        string owned = batch.Images.Single();
        Check(batch.Originals.Single() == temporary && owned != temporary, "受信元と所有画像を区別");
        File.Delete(temporary); // 送信元の1分後削除を待たず再現する。
        using (var bitmap = new Bitmap(owned)) Check(bitmap.Size == new Size(240, 120), "送信元が削除されても挿入用画像はデコード可能");
        batch.Dispose(); batch.Dispose();
        Check(!File.Exists(owned) && Directory.GetDirectories(root).Length == 0, "完了・キャンセルで自分の受信画像だけを片付けられる");
        using (var saved = ReceivedImageBatch.Capture([image], root))
            Check(saved.Originals.Single() == image && saved.Images.Single() != image, "Explorerリンク用の原本パスを保持");
        Check(File.Exists(image), "受信整理で原本を削除しない");
        Reject(() => ReceivedImageBatch.Capture([image, invalid], root), "壊れた受信画像をCopy応答前に拒否");
        Check(Directory.GetDirectories(root).Length == 0 && File.Exists(image) && File.Exists(invalid), "一括検証失敗で所有コピーだけを回収");
        bool failed = false;
        try { using var missing = ReceivedImageBatch.Capture([image, Path.Combine(output, "absent.png")], root); }
        catch (IOException) { failed = true; }
        Check(failed && Directory.GetDirectories(root).Length == 0, "コピー途中の消失でも一部の受信ファイルを残さない");

        using var externalDialog = new InsertOptionsForm("$A$1", 1, false);
        using var explorerDialog = new InsertOptionsForm("$A$1", 1, true);
        Check(!externalDialog.Controls.Find("LinkedPictureOption", true).Single().Enabled,
            "外部アプリの短命なファイルへリンクしない");
        Check(explorerDialog.Controls.Find("LinkedPictureOption", true).Single().Enabled,
            "Explorerの既存リンク挿入を維持");

        using var overlay = new DropOverlay();
        ReceivedImageBatch? accepted = null;
        bool accept = false;
        overlay.FilesDropped += (paths, _, _) =>
        {
            if (!accept) return false;
            accepted = ReceivedImageBatch.Capture(paths, root);
            Check(File.Exists(accepted.Images.Single()), "Copy応答より先に画像を確保");
            return true;
        };
        var data = new DataObject(); data.SetData(DataFormats.FileDrop, false, new[] { image });
        void Enter() => typeof(Control).GetMethod("OnDragEnter", BindingFlags.NonPublic | BindingFlags.Instance)!
            .Invoke(overlay, [new DragEventArgs(data, 1, 10, 20, DragDropEffects.Copy | DragDropEffects.Move, DragDropEffects.None)]);
        DragEventArgs Drop()
        {
            var args = new DragEventArgs(data, 0, 10, 20, DragDropEffects.Copy | DragDropEffects.Move, DragDropEffects.Move);
            typeof(Control).GetMethod("OnDragDrop", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(overlay, [args]);
            return args;
        }
        Enter(); Check(Drop().Effect == DragDropEffects.None, "画像を確保できなければCopy成功を応答しない");
        accept = true; Enter();
        Check(Drop().Effect == DragDropEffects.Copy && accepted is not null, "画像を確保した場合だけCopyを応答");
        accepted?.Dispose();

        // 実際のDropから呼ばれるサービス経路でもInvalidDataExceptionをOLEへ漏らさない。
        using var dispatcher = new Control();
        int reports = 0;
        using var service = new ExcelDropService(dispatcher, (_, isFailure) => { if (isFailure) reports++; }, startTimer: false);
        var queue = typeof(ExcelDropService).GetMethod("QueueInsert", BindingFlags.NonPublic | BindingFlags.Instance)!;
        Check(queue.Invoke(service, [new[] { invalid }, Point.Empty, false]) is false && !service.Busy && reports == 1,
            "壊れた画像はサービス経路でも例外を漏らさず受信拒否");
        string empty = Path.Combine(output, "empty.png"); File.WriteAllBytes(empty, []);
        Check(queue.Invoke(service, [new[] { empty }, Point.Empty, false]) is false && !service.Busy && reports == 2,
            "空画像の検証例外でもCopy成功や処理中状態にしない");
    }
}
