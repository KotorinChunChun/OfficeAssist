using System.Reflection;
using OfficeAssist;

internal static partial class Program
{
    static void DropTimingTests()
    {
        var gesture = new DropGesture();
        var explorer = new Point(10, 20);
        var excel = new Point(600, 400);
        static nint WindowAt(Point _) => 200;

        gesture.RecordPress(explorer, 100);
        Check(gesture.TakePress(true, excel, WindowAt) == new DropPress(explorer, 100),
            "タイマーの前にカーソルがExcelへ移動しても押下元を使う");
        gesture.RecordPress(explorer, 101);
        Check(gesture.HasPendingPress && gesture.TakePress(true, excel, WindowAt) == new DropPress(explorer, 101),
            "タイマーがボタン解放を取りこぼしても次の押下で受付を更新する");
        Check(!gesture.HasPendingPress && gesture.TakePress(true, excel, WindowAt) is null,
            "同じ押下を重複して処理しない");
        gesture.TakePress(false, excel, WindowAt);
        Check(gesture.WaitForDrop(100, true) && gesture.WaitForDrop(200, true) && !gesture.WaitForDrop(1100, true),
            "OLE受信中は解放後の遅い通知を待ち1秒で打ち切る");
        Check(gesture.TakePress(true, excel, WindowAt) == new DropPress(excel, 200), "押下通知が来ない場合もポーリングで受付を続ける");
        gesture.TakePress(false, excel, WindowAt);
        Check(gesture.WaitForDrop(2000, false) && !gesture.WaitForDrop(2045, false), "OLE受信前の待機は短時間で解除する");

        using var overlay = new DropOverlay();
        int received = 0;
        overlay.FilesDropped += (_, _, _) => { received++; return true; };
        var data = new DataObject(DataFormats.FileDrop, new[] { "fixture.csv" });
        var enter = typeof(Control).GetMethod("OnDragEnter", BindingFlags.NonPublic | BindingFlags.Instance)!;
        var drop = typeof(Control).GetMethod("OnDragDrop", BindingFlags.NonPublic | BindingFlags.Instance)!;
        DragEventArgs Event() => new(data, 0, 10, 20, DragDropEffects.Copy, DragDropEffects.None);
        enter.Invoke(overlay, [Event()]);
        Check(overlay.Receiving, "CSVのDragEnterで受信中になる");
        overlay.Reset();
        var stale = Event();
        drop.Invoke(overlay, [stale]);
        Check(!overlay.Receiving && received == 0 && stale.Effect == DragDropEffects.None, "中止したドラッグの遅いDropを受け付けない");
        enter.Invoke(overlay, [Event()]);
        var next = Event();
        drop.Invoke(overlay, [next]);
        Check(!overlay.Receiving && received == 1 && next.Effect == DragDropEffects.Copy, "中止の直後でも次のCSVを受け付ける");
    }
}
