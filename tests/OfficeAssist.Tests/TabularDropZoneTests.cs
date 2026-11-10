using OfficeAssist;
using System.Reflection;

internal static partial class Program
{
    static void TabularDropZoneTests()
    {
        var sheet = Rectangle.FromLTRB(1744, 249, 2598, 696);
        const int bottomBand = 31;
        ImportDestination? At(int x, int y, bool hasCell = false) =>
            ExcelTableDropSite.PreferredDestination(ExcelTableDropSite.Zone(new(x, y), sheet, bottomBand), hasCell);

        Check(At(1800, 300, true) is null && new TabularDropProfile(ImportDestination.Selection).ForNewDrop(At(1800, 300, true)).Destination == ImportDestination.Selection,
            "実セルへのドロップでは指定セルを含む従来の設定を使う");
        Check(At(1800, 665) == ImportDestination.NewSheet && At(2430, 680) == ImportDestination.NewSheet && At(1900, 704) == ImportDestination.NewSheet,
            "シートタブ・横スクロールバー・ステータスバーは新しいシート");
        Check(At(2000, 60) == ImportDestination.NewBook && At(2000, 224) == ImportDestination.NewBook && At(1800, 300) == ImportDestination.NewBook,
            "リボン・数式バー・セルを特定できない場所は新しいブック");
        Check(At(1800, 664, true) is null && ExcelTableDropSite.Zone(new(1800, 704), Rectangle.Empty, 0) == ExcelTableDropZone.Other,
            "下部の直前のセルを保持し、グリッドがない場合は新しいブック");
        var saved = new TabularDropProfile(ImportDestination.AppendSameSheet);
        Check(saved.ForNewDrop(At(1800, 704)).Destination == ImportDestination.NewSheet && saved.ForNewDrop(At(2000, 60)).Destination == ImportDestination.NewBook,
            "セル以外の位置による既定値は保存済みの展開先より優先");

        using var overlay = new DropOverlay { ImageBounds = sheet };
        int accepted = 0;
        overlay.FilesDropped += (_, _, _) => { accepted++; return true; };
        DragDropEffects DropOnRibbon(string file)
        {
            var data = new DataObject(); data.SetData(DataFormats.FileDrop, false, new[] { file });
            var args = new DragEventArgs(data, 0, 2000, 60, DragDropEffects.Copy, DragDropEffects.None);
            typeof(Control).GetMethod("OnDragEnter", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(overlay, [args]);
            typeof(Control).GetMethod("OnDragDrop", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(overlay, [args]);
            return args.Effect;
        }
        Check(DropOnRibbon("test.png") == DragDropEffects.None && DropOnRibbon("test.tsv") == DragDropEffects.Copy && accepted == 1,
            "リボンへのドロップはTSVを受け付け、画像の対象をシート内に保つ");
    }
}
