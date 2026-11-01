using OfficeAssist;
using System.Runtime.InteropServices;

internal static partial class Program
{
    static void DropTests()
    {
        const uint own = 100;
        Check(ImageDragOrigin.IsEligible("CabinetWClass", 200, own), "Explorerのフォルダーからの開始を維持");
        Check(ImageDragOrigin.IsEligible("Progman", 200, own) && ImageDragOrigin.IsEligible("WorkerW", 200, own), "デスクトップからの開始を維持");
        Check(ImageDragOrigin.IsEligible("WindowsForms10.Window.8.app.0.1", 300, own), "SSCapture等の外部WinForms通知をExplorer限定で拒否しない");
        Check(ImageDragOrigin.IsEligible("OtherImageApplication", 400, own), "画像データの判断を開始元のアプリ名に依存させない");
        Check(!ImageDragOrigin.IsEligible("XLMAIN", 500, own) && !ImageDragOrigin.IsEligible("xlmain", 501, own), "別Excelプロセスも含めExcel内操作を除外");
        Check(!ImageDragOrigin.IsEligible("WindowsForms10.Window.8", own, own), "OfficeAssist自身からの開始を除外");
        Check(!ImageDragOrigin.IsEligible("", 300, own) && !ImageDragOrigin.IsEligible("ValidWindow", 0, own), "消失窓や不明な開始元を拒否");
        Check(!Native.IsExternalImageDragSource(0), "無効HWNDからの開始を拒否");
        using (var form = new Form())
            Check(!Native.IsExternalImageDragSource(form.Handle), "実HWNDでも自分自身の操作を除外");

        string[] files = [@"C:\fixture\notify.png", @"C:\fixture\other.JPG"];
        var data = new DataObject(); data.SetData(DataFormats.FileDrop, false, files);
        using var preference = new MemoryStream(BitConverter.GetBytes((int)DragDropEffects.Move));
        data.SetData("Preferred DropEffect", false, preference);
        var result = ImageDropData.Read(data, DragDropEffects.Copy | DragDropEffects.Move);
        Check(result is not null && result.SequenceEqual(files), "既定MoveのSSCapture形式でもCopy可能な画像FileDropを受信");
        Check(ImageDropData.Read(data, DragDropEffects.Copy)?.Length == 2, "Copyだけを提示する送信元を受信");
        Check(ImageDropData.Read(data, DragDropEffects.Move) is null && ImageDropData.Read(data, DragDropEffects.Link) is null,
            "Copyできないデータは元画像保護のため拒否");
        files[0] = @"C:\fixture\changed.png";
        Check(result![0] == @"C:\fixture\notify.png", "受信配列を複製し送信元の変更から保護");
        Check(ImageDropData.Read(null, DragDropEffects.Copy) is null, "データなしを拒否");
        var text = new DataObject(); text.SetData(DataFormats.UnicodeText, false, "画像ではないドラッグ");
        Check(ImageDropData.Read(text, DragDropEffects.Copy) is null, "外部アプリの文字列ドラッグを画像扱いしない");
        var mixed = new DataObject(); mixed.SetData(DataFormats.FileDrop, false, new[] { "image.png", "book.xlsx" });
        Check(ImageDropData.Read(mixed, DragDropEffects.Copy) is null, "非画像を含む外部ドラッグを拒否");
        var folder = new DataObject(); folder.SetData(DataFormats.FileDrop, false, new[] { @"C:\fixture\folder" });
        Check(ImageDropData.Read(folder, DragDropEffects.Copy) is null, "フォルダーのドロップを拒否");
        var many = new DataObject(); many.SetData(DataFormats.FileDrop, false, Enumerable.Repeat("image.png", 51).ToArray());
        Check(ImageDropData.Read(many, DragDropEffects.Copy) is null, "外部アプリでも最大50枚を維持");
        var invalid = new DataObject(); invalid.SetData(DataFormats.FileDrop, false, new MemoryStream(new byte[4]));
        Check(ImageDropData.Read(invalid, DragDropEffects.Copy) is null, "FileDropを名乗る非パス配列を拒否");
        Check(ImageDropData.Read(new BrokenDropData(), DragDropEffects.Copy) is null, "OLE送信元の読取例外を受信拒否として処理");
    }

    sealed class BrokenDropData : DataObject
    {
        [Obsolete("IDataObject互換経路の読取失敗を検証するテスト専用です。")]
        public override object? GetData(string format, bool autoConvert) => throw new COMException("送信元終了");
        public override bool GetDataPresent(string format, bool autoConvert) => true;
    }
}
