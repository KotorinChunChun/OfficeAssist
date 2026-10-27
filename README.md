# OfficeAssist

Officeをカスタマイズして便利に使うためのWindows用常駐アプリです。

## Excelへ画像を挿入する

1. OfficeAssistを起動し、「Excelへの画像ドロップ」をONにします（初期値ON）。
2. デスクトップ版Excelで編集するブックを開きます。
3. エクスプローラーのフォルダーから画像をシートへドラッグ＆ドロップします。

ドロップしたセルの左上を基準に画像を埋め込みます。複数画像は縦に並びます。元画像は移動・変更しません。対応形式はPNG、JPEG、BMP、GIF、TIFFです。

画面を閉じると通知領域に常駐します。通知領域のメニューから再表示・一時停止・終了できます。Windowsへの自動起動登録は行いません。

## 対応環境

Windows 11 x64、.NET Desktop Runtime 10 x64、インストール済みのデスクトップ版Microsoft Excelが必要です。ExcelとOfficeAssistは通常権限で実行してください。Web版Excelは対象外です。

- [使い方と制限](docs/guide.md)
- [開発者ガイド](docs/development.md)
