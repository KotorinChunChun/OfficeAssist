# OfficeAssist

Officeをカスタマイズして便利に使うためのWindows用常駐アプリです。

## Excelへ画像を挿入する

1. OfficeAssistを起動します。通知領域に常駐し、画像ドロップは初期状態でONです。
2. デスクトップ版Excelで編集するブックを開きます。
3. エクスプローラーのフォルダーや、SSCaptureなどの画像ファイルをドラッグできるアプリから、画像をシートへドラッグ＆ドロップします。

通常はダイアログを表示せず、選択範囲に収まる大きさで画像を埋め込みます。修飾キー（初期値はShift）を押しながらドロップすると、挿入先と形式を選び、次回以降の既定として保存できます。挿入先はドロップ位置（原寸大）・選択位置（原寸大）・選択範囲から選べます。リンクされた図はExplorerから渡した画像だけが対象です。元画像は移動・変更しません。対応形式はPNG、JPEG、BMP、GIF、TIFFです。

ウィンドウを閉じても通知領域に常駐します。通知領域のアイコンを右クリックすると、画面表示・一時停止・Excelドロップ設定・終了ができます。設定画面を開く修飾キーはShift・Ctrl・Altから選べます。Windowsへの自動起動登録は行いません。

## インストール

[GitHub Releases](https://github.com/KotorinChunChun/OfficeAssist/releases)からセットアップファイルをダウンロードして実行します。v0.2.0以降のセットアップには.NETランタイムを同梱しています。

## 対応環境

Windows 11 x64とインストール済みのデスクトップ版Microsoft Excelが必要です。ExcelとOfficeAssistは通常権限で実行してください。Web版Excelは対象外です。

- [使い方と制限](docs/guide.md)
- [開発者ガイド](docs/development.md)
