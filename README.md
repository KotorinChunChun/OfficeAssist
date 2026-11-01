# OfficeAssist

Officeをカスタマイズして便利に使うためのWindows用常駐アプリです。

## Excelへ画像を挿入する

1. OfficeAssistを起動します。通知領域に常駐し、画像ドロップは初期状態でONです。
2. デスクトップ版Excelで編集するブックを開きます。
3. エクスプローラーのフォルダーや、SSCaptureなどの画像ファイルをドラッグできるアプリから、画像をシートへドラッグ＆ドロップします。

画像をドロップすると挿入方法を選べます。挿入先はドロップ位置またはExcelで選択した範囲、サイズは原寸大または選択範囲、形式は埋め込み・リンクされた図・セル内IMAGE関数式を選択できます。リンクされた図はExplorerから渡した画像だけが対象です。元画像は移動・変更しません。対応形式はPNG、JPEG、BMP、GIF、TIFFです。

通知領域のアイコンを右クリックすると、一時停止・画面表示・終了ができます。Windowsへの自動起動登録は行いません。

## インストール

[GitHub Releases](https://github.com/KotorinChunChun/OfficeAssist/releases)からセットアップファイルをダウンロードして実行します。v0.2.0以降のセットアップには.NETランタイムを同梱しています。

## 対応環境

Windows 11 x64とインストール済みのデスクトップ版Microsoft Excelが必要です。ExcelとOfficeAssistは通常権限で実行してください。Web版Excelは対象外です。

- [使い方と制限](docs/guide.md)
- [開発者ガイド](docs/development.md)
