# 開発者ガイド

.NET 10 SDKとWindows 11 x64を使用します。OfficeのInterop NuGetパッケージは使わず、実行時にインストール済みExcelのCOMインターフェースへ接続します。

```powershell
cd C:\develop\KotorinToys\tools\OfficeAssist
dotnet build src/OfficeAssist/OfficeAssist.csproj
dotnet run --project src/OfficeAssist/OfficeAssist.csproj
```

管理用ビルドは`dotnet build src/OfficeAssist/OfficeAssist.csproj --artifacts-path artifacts/management`です。通常起動と`--startup`はいずれも画面を表示せず常駐します。二重起動では既存画面を開かずに終了します。管理窓口の`show`または通知領域メニューの「開く」で画面を表示できます。この通常起動は利用者の2026-10-28の明示指示による共通起動規約の例外です。

管理窓口はKotorinToys管理契約2（status/show/stop）です。Integration/ManagementEndpoint.csは親のcontracts/ManagementEndpoint.Resilient.csのコピーです。独自アイコンは`src/OfficeAssist/assets/`に同梱しており、親リポジトリがなくてもビルドできます。

独自アイコンのマスターは`src/OfficeAssist/assets/OfficeAssist.png`です。`./dev/scripts/build-app-icon.ps1`で正方形マスターの周囲の透明余白を調整し、16～256pxの10サイズを含むICOを再生成できます。

バージョンごとの結果は[開発資料](../dev/README.md)を参照してください。

## 検証

### プレビューの性能測定

`--preview-benchmark`は合成データだけを使い、機能別に3回測定して`benchmark.json`へ全試行・中央値・割当量・処理結果SHA-256・実行環境・測定DLLのSHA-256を記録します。保存先は毎回未使用のフォルダーを指定してください。`--quick`は動作確認用の1回測定で、性能判断には使いません。

```powershell
cd C:\develop\KotorinToys\tools\OfficeAssist
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release -- --preview-benchmark dev/history/v0.2.2/logs/preview-benchmark/run-a
./dev/scripts/Compare-PreviewBenchmarks.ps1 -Before dev/history/v0.2.2/logs/preview-benchmark/run-a/benchmark.json -After dev/history/v0.2.2/logs/preview-benchmark/run-b/benchmark.json -OutputDirectory dev/history/v0.2.2/logs/preview-benchmark/comparison
```

改修前後の測定結果を同じ環境で別々に保存して比較します。元データのみ／書式あり、対象属性なし／あり、検査、推定、行加工・結合、表示行の選択、文字コード自動／手動、繰返し／固有値、日付／不正日付、長文、縦長／横長、初回／小数桁変更を含みます。`features/`では、結果キャッシュ、範囲設定、短い文字列の処理を無効にした場合を測定し、影響を切り分けます。共通ケースの条件と結果が違う場合、比較スクリプトは失敗します。

書式の測定は結果文字列の構築を含みます。`excel/cold-start`だけは専用Excelの起動込み、`form/`は専用Excel起動後の非表示コントロールの構築と表示完了までです。表示結果の再利用が効く反復の中央値と、未登録の値を含む最初の試行を分けて評価してください。`--preview-final-check --selection-com --large-preview`は性能対策の一致・上限・終了処理と既存機能をまとめてCLI確認します。先に大型サンプルを同じ出力先の`files/`へ生成します。GUIの実描画・操作の受入確認は別です。

書式付きプレビューは次のコマンドで確認します。先に示すコマンドは専用の非表示Excelを使い、実際に書き込んだセルとプレビューのText・Value2・NumberFormat・PrefixCharacterを比較します。取り込み先の列幅、書体、1904年日付方式、地域設定も確認します。続くコマンドは画面を表示せずにコントロールを試験し、形式の変更、エラーからの復帰、全件検査、遅れて届いた結果の破棄を確認します。GUIの実描画やドラッグ操作は別に検証します。

```powershell
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release -- --format-preview-com dev/history/v0.2.2/logs/format-check
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release -- --format-preview-ui dev/history/v0.2.2/logs/format-check
```

大型プレビュー用サンプルは次のコマンドで作成します。40列と64列、それぞれ見出し＋1500行で、２ファイル合計156104セルです。番号・日付・数値・日本語・改行・引用符・数式風文字列・空欄を含み、既存ファイルは上書きしません。

```powershell
./dev/scripts/New-TabularPreviewSamples.ps1 -OutputDirectory ./dev/history/v0.2.2/logs/large-table-20261105/files
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release -- --large-preview dev/history/v0.2.2/logs/large-table-20261105
```

CSV取り込みCLIは常駐アプリの排他制御とは独立して新規Excelプロセスを作り、新しいブックへ展開します。全値を読み戻して照合し、`--show`指定時はブックを閉じずに表示します。保存先と報告先が既に存在する場合は上書きしません。

```powershell
dotnet run --project src/OfficeAssist/OfficeAssist.csproj -c Release -- --import-table --new-book --header auto --show --save-as C:\path\確認.xlsx --report C:\path\result.json -- C:\path\source.csv
```

`--header`は`auto`・`yes`・`no`を指定します。列形式を上書きする場合は`--types text-format,date,number,general,text-quote`のように全列分をカンマ区切りで指定します。複数ファイルの場合も同じ列数・形式を適用します。未指定なら各ファイルの列から推定します。`--save-as`を省略した`--show`も使用できます。`--show`を省略する場合は保存先を必須とし、保存・読戻し後に専用Excelを終了します。報告には照合結果、ブックのパス、PID、件数と列形式を記録します。`--report-values`指定時だけ先頭10行・20列の値・書式・PrefixCharacterを含めます。個人データを使う場合は報告も公開しないでください。

追加の列形式は`date-weekday,datetime,time,grouped-number,percent,yen,duration`です。`--decimals 0,2,0`のように全列分の桁数（0～15）を指定できます。省略時は値から小数桁を推定します。報告には小数桁を含み、`--report-values`指定時は実際の表示文字列も含みます。

v0.2.2では`--skip-first N --skip-last N --drop-empty-rows`で各ファイルを加工し、`--merge --merge-by order|name`で結合します。列名結合は`--header yes`が必要です。`--preview none|first10|all`はGUIとの設定比較用で、非表示CLIはプレビューを生成しません。出力の全件読戻しは常に行います。レポートは既定で件数・列形式等のメタデータだけとし、値・書式・PrefixCharacter等のサンプルが必要な場合は`--report-values`を明示します。

`--list-inputs -- C:\path\folder`でフォルダ内の候補と探索エラーをJSONで出力します。`--input-list C:\path\selected.json`には、フルパスの文字列を取り込み順に並べたJSON配列を指定できます。`--include-pattern '*2026*.csv'`は取り込む対象を絞ります。GUIの検索で絞るのは一覧表示だけなので、用途に応じて使い分けてください。CLIのフォルダ探索で未解決のエラーがあれば書込みを中止します。

`--list-excel-targets`で実行中ExcelのPID・bookId・シート名を取得できます。既存ブックへのCLI書込みは`--excel-instance PID --book-id ID`を必須とし、`--new-sheet`、`--append-same-sheet`、または`--sheet 名前 --cell B3`で対象を指定します。追記と結合は排他です。試験には専用Excelプロセスと合成データを使用してください。

文字コードは`--encoding auto|utf8|utf16le|utf16be|cp932`で指定します。既定はautoです。ファイル別指定は`--encoding-map C:\path\encodings.json`を使い、JSONオブジェクトのキーを入力のフルパス、値を文字コード名とします。mapで指定されていないファイルには`--encoding`を適用します。レポートのencodingsに指定値と実際のデコーダー名を記録します。

行除外・結合・追記・文字コードの検証用ファイル20ケースは、次のコマンドで生成できます。UTF-8（BOM有無）、UTF-16 LE/BE（BOM有無）、CP932、引用符内改行、空欄と空白、列順違い、重複見出し、同名別パス、不正UTF-8等を含み、manifest.jsonで用途とSHA-256を確認できます。rejectフォルダは拒否の確認用です。

```powershell
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release -- --create-import-cases dev/history/v0.2.2/logs/new-cases
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release -- --selection-com dev/history/v0.2.2/logs/encoding-check
```

`--selection-com`は専用の非表示Excelで、既存データとの衝突、新規ブック・新規シート、列形式と引用符の有無を検証します。

ドロップ受付のタイミングは `--drop-timing` で検証します。カーソルの早い移動、解放と押下の連続、遅いOLE通知、キャンセル後の次回受付を試験し、Windowsフックを専用スレッドに登録・解除します。`--selection-com`では新しいブックの作成前に元のロックを解除する順序と、新しいブックへの書込み・操作の復元も確認します。実ドラッグは `--fixture` を使うGUI受入項目です。

```powershell
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release -- --drop-timing dev/history/v0.2.2/logs/drop-timing-check
```

```powershell
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release -- --selection-com dev/history/v0.2.2/logs/column-check
```

```powershell
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -- dev/history/v0.1.0/logs/unit
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -- --excel dev/history/v0.1.0/logs/com
```

`--excel`は新しいExcelプロセスで専用ブックを作り、試験後に保存・終了します。既存ExcelのPIDと一致したら中止します。GUI試験用の`--fixture`は専用Excelを最大25分間開き、出力先へ`stop-fixture`ファイルを作成すると試験ブックを保存して終了します。毎回新しい出力先を指定してください。

### サンプル画像の挿入・保存をCLIで検証する

`--excel --picture-case`は指定した試験用`.xlsx`を複製し、指定した画像を埋め込みます。専用の非表示Excelで動作し、セル番地から通常の画像配置処理を呼び出します。次の例ではケースJSONと同じフォルダーに試験ブックと画像を置きます。絶対パスも指定できます。

```json
{
  "Workbook": "取り込み結果.xlsx",
  "Sheet": "売上サンプル",
  "Images": ["01-blue.png", "02-orange.png"],
  "Destination": "selection-range",
  "Selection": "E2:H9,J2:M9"
}
```

`Destination`は `drop`（ドロップ位置に原寸）、`selection-position`（選択範囲の左上に原寸）、`selection-range`（選択範囲にフィット）から選びます。`drop`では `"DropCell": "E2"` のように1セルを追加指定します。`Selection`は実行時の選択を表し、複数範囲はカンマで区切ります。

```powershell
cd C:\develop\KotorinToys\tools\OfficeAssist
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release -- --excel --picture-case C:\path\picture-case.json dev/history/v0.2.2/logs/picture-case-01
```

出力先には `picture-result.xlsx` と `picture-report.json` が作られます。JSONには画像ごとの期待値と実測値（左・上・幅・高さ、単位はポイント）、警告、複製元と結果のSHA-256を記録します。画像配置と埋め込みを照合し、保存後に開き直して図形の保持と複製元のハッシュを確認できた場合に `verified: true` を返します。出力名が使用済みの場合は別のフォルダーを指定してください。

エクスプローラーからの実ドラッグ、修飾キー、ダイアログの表示は、`--fixture`を使うGUI受入項目として別に確認します。

通常試験ではExplorer以外の外部アプリから始めた操作、Excel内の操作とOfficeAssist自身の除外、Copy可能な画像FileDropの判定を確認します。画像以外の入力と送信元の読み取りで例外が起きる場合の回帰テストも含みます。実入力・クリップボード・既存Excelは操作しません。次のコマンドはPR用に出力を分ける例です。

```powershell
dotnet build tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release --artifacts-path artifacts/external-drop-pr -warnaserror
dotnet artifacts/external-drop-pr/bin/OfficeAssist.Tests/release/OfficeAssist.Tests.dll dev/history/v0.2.2/logs/unit
```

管理用ビルド後、OfficeAssistが停止している状態で次を実行すると、無表示起動・状態・誤PID拒否・不正要求後の受付・二重起動・正常停止を確認します。起動中のものがあれば試験起動を中止します。

```powershell
./dev/scripts/test-management.ps1 -Executable ./artifacts/management/bin/OfficeAssist/debug/OfficeAssist.exe -StartForTest
```


## README用の安全な画面プレビュー

`OfficeAssist.Tests.exe --manual-preview <main|settings|image|table> <専用出力先>`で実際のフォームを表示します。利用者設定・常駐サービス・実Excelには接続しません。保存や挿入を確定しても、製品設定やブックには書き込みません。画像や表の結果はCLIで用意した専用ブックで別に確認します。このコマンドはテストEXE専用です。
