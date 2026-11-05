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

版別結果は[開発資料](../dev/README.md)を参照してください。

## 検証

大型プレビュー用サンプルは次のコマンドで作成します。40列と64列、それぞれ見出し＋1500行で、２ファイル合計156104セルです。番号・日付・数値・日本語・改行・引用符・数式風文字列・空欄を含み、既存ファイルは上書きしません。

```powershell
./dev/scripts/New-TabularPreviewSamples.ps1 -OutputDirectory ./dev/history/v0.3.0/logs/large-table-20261105/files
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release -- --large-preview dev/history/v0.3.0/logs/large-table-20261105
```

CSV取り込みCLIは常駐アプリの排他制御とは独立して新規Excelプロセスを作り、新しいブックへ展開します。全値を読み戻して照合し、`--show`指定時はブックを閉じずに表示します。保存先と報告先が既に存在する場合は上書きしません。

```powershell
dotnet run --project src/OfficeAssist/OfficeAssist.csproj -c Release -- --import-table --new-book --header auto --show --save-as C:\path\確認.xlsx --report C:\path\result.json -- C:\path\source.csv
```

`--header`は`auto`・`yes`・`no`を指定します。列形式を上書きする場合は`--types text-format,date,number,general,text-quote`のように全列分をカンマ区切りで指定します。複数ファイルの場合も同じ列数・形式を適用します。未指定なら各ファイルの列から推定します。`--save-as`を省略した`--show`も使用できます。`--show`を省略する場合は保存先を必須とし、保存・読戻し後に専用Excelを終了します。報告には照合結果、ブックのパス、PID、件数と列形式を記録します。`--report-values`指定時だけ先頭10行・20列の値・書式・PrefixCharacterを含めます。個人データを使う場合は報告も公開しないでください。

追加の列形式は`date-weekday,datetime,time,grouped-number,percent,yen,duration`です。`--decimals 0,2,0`のように全列分の桁数（0～15）を指定できます。省略時は値から小数桁を推定します。報告には小数桁を含み、`--report-values`指定時は実際の表示文字列も含みます。

v0.3.0では`--skip-first N --skip-last N --drop-empty-rows`で各ファイルを加工し、`--merge --merge-by order|name`で結合します。列名結合は`--header yes`が必要です。`--preview none|first10|all`はGUIとの設定比較用で、非表示CLIはプレビューを生成しません。出力の全件読戻しは常に行います。レポートは既定で件数・列形式等のメタデータだけとし、値・書式・PrefixCharacter等のサンプルが必要な場合は`--report-values`を明示します。

`--list-inputs -- C:\path\folder`でフォルダ内の候補と探索エラーをJSON出力します。`--input-list C:\path\selected.json`にはフルパス文字列のJSON配列を取込順で指定できます。`--include-pattern '*2026*.csv'`は対象選択条件として作用します。GUIの検索は表示だけを絞るので用途を区別してください。CLIのフォルダ探索で未解決のエラーがあれば書込みを中止します。

`--list-excel-targets`で実行中ExcelのPID・bookId・シート名を取得できます。既存ブックへのCLI書込みは`--excel-instance PID --book-id ID`を必須とし、`--new-sheet`、`--append-same-sheet`、または`--sheet 名前 --cell B3`で対象を指定します。追記と結合は排他です。試験には専用Excelプロセスと合成データを使用してください。

`--selection-com`は専用の非表示Excelで、既存データとの衝突、新規ブック・新規シート、列形式と引用符の有無を検証します。

```powershell
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release -- --selection-com dev/history/v0.3.0/logs/column-check
```

```powershell
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -- dev/history/v0.1.0/logs/unit
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -- --excel dev/history/v0.1.0/logs/com
```

`--excel`は新しいExcelプロセスで専用ブックを作り、試験後に保存・終了します。既存ExcelのPIDと一致したら中止します。GUI試験用の`--fixture`は専用Excelを最大25分間開き、出力先へ`stop-fixture`ファイルを作成すると試験ブックを保存して終了します。毎回新しい出力先を指定してください。

通常試験には、Explorer以外の外部起点、Excel内操作とOfficeAssist自身の除外、Copy可能な画像FileDropの判定、非画像と送信元読取例外の回帰も含みます。実入力・クリップボード・既存Excelは操作しません。PR用に出力を分ける例:

```powershell
dotnet build tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -c Release --artifacts-path artifacts/external-drop-pr -warnaserror
dotnet artifacts/external-drop-pr/bin/OfficeAssist.Tests/release/OfficeAssist.Tests.dll dev/history/v0.3.0/logs/unit
```

管理用ビルド後、OfficeAssistが停止している状態で次を実行すると、無表示起動・状態・誤PID拒否・不正要求後の受付・二重起動・正常停止を確認します。起動中のものがあれば試験起動を中止します。

```powershell
./dev/scripts/test-management.ps1 -Executable ./artifacts/management/bin/OfficeAssist/debug/OfficeAssist.exe -StartForTest
```
