# 開発者ガイド

.NET 10 SDKとWindows 11 x64を使用します。OfficeのInterop NuGetパッケージは使わず、実行時にインストール済みExcelのCOMインターフェースへ接続します。

```powershell
cd C:\develop\KotorinToys\tools\OfficeAssist
dotnet build src/OfficeAssist/OfficeAssist.csproj
dotnet run --project src/OfficeAssist/OfficeAssist.csproj
```

管理用ビルドは`dotnet build src/OfficeAssist/OfficeAssist.csproj --artifacts-path artifacts/management`です。`--startup`を付けると画面を表示せず常駐します。手動の二重起動は既存画面を表示し、無表示の二重起動では画面を開きません。

管理窓口はKotorinToys管理契約2（status/show/stop）です。Integration/ManagementEndpoint.csは親のcontracts/ManagementEndpoint.Resilient.csのコピーです。独自アイコンは`src/OfficeAssist/assets/`に同梱しており、親リポジトリがなくてもビルドできます。

独自アイコンのマスターは`src/OfficeAssist/assets/OfficeAssist.png`です。`./dev/scripts/build-app-icon.ps1`で16～256pxの10サイズを含むICOを再生成できます。画像生成ツールで作成したマスターを使用しています。

版別結果は[開発資料](../dev/README.md)を参照してください。

## 検証

```powershell
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -- dev/history/v0.1.0/logs/unit
dotnet run --project tests/OfficeAssist.Tests/OfficeAssist.Tests.csproj -- --excel dev/history/v0.1.0/logs/com
```

`--excel`は新しいExcelプロセスで専用ブックを作り、試験後に保存・終了します。既存ExcelのPIDと一致したら中止します。GUI試験用の`--fixture`は専用Excelを最大25分間開き、出力先へ`stop-fixture`ファイルを作成すると試験ブックを保存して終了します。毎回新しい出力先を指定してください。

管理用ビルド後、OfficeAssistが停止している状態で次を実行すると、無表示起動・状態・誤PID拒否・不正要求後の受付・二重起動・正常停止を確認します。起動中のものがあれば試験起動を中止します。

```powershell
./dev/scripts/test-management.ps1 -Executable ./artifacts/management/bin/OfficeAssist/debug/OfficeAssist.exe -StartForTest
```
