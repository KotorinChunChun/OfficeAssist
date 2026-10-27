# 開発者ガイド

.NET 10 SDKとWindows 11 x64を使用します。OfficeのInterop NuGetパッケージは使わず、実行時にインストール済みExcelのCOMインターフェースへ接続します。

```powershell
cd C:\develop\KotorinToys\tools\OfficeAssist
dotnet build src/OfficeAssist/OfficeAssist.csproj
dotnet run --project src/OfficeAssist/OfficeAssist.csproj
```

管理用ビルドは`dotnet build src/OfficeAssist/OfficeAssist.csproj --artifacts-path artifacts/management`です。`--startup`を付けると画面を表示せず常駐します。手動の二重起動は既存画面を表示し、無表示の二重起動では画面を開きません。

管理窓口はKotorinToys管理契約2（status/show/stop）です。Integration/ManagementEndpoint.csは親のcontracts/ManagementEndpoint.Resilient.csのコピーです。共通アイコンは本体の採用04をコピーして同梱しており、親リポジトリがなくてもビルドできます。

版別結果は[開発資料](../dev/README.md)を参照してください。
