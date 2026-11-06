using OfficeAssist;
using System.Text.Json;

internal static partial class Program
{
    static void DropDialogTests(string output)
    {
        string imageSettings = Path.Combine(output, "dialog-settings.json");
        File.WriteAllText(imageSettings, "{\"Destination\":2,\"Kind\":0}");
        var legacy = InsertSettings.LoadBehavior(imageSettings);
        Check(!legacy.ImagesNormally && legacy.TablesNormally && legacy.Modifier == DialogModifier.Shift, "従来の設定は画像通常省略・CSV通常表示・Shiftを維持する");
        foreach (bool normally in new[] { false, true }) foreach (bool pressed in new[] { false, true })
        {
            var behavior = new DropDialogBehavior(DialogModifier.Shift, normally, normally);
            Check(behavior.Show(false, pressed ? 4 : 0) == (normally != pressed) && behavior.Show(true, pressed ? 4 : 0) == (normally != pressed), $"通常表示={normally}・Shift={pressed}で画像とCSVを切り替える");
        }
        Check(!new DropDialogBehavior(DialogModifier.Shift, false, false).Show(true, 8) && new DropDialogBehavior(DialogModifier.Ctrl, false, false).Show(true, 8), "設定した修飾キーだけが表示・省略を反転する");
        InsertSettings.Save(new(Destination.SelectionRange, PictureKind.Embedded, []), DialogModifier.Shift, imageSettings, imagesNormally: true, tablesNormally: false);
        InsertSettings.Save(new(Destination.SelectionPosition, PictureKind.Linked, []), path: imageSettings);
        var saved = InsertSettings.LoadBehavior(imageSettings);
        Check(saved.ImagesNormally && !saved.TablesNormally && InsertSettings.Load(imageSettings).Kind == PictureKind.Linked, "画像の挿入設定を更新しても表示・省略の既定を保持する");
        using (var settings = new InsertOptionsForm("", 0, settingsMode: true, behavior: saved))
        {
            Check(settings.ImagesNormally && !settings.TablesNormally, "設定画面へ画像・CSVそれぞれの既定を復元する");
            ((CheckBox)settings.Controls.Find("TablesDialogNormally", true).Single()).Checked = true;
            Check(settings.TablesNormally, "設定画面から通常表示を変更できる");
        }
        var source = new TabularData("source.csv", [["説明"], ["ID", "金額"], ["001", "12.50"], ["002", "123.4"], ["集計"]]);
        var profile = new TabularDropProfile(Header: true, Rows: new(1, 1), Encoding: TableEncoding.Cp932,
            Columns: [new("source.csv", ["ID", "金額"], [ColumnType.TextFormat, ColumnType.Number], [null, 0])]);
        string tableSettings = Path.Combine(output, "table-settings.json"); TabularDropSettings.Save(profile, tableSettings);
        var loaded = TabularDropSettings.Load(tableSettings);
        Check(loaded.Rows == profile.Rows && loaded.Encoding == TableEncoding.Cp932 && loaded.Header == true, "CSVの行加工・見出し・文字コードを保存して復元する");
        var automatic = loaded.Prepare([source]);
        Check(automatic.Tables[0].RowCount == 3 && automatic.Options.Types[0][0] == ColumnType.TextFormat && automatic.Options.Decimals(0, 1) == 0, "ダイアログなしでも保存した行加工・明示列設定を適用する");
        using (var dialog = new TabularImportForm([source], defaults: loaded))
        {
            Check(dialog.Tables[0].Rows[1][0] == "001" && dialog.TryImport(), "保存設定を復元した非表示の取り込み画面を確定できる");
            Check(JsonSerializer.Serialize(dialog.Options!.Types) == JsonSerializer.Serialize(automatic.Options.Types)
                && JsonSerializer.Serialize(dialog.Options.DecimalPlaces) == JsonSerializer.Serialize(automatic.Options.DecimalPlaces)
                && dialog.Tables[0].Rows.SelectMany(r => r).SequenceEqual(automatic.Tables[0].Rows.SelectMany(r => r)), "画面の有無で取り込む値・型・小数桁が一致する");
            var exported = dialog.CreateDropProfile(TableEncoding.Cp932);
            Check(exported.Rows == profile.Rows && exported.Columns![0].Types[1] == ColumnType.Number && exported.Columns[0].Decimals[1] == 0, "確定画面から次回の既定を取得できる");
        }
        var renamed = loaded.Prepare([source with { Name = "next.csv" }]);
        Check(renamed.Options.Decimals(0, 1) == 0, "ファイル名が違っても同じ列構成なら明示設定を再利用する");
        var changed = new TabularData("changed.csv", [["説明"], ["別列", "金額"], ["003", "1.23"], ["集計"]]);
        Reject(() => loaded.Prepare([changed]), "列構成の異なる入力へ保存した列設定を自動適用しない");
        Reject(() => (loaded with { MixedEncoding = true }).Prepare([source]), "文字コード混在の曖昧な設定でダイアログを省略しない");
        var inferred = new TabularDropProfile().Prepare([new("auto.csv", [["値"], ["12.345"]])]);
        Check(inferred.Options.Types[0][0] == ColumnType.Number && inferred.Options.Decimals(0, 0) == 3, "自動列設定は新しいデータから型と桁数を再推定する");
        string text = File.ReadAllText(tableSettings);
        Check(!text.Contains("001") && !text.Contains("12.50") && !text.Contains("BookId") && !text.Contains("CellAddress"), "既定にはデータ値や固定のExcel書込み先を保存しない");
        File.WriteAllText(tableSettings, "{\"Destination\":99}");
        Reject(() => TabularDropSettings.Load(tableSettings), "破損した保存設定を黙って別の既定へ置き換えない");
        int reads = 0;
        var rows = new TabularRows(11, 3, r => { reads++; return r == 0 ? ["a", "b", "c"] : ["0.1234", "0.1234", "0.1234"]; });
        var wide = new TabularData("wide.csv", rows); var model = new TabularImportSelection([wide], header: true); reads = 0;
        new TabularDropProfile(Columns: [new("wide.csv", ["a", "b", "c"], [ColumnType.Percent, ColumnType.Percent, ColumnType.Percent], [null, null, null])]).ApplyColumns([wide], model, strict: true);
        Check(reads <= 20 && model.DecimalPlaces[0].All(d => d == 2), "保存した複数列の再推定でデータ全体を列ごとに読み直さない");
    }

    static void DropDialogExcelTests(ComScope com, dynamic window, dynamic sheet)
    {
        var source = new TabularData("auto.csv", [["説明"], ["ID", "値"], ["001", "1234.56"], ["集計"]]);
        var profile = new TabularDropProfile(Header: true, Rows: new(1, 1), Columns: [new("auto.csv", ["ID", "値"], [ColumnType.TextFormat, ColumnType.Number], [null, 0])]);
        var prepared = profile.Prepare([source]); sheet.Activate(); sheet.Range["AE1000"].Select();
        ExcelTabularInserter.Import(com, com.Own(window.Application), window, prepared.Tables, prepared.Options);
        Check((string)sheet.Range["AE1001"].Value2 == "001" && (string)sheet.Range["AE1001"].PrefixCharacter == "" && (double)sheet.Range["AF1001"].Value2 == 1234.56
            && (string)sheet.Range["AF1001"].NumberFormat == "0", "省略経路の値・文字列書式・小数桁を専用Excelへ書き込む");
        sheet.Range["AE1000"].Select();
        Reject(() => ExcelTabularInserter.Import(com, com.Own(window.Application), window, prepared.Tables, prepared.Options), "省略経路でも既存データの衝突を拒否する");
        Check((string)sheet.Range["AE1001"].Value2 == "001", "省略経路の衝突拒否で既存データを保持する");
        sheet.Range["AH1000"].Select();
        ExcelTabularInserter.Import(com, com.Own(window.Application), window, prepared.Tables, prepared.Options);
        Check((string)sheet.Range["AH1001"].Value2 == "001", "次のドロップは保存時の位置ではなく現在の選択セルへ書く");
    }
}
