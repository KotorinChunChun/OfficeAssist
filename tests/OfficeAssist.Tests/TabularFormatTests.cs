using OfficeAssist;

internal static partial class Program
{
    static void TabularFormatTests()
    {
        Check(ColumnConversion.Format(ColumnType.Number, 0) == "0" && ColumnConversion.Format(ColumnType.Number, 2) == "0.00", "小数0桁は小数点なし、2桁は固定2桁");
        Check(ColumnConversion.Format(ColumnType.Date) == "yyyy/mm/dd" && ColumnConversion.Format(ColumnType.DateTime) == "yyyy/mm/dd hh:mm:ss", "日付と日時の書式を分離");
        Check((double)ColumnConversion.Convert("12.5%", ColumnType.Percent, false)! == 0.125 && (double)ColumnConversion.Convert("¥1,234", ColumnType.Yen, false)! == 1234, "パーセントと円記号の値を変換");
        Check(Math.Abs((double)ColumnConversion.Convert("25:30:00", ColumnType.Duration, false)! - 25.5 / 24) < 1e-12, "24時間を超える経過時間を保持");
        Check(ColumnConversion.Infer(["2026/11/05 08:09:10"]) == ColumnType.DateTime && ColumnConversion.Infer(["08:09:10"]) == ColumnType.Time
            && ColumnConversion.Infer(["25:30:00"]) == ColumnType.Duration, "日付時刻と時刻と経過時間を推定");
        Check(ColumnConversion.Infer(["12.5%", "0%"]) == ColumnType.Percent && ColumnConversion.Infer(["¥1,234"]) == ColumnType.Yen, "パーセント列と円の列を推定");
        var table = new TabularData("decimals.csv", [["値"], ["12.34"]]);
        var model = new TabularImportSelection([table, table]);
        Check(model.DecimalPlaces[0][0] == 2, "値から小数桁を推定");
        model.SetApplyToAll(true, 0); model.SetDecimals(0, 0, 0);
        Check(model.DecimalPlaces[1][0] == 0, "0桁設定を全ファイルへ適用");
        model.SetHeader(false);
        Check(model.DecimalPlaces[0][0] == 0 && model.DecimalPlaces[1][0] == 0, "見出し変更でも小数桁を維持");
        using var form = new TabularImportForm([table]);
        var digits = (NumericUpDown)form.Controls.Find("ColumnDecimals0", true).Single(); digits.Value = 3;
        var files = (ComboBox)form.Controls.Find("ColumnType0", true).Single(); files.SelectedIndex = (int)ColumnType.Date;
        Check(!digits.Enabled, "小数桁は数値系の形式だけで設定する");
    }

    static void TabularFormatExcelTests(ComScope com, dynamic window, dynamic sheet)
    {
        ColumnType[] types = [ColumnType.Number, ColumnType.Number, ColumnType.GroupedNumber, ColumnType.Percent, ColumnType.Yen,
            ColumnType.Date, ColumnType.DateWeekday, ColumnType.DateTime, ColumnType.Time, ColumnType.Duration, ColumnType.TextFormat, ColumnType.TextQuote];
        int[] decimals = [0, 2, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0];
        string[] row = ["1234.56", "1234.56", "1234.56", "12.5%", "1234.56", "2026/11/05 08:09:10", "2026/11/05", "2026/11/05 08:09:10", "08:09:10", "25:30:00", "001", "001"];
        var data = new TabularData("formats.csv", [types.Select(t => t.ToString()).ToArray(), row]);
        var options = new TabularImportOptions(ImportDestination.Selection, [types], true) { DecimalPlaces = [decimals] };
        sheet.Activate(); dynamic start = com.Own(sheet.Range["A100"]); start.Select();
        ExcelTabularInserter.InsertIntoWindow(com, window, new[] { data }, options);
        dynamic target = com.Own(sheet.Range["A100:L101"]), columns = com.Own(target.Columns); columns.AutoFit();
        dynamic cell = com.Own(sheet.Range["A101"]);
        Check((string)cell.Text == "1235" && Math.Abs((double)cell.Value2 - 1234.56) < 1e-9, "実Excelで0桁は小数点なし、元の値は不変");
        cell = com.Own(sheet.Range["B101"]); Check((string)cell.Text == "1234.56", "実Excelで固定2桁表示");
        cell = com.Own(sheet.Range["C101"]); Check((string)cell.Text == "1,234.56", "実Excelで桁区切り表示");
        cell = com.Own(sheet.Range["D101"]); Check((string)cell.Text == "12.5%" && (double)cell.Value2 == 0.125, "実Excelでパーセント表示");
        cell = com.Own(sheet.Range["E101"]); Check(((string)cell.Text).EndsWith("1,235"), "実Excelで円表示");
        cell = com.Own(sheet.Range["F101"]); Check((string)cell.Text == "2026/11/05", "実Excelで日付のみ表示");
        cell = com.Own(sheet.Range["G101"]); Check((string)cell.Text == "2026/11/05 (木)", "実Excelで日本語の曜日表示");
        cell = com.Own(sheet.Range["H101"]); Check((string)cell.Text == "2026/11/05 08:09:10", "実Excelで日付と時刻表示");
        cell = com.Own(sheet.Range["I101"]); Check((string)cell.Text == "08:09:10", "実Excelで時刻のみ表示");
        cell = com.Own(sheet.Range["J101"]); Check((string)cell.Text == "25:30:00", "実Excelで24時間を超える経過時間表示");
        cell = com.Own(sheet.Range["K101"]); Check((string)cell.PrefixCharacter == "" && (string)cell.NumberFormat == "@", "追加書式後も文字列の引用符なし方式を維持");
    }
}
