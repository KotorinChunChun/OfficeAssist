using OfficeAssist;

internal static partial class Program
{
    // README撮影用。常駐サービス・利用者設定・実Excelへ接続しない。
    static int ManualPreview(string[] args, string output)
    {
        int index = Array.IndexOf(args, "--manual-preview");
        string scene = index + 1 < args.Length ? args[index + 1] : "";
        using Form form = scene switch
        {
            "main" => new MainForm { AllowExit = true },
            "settings" => new InsertOptionsForm("", 0, settingsMode: true),
            "image" => new InsertOptionsForm("C5:F12", 2),
            "table" => new TabularImportForm([new TabularData("売上サンプル.csv", [
                ["商品", "数量", "単価"], ["色紙", "3", "120"], ["木の輪", "1", "350"], ["ひも", "2", "80"]])]),
            _ => throw new ArgumentException("場面はmain/settings/image/tableを指定してください。")
        };
        form.ShowInTaskbar = true;
        form.ShowDialog();
        File.WriteAllText(Path.Combine(output, "preview-" + scene + ".txt"), "利用者設定・Excelへの保存なしで終了");
        return 0;
    }
}
