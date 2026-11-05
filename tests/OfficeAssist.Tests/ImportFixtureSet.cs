using System.Text;
using System.Text.Json;
using OfficeAssist;

internal sealed record ImportFixture(string RelativePath, TableEncoding Encoding, bool Auto, string Purpose);

internal static class ImportFixtureSet
{
    internal static ImportFixture[] Create(string root)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var fixtures = new List<ImportFixture>();
        void Write(string relative, string text, Encoding encoding, TableEncoding choice, bool auto, string purpose)
        {
            string path = Path.Combine(root, relative); Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, text, encoding); fixtures.Add(new(relative, choice, auto, purpose));
        }
        string basic = "ID,氏名,金額\r\n001,日本語,12.34\r\n002,\"前半\r\n後半\",0\r\n";
        Write("valid/01-utf8-bom.csv", basic, new UTF8Encoding(true), TableEncoding.Utf8, true, "UTF-8 BOM・先頭ゼロ・引用符内改行");
        Write("valid/02-utf8-no-bom.tsv", basic.Replace(',', '\t'), new UTF8Encoding(false), TableEncoding.Utf8, true, "UTF-8 BOMなし・TSV");
        Write("valid/03-utf16le-bom.csv", basic, new UnicodeEncoding(false, true), TableEncoding.Utf16Le, true, "UTF-16 LE BOM");
        Write("valid/04-utf16be-bom.TSV", basic.Replace(',', '\t'), new UnicodeEncoding(true, true), TableEncoding.Utf16Be, true, "UTF-16 BE BOM・大文字拡張子");
        Write("valid/05-cp932.csv", basic, Encoding.GetEncoding(932), TableEncoding.Cp932, true, "CP932自動判定と手動指定");
        Write("manual/06-utf16le-no-bom.csv", basic, new UnicodeEncoding(false, false), TableEncoding.Utf16Le, false, "BOMなしUTF-16 LEは手動で読む");
        Write("manual/07-utf16be-no-bom.tsv", basic.Replace(',', '\t'), new UnicodeEncoding(true, false), TableEncoding.Utf16Be, false, "BOMなしUTF-16 BEは手動で読む");
        Write("valid/08-emoji.csv", "ID,氏名\n001,🕊️\n", new UTF8Encoding(false), TableEncoding.Utf8, true, "CP932にない文字をUTF-8として保持");
        Write("edge/09-trim.csv", "説明1\r\n説明2\r\nID,値\r\n001,3\r\n,,\r\n002,4\r\n集計,7\r\n", new UTF8Encoding(true), TableEncoding.Utf8, true, "先頭2・末尾1・空欄除外");
        Write("edge/10-trim-swapped.tsv", "説明1\n説明2\n値\tID\t備考\n5\t003\t追加\n集計\t5\n", Encoding.GetEncoding(932), TableEncoding.Cp932, true, "混在文字コード・列順違い・追加列");
        Write("reject/11-duplicate-headers.csv", "ID, ID \n001,002\n", new UTF8Encoding(false), TableEncoding.Utf8, true, "正規化後の重複見出し");
        Write("edge/12-unknown-append.csv", "未知列\n値\n", new UTF8Encoding(false), TableEncoding.Utf8, true, "追記の未知列拒否");
        Write("edge/13-header-only.csv", "ID,値,追加\n", new UTF8Encoding(true), TableEncoding.Utf8, true, "データなしの見出し");
        Write("edge/14-empty-space.csv", "ID,値\n001,\"\"\n,\n002, \n", new UTF8Encoding(false), TableEncoding.Utf8, true, "空欄行と空白値の区別");
        Write("reject/15-malformed.csv", "ID,値\n001,\"閉じない", new UTF8Encoding(false), TableEncoding.Utf8, true, "閉じ引用符なし");
        Write("edge/16-formula-like.csv", "ID,文字\n0000000000000001,=1+2\n0007,'quoted\n", new UTF8Encoding(false), TableEncoding.Utf8, true, "長い番号と数式風文字列");
        Write("folders/a/common.csv", "ID,値\n001,1\n", new UTF8Encoding(false), TableEncoding.Utf8, true, "同名別パスの保持");
        Write("folders/b/common.csv", "ID,値\n002,2\n", new UTF8Encoding(false), TableEncoding.Utf8, true, "同名別パスの順次追記");
        Directory.CreateDirectory(Path.Combine(root, "reject"));
        File.WriteAllBytes(Path.Combine(root, "reject/17-invalid-utf8.csv"), [0xF0, 0x28, 0x8C, 0x28]);
        fixtures.Add(new("reject/17-invalid-utf8.csv", TableEncoding.Utf8, false, "不正UTF-8を置換せず拒否"));
        File.WriteAllBytes(Path.Combine(root, "reject/18-empty.csv"), []);
        fixtures.Add(new("reject/18-empty.csv", TableEncoding.Auto, false, "空ファイル拒否"));
        File.WriteAllText(Path.Combine(root, "folders/ignored.txt"), "探索対象外");
        var metadata = fixtures.Select(f => new { f.RelativePath, encoding = TabularEncoding.Labels[(int)f.Encoding], auto = f.Auto, f.Purpose,
            bytes = new FileInfo(Path.Combine(root, f.RelativePath)).Length,
            sha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(Path.Combine(root, f.RelativePath)))).ToLowerInvariant() });
        File.WriteAllText(Path.Combine(root, "manifest.json"), JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping }));
        return fixtures.ToArray();
    }
}
