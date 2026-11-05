using System.Text;

namespace OfficeAssist;

internal enum TableEncoding { Auto, Utf8, Utf16Le, Utf16Be, Cp932 }

internal static class TabularEncoding
{
    internal static readonly string[] Labels = ["自動判定", "UTF-8", "UTF-16 LE", "UTF-16 BE", "Shift_JIS（CP932）"];
    internal static TableEncoding Parse(string value) => value.ToLowerInvariant() switch
    {
        "auto" => TableEncoding.Auto, "utf8" or "utf-8" => TableEncoding.Utf8,
        "utf16le" or "utf-16le" => TableEncoding.Utf16Le, "utf16be" or "utf-16be" => TableEncoding.Utf16Be,
        "cp932" or "shift_jis" or "sjis" => TableEncoding.Cp932,
        _ => throw new InvalidDataException("文字コードはauto・utf8・utf16le・utf16be・cp932から選んでください。")
    };
    internal static Encoding Explicit(Stream stream, TableEncoding choice)
    {
        if (!Enum.IsDefined(choice) || choice == TableEncoding.Auto) throw new InvalidDataException("文字コードの指定が不正です。");
        stream.Position = 0; byte[] prefix = new byte[4]; int n = stream.Read(prefix); stream.Position = 0;
        if (n >= 4 && (prefix.AsSpan().SequenceEqual(new byte[] { 255, 254, 0, 0 }) || prefix.AsSpan().SequenceEqual(new byte[] { 0, 0, 254, 255 }))) throw new InvalidDataException("UTF-32は対象外です。");
        TableEncoding? bom = n >= 3 && prefix[0] == 239 && prefix[1] == 187 && prefix[2] == 191 ? TableEncoding.Utf8
            : n >= 2 && prefix[0] == 255 && prefix[1] == 254 ? TableEncoding.Utf16Le
            : n >= 2 && prefix[0] == 254 && prefix[1] == 255 ? TableEncoding.Utf16Be : null;
        if (bom.HasValue && bom != choice) throw new InvalidDataException($"BOMの文字コード（{Labels[(int)bom.Value]}）と指定（{Labels[(int)choice]}）が異なります。");
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return choice switch
        {
            TableEncoding.Utf8 => new UTF8Encoding(true, true),
            TableEncoding.Utf16Le => new UnicodeEncoding(false, true, true),
            TableEncoding.Utf16Be => new UnicodeEncoding(true, true, true),
            _ => Encoding.GetEncoding(932, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback)
        };
    }
}
