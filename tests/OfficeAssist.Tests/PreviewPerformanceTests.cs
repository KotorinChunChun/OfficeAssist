using OfficeAssist;

internal static partial class Program
{
    static void PreviewPerformanceTests()
    {
        using var exact = new ExcelFormatPreview(cacheResults: false, batchStyles: false, literalText: false);
        using var fast = new ExcelFormatPreview();
        string[] values = ["", "001", "=1+2", "'abc", "''abc", "前半\r\n後半", "前半\n後半", "\tタブ", new string('文', 255), new string('文', 1024), new string('文', 32760)];
        var inputs = values.SelectMany((v, i) => new[] { new ExcelPreviewInput(i, 0, v, ColumnType.TextFormat, 0, new(WidthCharacters: 2)),
            new ExcelPreviewInput(i, 1, v, ColumnType.TextQuote, 0, new(WidthCharacters: 30)), new ExcelPreviewInput(i, 2, v, ColumnType.General, 0, new(WidthCharacters: 30)) }).ToArray();
        var expected = exact.RenderAsync(inputs, new(), default).GetAwaiter().GetResult();
        var actual = fast.RenderAsync(inputs, new(), default).GetAwaiter().GetResult();
        for (int i = 0; i < expected.Count; i++)
        {
            if (actual[i] != expected[i]) Console.WriteLine($"差分 {i}: type={inputs[i].Type}, rawLength={inputs[i].Raw.Length}, displayLength={expected[i].Display.Length}/{actual[i].Display.Length}, valueLength={(expected[i].StoredValue as string)?.Length}/{(actual[i].StoredValue as string)?.Length}, code={expected[i].NumberFormat}/{actual[i].NumberFormat}, prefix={expected[i].Prefix}/{actual[i].Prefix}, warning={expected[i].Warning}/{actual[i].Warning}");
            Check(actual[i] == expected[i], $"高速表示とExcelの一致: {inputs[i].Type}・長さ{inputs[i].Raw.Length}");
        }
        var longInputs = BenchmarkInputs(10, true, false, true);
        expected = exact.RenderAsync(longInputs, new(), default).GetAwaiter().GetResult();
        actual = fast.RenderAsync(longInputs, new(), default).GetAwaiter().GetResult();
        Check(actual.SequenceEqual(expected), "最大長付近の固有文字列80セルもExcelと一致");
        var styled = new[] { new ExcelPreviewStyle(WidthPoints: 0, RowHeight: 0), new ExcelPreviewStyle(WidthCharacters: 2, Wrap: true),
            new ExcelPreviewStyle(WidthCharacters: 1, Shrink: true, FontSize: 24, Bold: true) };
        var styledInputs = styled.SelectMany((s, i) => new[] { new ExcelPreviewInput(i, 0, "001", ColumnType.TextFormat, 0, s), new ExcelPreviewInput(i, 1, "前半\n後半", ColumnType.TextFormat, 0, s) }).ToArray();
        Check(fast.RenderAsync(styledInputs, new(), default).GetAwaiter().GetResult().SequenceEqual(exact.RenderAsync(styledInputs, new(), default).GetAwaiter().GetResult()), "非表示幅・行高・折り返し・縮小でも文字列の表示がExcelと一致");
        var unique = BenchmarkInputs(10, true, false, false);
        expected = exact.RenderAsync(unique, new(), default).GetAwaiter().GetResult();
        actual = fast.RenderAsync(unique, new(), default).GetAwaiter().GetResult();
        Check(actual.SequenceEqual(expected), "数値・日付・曜日・日時・時刻・率・経過時間の範囲設定が個別設定と一致");
        fast.RenderAsync(unique, new(), default).GetAwaiter().GetResult();
        Check(fast.LastMetrics.ExcelCells == 0 && fast.LastMetrics.Cached > 0, "同じ値と属性はExcelへ再転送しない");
        var edited = unique.Select(i => i.Column == 1 ? i with { Decimals = 0 } : i).ToArray();
        actual = fast.RenderAsync(edited, new(), default).GetAwaiter().GetResult();
        expected = exact.RenderAsync(edited, new(), default).GetAwaiter().GetResult();
        Check(actual.SequenceEqual(expected) && fast.LastMetrics.ExcelCells == 10, "1列変更時はその列だけExcelで再計算する");
        var moved = unique.Select(i => i with { Row = i.Row + 100 }).ToArray();
        actual = fast.RenderAsync(moved, new(), default).GetAwaiter().GetResult();
        Check(fast.LastMetrics.ExcelCells == 0 && actual[0].Row == 100, "同じ表示の別行は結果を再利用し位置だけ置換する");
        using var textOnly = new ExcelFormatPreview();
        var literal = textOnly.RenderAsync([new(0, 0, "001", ColumnType.TextFormat, 0)], new(), default).GetAwaiter().GetResult();
        Check(literal[0].Display == "001" && textOnly.OwnedPid == 0, "書式設定方式の文字列だけなら専用Excelを起動しない");
        string[][] samples = [["001", "002"], ["0.1234", "12.5%"], ["001", "¥2.50"], ["2026/11/06", "2026/11/07 08:09:10"], ["25:30:00", "08:09:10"], ["1", "文字", "¥2"]];
        foreach (var sample in samples)
        {
            var table = new TabularData("inference.csv", sample.Select(s => new[] { s }).ToArray()); var inferred = ColumnInference.Scan(table, false);
            var type = ColumnConversion.Infer(sample); int decimals = ColumnConversion.InferDecimals(sample, type);
            Check(inferred.Types[0] == type && inferred.Decimals[0] == decimals, "高速列推定は従来の型・小数桁と一致: " + string.Join('/', sample));
        }
        var bad = new TabularData("invalid.csv", Enumerable.Range(0, 1000).Select(i => new[] { i % 2 == 0 ? "003" : "2026/11/06" }).ToArray());
        var validation = ImportPreviewValidation.Check([bad], new(ImportDestination.NewBook, [[ColumnType.Date]], false), default);
        Check(validation.ErrorCount == 500 && validation.Examples.Count == 200, "例外を使わない全件検査も件数と先頭200件を維持する");
        int reads = 0;
        var guaranteed = new TabularData("literal.csv", new TabularRows(10000, 3, _ => { reads++; return ["001", "不正な日付", "=1+1"]; }));
        var accepted = ImportPreviewValidation.Check([guaranteed], new(ImportDestination.NewBook, [[ColumnType.TextFormat, ColumnType.TextQuote, ColumnType.General]], false), default);
        Check(accepted.Cells == 30000 && accepted.ErrorCount == 0 && reads == 0, "任意文字列を受理する形式は全件成功を確定してデータを再読込みしない");
        using var firstLease = SharedExcelFormatPreview.Acquire();
        firstLease.RenderAsync([new(0, 0, "42", ColumnType.Number, 0)], new(), default).GetAwaiter().GetResult();
        int sharedPid = SharedExcelFormatPreview.OwnedPid; firstLease.Dispose();
        using var secondLease = SharedExcelFormatPreview.Acquire();
        secondLease.RenderAsync([new(0, 0, "43", ColumnType.Number, 0)], new(), default).GetAwaiter().GetResult();
        Check(sharedPid > 0 && SharedExcelFormatPreview.OwnedPid == sharedPid, "ダイアログを閉じても専用Excelを再利用し初回起動を繰り返さない");
        SharedExcelFormatPreview.Shutdown();
        bool ended; try { using var process = System.Diagnostics.Process.GetProcessById(sharedPid); ended = process.HasExited; } catch (ArgumentException) { ended = true; }
        Check(ended && SharedExcelFormatPreview.OwnedPid == 0, "常駐終了に相当する処理で共有の専用Excelだけを回収する");
        for (int chunk = 0; chunk < 9; chunk++)
        {
            var valuesForCache = Enumerable.Range(chunk * 512, 512).Select(i => new ExcelPreviewInput(i, 0, (100000 + i).ToString(), ColumnType.Number, 0)).ToArray();
            fast.RenderAsync(valuesForCache, new(), default).GetAwaiter().GetResult();
        }
        fast.RenderAsync(Enumerable.Range(0, 512).Select(i => new ExcelPreviewInput(i, 0, (100000 + i).ToString(), ColumnType.Number, 0)).ToArray(), new(), default).GetAwaiter().GetResult();
        Check(fast.LastMetrics.Cached == 0 && fast.LastMetrics.ExcelCells == 512, "4096件を超えた表示キャッシュの古い値を回収する");
    }
}
