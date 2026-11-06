using System.Globalization;

namespace OfficeAssist;

internal sealed partial class TabularImportForm
{
    readonly Button validationDetails = new() { Name = "PreviewValidation", Text = "検証結果…", AutoSize = true };
    FormattedPreviewController? formattedPreview;
    Func<IReadOnlyList<TabularData>, TabularImportOptions, int, IReadOnlyList<(int Row, int Column)>, ExcelPreviewSnapshot>? capturePreview;
    ImportPreviewValidation? validationResult;
    CancellationTokenSource? validationCancellation;
    Task validationWork = Task.CompletedTask;
    string basePreviewStatus = "";
    string? validationFailure;
    bool validating;
    int validationGeneration;
    bool FormattingAllowsImport => formattedPreview is null || (!validating && validationResult is not null && validationResult.ErrorCount == 0
        && validationFailure is null && formattedPreview.Failure is null && !formattedPreview.Busy);
    internal bool IsFormattingReady => valid && FormattingAllowsImport;

    void InitializeFormatting(IExcelFormatPreview? renderer,
        Func<IReadOnlyList<TabularData>, TabularImportOptions, int, IReadOnlyList<(int Row, int Column)>, ExcelPreviewSnapshot>? capture)
    {
        capturePreview = capture;
        if (renderer is null) { validationDetails.Visible = false; return; }
        formattedPreview = new(this, renderer, () => { preview.Invalidate(); UpdateFormattingState(); });
        if (targets is not null) targets.Changed += RefreshFormatting;
        preview.CellFormatting += (_, e) =>
        {
            if (e.CellStyle is null || e.RowIndex < 0 || e.RowIndex >= previewRows.Length || e.ColumnIndex < 0) return;
            var cell = formattedPreview.Get(previewRows[e.RowIndex], e.ColumnIndex);
            e.CellStyle.BackColor = cell?.Error is not null ? Color.MistyRose : cell?.Warning is not null ? Color.LemonChiffon : SystemColors.Window;
            e.CellStyle.ForeColor = cell?.Error is not null ? Color.DarkRed : SystemColors.WindowText;
        };
        validationDetails.Click += (_, _) => { using var result = new PreviewValidationForm(validationResult, formattedPreview.Errors, validationFailure ?? formattedPreview.Failure); result.ShowDialog(this); };
    }

    void RefreshFormatting()
    {
        if (tables.Count == 0 || updating || working || !valid) return;
        foreach (var field in settings.Controls.OfType<TextBox>())
            if (field.Name.StartsWith("ColumnFormat", StringComparison.Ordinal) && int.TryParse(field.Name[12..], out int column))
                field.Text = ExcelValueFormat.Code(selection.Types[SelectedFile][column], selection.DecimalPlaces[SelectedFile][column]);
        if (formattedPreview is null) return;
        CancelPreviewWork();
        validationResult = null; validationFailure = null;
        TabularImportOptions options;
        try { options = selection.CreateOptions((ImportDestination)destination.SelectedIndex) with { Merged = merge.Checked }; if (targets is not null) options = targets.Apply(options); }
        catch (Exception ex) { validationFailure = ex.Message; UpdateFormattingState(); return; }
        int file = SelectedFile; var data = tables;
        if (previewMode.SelectedIndex != 0)
            formattedPreview.Reset(data[file], options.Types[file], Enumerable.Range(0, data[file].ColumnCount).Select(c => options.Decimals(file, c)).ToArray(),
                positions => capturePreview?.Invoke(data, options, file, positions) ?? new(new(), positions.Select(_ => new ExcelPreviewStyle()).ToArray()));
        int revision = ++validationGeneration; var cancellation = new CancellationTokenSource(); validationCancellation = cancellation;
        validating = true; tasks++;
        validationWork = ValidatePreview(data, options, revision, cancellation);
        preview.Invalidate(); UpdateFormattingState();
    }

    async Task ValidatePreview(IReadOnlyList<TabularData> data, TabularImportOptions options, int revision, CancellationTokenSource cancellation)
    {
        try
        {
            await Task.Delay(150, cancellation.Token);
            var result = await Task.Run(() => ImportPreviewValidation.Check(data, options, cancellation.Token), cancellation.Token);
            if (revision == validationGeneration && !IsDisposed) validationResult = result;
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { if (revision == validationGeneration && !IsDisposed) validationFailure = ex.Message; }
        finally
        {
            tasks--; cancellation.Dispose();
            if (revision == validationGeneration) { validationCancellation = null; validating = false; if (!IsDisposed) UpdateFormattingState(); }
            if (pendingClose && tasks == 0 && !IsDisposed) { DialogResult = DialogResult.Cancel; Close(); }
        }
    }

    void UpdateFormattingState()
    {
        if (IsDisposed || working || !valid || formattedPreview is null) return;
        string detail = validating ? "全件の変換検査中" : validationFailure is not null ? "検証エラー: " + validationFailure
            : validationResult?.ErrorCount > 0 ? $"変換エラー {validationResult.ErrorCount}件"
            : formattedPreview.Failure is not null ? "表示確認エラー: " + formattedPreview.Failure
            : formattedPreview.Busy ? "Excelの表示を取得中" : "全件の変換検査完了";
        if (formattedPreview.WarningCount > 0) detail += $" / 列幅不足 {formattedPreview.WarningCount}セル";
        status.Text = basePreviewStatus + " / " + detail;
        validationDetails.Enabled = !validating && (validationResult is not null || validationFailure is not null || formattedPreview.Failure is not null);
        import.Enabled = tasks == 0 && FormattingAllowsImport;
    }

    string PreviewToolTip(int row, int column)
    {
        string origin = tables[SelectedFile].Origin?.Invoke(row) ?? tables[SelectedFile].Name + ": 行" + (row + 1);
        if (column < 0 || formattedPreview is null) return origin;
        var cell = formattedPreview.Get(row, column); if (cell is null) return origin + "\nExcelの表示を取得中";
        static string Limit(string value) => value.Length <= 1000 ? value : value[..1000] + "…";
        string value = Convert.ToString(cell.StoredValue, CultureInfo.CurrentCulture) ?? "（空欄）";
        return origin + $"\n元の値: {Limit(cell.Raw)}\n書式コード: {cell.FormatCode}\n実際のセル書式: {cell.NumberFormat}\n格納値: {Limit(value)}"
            + (cell.Prefix.Length == 0 ? "" : "\n接頭辞: " + cell.Prefix)
            + (cell.Error is null ? "" : "\nエラー: " + cell.Error) + (cell.Warning is null ? "" : "\n注意: " + cell.Warning);
    }

    void CancelPreviewWork()
    {
        validationGeneration++; validationCancellation?.Cancel(); validationCancellation = null; validating = false;
        formattedPreview?.Suspend();
    }
}
