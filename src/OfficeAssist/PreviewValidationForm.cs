namespace OfficeAssist;

internal sealed class PreviewValidationForm : BrandedForm
{
    internal PreviewValidationForm(ImportPreviewValidation? validation, IEnumerable<ExcelPreviewCell>? displayErrors, string? failure)
    {
        Text = WindowTitle.Create("取り込み前の検証結果"); ClientSize = new(1040, 540); MinimumSize = new(700, 350); MinimizeBox = false; TopMost = true; ShowInTaskbar = false;
        var grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            AllowUserToOrderColumns = false, AutoGenerateColumns = false, RowHeadersVisible = false };
        foreach (var item in new[] { ("元の位置", 320), ("列", 60), ("元の値", 230), ("理由", 500) })
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = item.Item1, Width = item.Item2, SortMode = DataGridViewColumnSortMode.NotSortable });
        if (validation is not null) foreach (var error in validation.Examples) grid.Rows.Add(error.Origin, error.Column + 1, error.Raw, error.Reason);
        if (displayErrors is not null) foreach (var error in displayErrors.Take(200)) grid.Rows.Add($"プレビュー行{error.Row + 1}", error.Column + 1, error.Raw, error.Error ?? "");
        if (failure is not null && grid.Rows.Count == 0) grid.Rows.Add("", "", "", failure);
        var summary = new Label { Dock = DockStyle.Top, Height = 32, Text = validation is null ? "全件の変換検査を実行中" : $"検査 {validation.Cells}セル / 変換エラー {validation.ErrorCount}件（先頭200件まで表示）" };
        var close = new Button { Text = "閉じる", Dock = DockStyle.Bottom, DialogResult = DialogResult.OK, Height = 34 }; Controls.Add(grid); Controls.Add(summary); Controls.Add(close); AcceptButton = close;
    }
}
