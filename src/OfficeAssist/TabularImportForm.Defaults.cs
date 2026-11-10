namespace OfficeAssist;

internal sealed partial class TabularImportForm
{
    TabularDropProfile? pendingDefaults;
    internal TabularDropProfile CreateDropProfile() => new(
        (ImportDestination)destination.SelectedIndex, headerEdited ? selection.Header : null,
        new((int)skipFirst.Value, (int)skipLast.Value, dropEmpty.Checked, merge.Checked, merge.Checked && mergeBy.SelectedIndex == 1),
        (PreviewMode)previewMode.SelectedIndex, applyToAll.Checked, Columns: selection.ExportColumns());

    void RestoreDefaultColumns()
    {
        if (pendingDefaults is null) return;
        pendingDefaults.ApplyColumns(tables, selection, strict: false); pendingDefaults = null;
        var columns = selection.ExportColumns();
        for (int f = 0; f < columns.Length; f++) for (int c = 0; c < columns[f].Types.Length; c++)
        { if (columns[f].Types[c].HasValue) Remember(f, c, true); if (columns[f].Decimals[c].HasValue) Remember(f, c, false); }
    }
}
