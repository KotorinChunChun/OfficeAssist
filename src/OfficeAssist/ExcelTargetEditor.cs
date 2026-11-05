namespace OfficeAssist;

internal sealed class ExcelTargetEditor : FlowLayoutPanel
{
    readonly ImportComboBox books = new() { Name = "TargetBook", Width = 190 };
    readonly ImportComboBox sheets = new() { Name = "TargetSheet", Width = 140 };
    readonly TextBox address = new() { Name = "TargetAddress", Width = 180, Text = "A1" };
    readonly Label plannedBook = new() { Text = "新規作成", Width = 190, Padding = new(0, 5, 0, 0) };
    readonly Label plannedSheet = new() { Width = 160, Padding = new(0, 5, 0, 0) };
    readonly Label plannedCell = new() { Width = 140, Padding = new(0, 5, 0, 0) };
    readonly List<string> bookIds = new();
    readonly ComScope com;
    readonly object app;
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    internal Action<nint>? TargetWindowChanged { get; set; }
    internal ExcelTargetEditor(ComScope com, dynamic window)
    {
        this.com = com; app = com.Own(window.Application); AutoSize = true; WrapContents = true;
        dynamic application = app, allBooks = com.Own(application.Workbooks), sourceSheet = com.Own(window.ActiveSheet), sourceBook = com.Own(sourceSheet.Parent), selected = com.Own(window.RangeSelection);
        for (int i = 1; i <= (int)allBooks.Count; i++) { dynamic book = com.Own(allBooks.Item(i)); bookIds.Add((string)book.FullName); books.Items.Add((string)book.Name); }
        Controls.AddRange([Label("ブック"), books, plannedBook, Label("シート"), sheets, plannedSheet, Label("セル"), address, plannedCell]);
        books.SelectedIndexChanged += (_, _) => { try { LoadSheets(); } catch (Exception ex) { sheets.Items.Clear(); MessageBox.Show(FindForm(), ex.Message, WindowTitle.Create("出力先"), MessageBoxButtons.OK, MessageBoxIcon.Warning); } };
        books.SelectedIndex = bookIds.IndexOf((string)sourceBook.FullName); sheets.SelectedItem = (string)sourceSheet.Name; address.Text = ExcelAddress.Display(selected);
    }
    static Label Label(string text) => new() { Text = text, AutoSize = true, Padding = new(0, 5, 0, 0) };
    void LoadSheets()
    {
        sheets.Items.Clear(); if (books.SelectedIndex < 0) return;
        dynamic application = app, book = ExcelImportEngine.ResolveBook(com, application, null, bookIds[books.SelectedIndex]), all = com.Own(book.Worksheets);
        for (int i = 1; i <= (int)all.Count; i++) { dynamic sheet = com.Own(all.Item(i)); sheets.Items.Add((string)sheet.Name); }
        if (sheets.Items.Count != 0) sheets.SelectedIndex = 0;
        LockCurrentWindows();
    }
    internal void LockCurrentWindows()
    {
        if (books.SelectedIndex < 0) return;
        dynamic application = app, book = ExcelImportEngine.ResolveBook(com, application, null, bookIds[books.SelectedIndex]), windows = com.Own(book.Windows);
        for (int i = 1; i <= (int)windows.Count; i++) { dynamic window = com.Own(windows.Item(i)); TargetWindowChanged?.Invoke((nint)(int)window.Hwnd); }
    }
    internal void SetMode(ImportDestination mode)
    {
        books.Visible = books.Enabled = mode != ImportDestination.NewBook; plannedBook.Visible = mode == ImportDestination.NewBook;
        sheets.Visible = sheets.Enabled = address.Visible = address.Enabled = mode == ImportDestination.Selection;
        plannedSheet.Visible = plannedCell.Visible = mode != ImportDestination.Selection;
        plannedSheet.Text = mode == ImportDestination.AppendSameSheet ? "各ファイルと同名" : "新規作成";
        plannedCell.Text = mode == ImportDestination.AppendSameSheet ? "最終行の次" : "A1";
    }
    internal TabularImportOptions Apply(TabularImportOptions options)
    {
        if (options.Destination == ImportDestination.NewBook) return options;
        if (books.SelectedIndex < 0) throw new InvalidDataException("出力先ブックを選択してください。");
        return options with { BookId = bookIds[books.SelectedIndex], SheetName = options.Destination == ImportDestination.Selection ? sheets.Text : null,
            CellAddress = options.Destination == ImportDestination.Selection ? ExcelImportEngine.NormalizeAddress(address.Text) : null };
    }
}
