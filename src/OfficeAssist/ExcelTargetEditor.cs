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
    internal event Action? Changed;
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    internal Action<nint>? TargetWindowChanged { get; set; }
    internal ExcelTargetEditor(ComScope com, dynamic window)
    {
        this.com = com; app = com.Own(window.Application); AutoSize = true; WrapContents = true;
        dynamic application = app, allBooks = com.Own(application.Workbooks);
        for (int i = 1; i <= (int)allBooks.Count; i++)
        {
            dynamic book = com.Own(allBooks.Item(i));
            if (!VisibleWindows(com, (object)book).Any()) continue;
            bookIds.Add((string)book.FullName);
            books.Items.Add((string)book.Name);
        }
        Controls.AddRange([Label("ブック"), books, plannedBook, Label("シート"), sheets, plannedSheet, Label("セル"), address, plannedCell]);
        books.SelectedIndexChanged += (_, _) => { try { LoadSheets(); } catch (Exception ex) { sheets.Items.Clear(); MessageBox.Show(FindForm(), ex.Message, WindowTitle.Create("出力先"), MessageBoxButtons.OK, MessageBoxIcon.Warning); } };
        if (bookIds.Count > 0)
        {
            object? active = window.ActiveSheet;
            if (active is not null)
            {
                dynamic sourceSheet = com.Own(active), sourceBook = com.Own(sourceSheet.Parent);
                books.SelectedIndex = bookIds.IndexOf((string)sourceBook.FullName);
                if (books.SelectedIndex >= 0 && sheets.Items.Contains((string)sourceSheet.Name))
                {
                    sheets.SelectedItem = (string)sourceSheet.Name;
                    object? selected = window.RangeSelection;
                    if (selected is not null) address.Text = ExcelAddress.Display(com.Own(selected));
                }
            }
            if (books.SelectedIndex < 0) books.SelectedIndex = 0;
        }
        books.SelectedIndexChanged += (_, _) => Changed?.Invoke(); sheets.SelectedIndexChanged += (_, _) => Changed?.Invoke(); address.TextChanged += (_, _) => Changed?.Invoke();
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
        if (books.SelectedIndex < 0 || TargetWindowChanged is null) return;
        dynamic application = app;
        object book = ExcelImportEngine.ResolveBook(com, application, null, bookIds[books.SelectedIndex]);
        foreach (dynamic window in VisibleWindows(com, book)) TargetWindowChanged((nint)(int)window.Hwnd);
    }
    internal static IEnumerable<object> VisibleWindows(ComScope com, object workbook)
    {
        dynamic book = workbook;
        if ((bool)book.IsAddin) yield break;
        dynamic windows = com.Own(book.Windows);
        for (int i = 1; i <= (int)windows.Count; i++)
        {
            dynamic window = com.Own(windows.Item(i));
            if ((bool)window.Visible) yield return window;
        }
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
