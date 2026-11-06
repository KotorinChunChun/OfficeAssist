namespace OfficeAssist;

// 専用Excelの初回起動をドロップのたびに繰り返さない。常駐アプリ終了時に回収する。
internal static class SharedExcelFormatPreview
{
    static ExcelFormatPreview? shared;
    static SharedExcelFormatPreview() { Application.ApplicationExit += (_, _) => Shutdown(); }
    internal static int OwnedPid => shared?.OwnedPid ?? 0;
    internal static void Shutdown() { var renderer = shared; shared = null; renderer?.Dispose(); }
    internal static IExcelFormatPreview Acquire()
    {
        if (shared is null)
        {
            shared = new();
        }
        return new Lease(shared);
    }
    sealed class Lease(ExcelFormatPreview renderer) : IExcelFormatPreview
    {
        public Task<IReadOnlyList<ExcelPreviewCell>> RenderAsync(IReadOnlyList<ExcelPreviewInput> cells, ExcelPreviewEnvironment environment, CancellationToken token) => renderer.RenderAsync(cells, environment, token);
        public void Dispose() { }
    }
}
