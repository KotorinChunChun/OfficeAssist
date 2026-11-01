namespace OfficeAssist;

internal static class ImageDragOrigin
{
    // プロセス名を限定すると、SSCapture等が渡す正しいFileDropも受信前に拒否してしまう。
    // Excel内の通常操作とOfficeAssist自身は候補にせず、データの可否はDragEnterで判定する。
    internal static bool IsEligible(string rootClass, uint processId, uint ownProcessId) =>
        processId != 0 && processId != ownProcessId && !string.IsNullOrEmpty(rootClass)
        && !rootClass.Equals("XLMAIN", StringComparison.OrdinalIgnoreCase);
}
