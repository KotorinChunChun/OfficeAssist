namespace OfficeAssist;

internal static class ImageAssignment
{
    internal static List<(int Area, int File)> Create(int areaCount, int fileCount)
    {
        if (areaCount < 1 || fileCount < 1) throw new ArgumentOutOfRangeException();
        return Enumerable.Range(0, fileCount).Select(i => (Math.Min(i, areaCount - 1), i)).ToList();
    }
}
