namespace OfficeAssist;

internal sealed record DropPress(Point Position, nint Source);

internal sealed class DropGesture
{
    DropPress? pendingPress;
    bool wasDown;
    long? releasedAt;

    internal bool HasPendingPress => Volatile.Read(ref pendingPress) is not null;
    internal void RecordPress(Point position, nint source) => Interlocked.Exchange(ref pendingPress, new(position, source));

    internal DropPress? TakePress(bool down, Point cursor, Func<Point, nint> findWindow)
    {
        var press = Interlocked.Exchange(ref pendingPress, null);
        if (press is null && down && !wasDown) press = new(cursor, findWindow(cursor));
        wasDown = down;
        if (down || press is not null) releasedAt = null;
        return press;
    }

    internal bool WaitForDrop(long now, bool receiving)
    {
        releasedAt ??= now;
        // DragEnterが届いた場合は、ボタン解放より後に届くOLEのDropを待つ。
        return now - releasedAt.Value < (receiving ? 1000 : 45);
    }
}
