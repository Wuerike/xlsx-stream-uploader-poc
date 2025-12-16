namespace XlsxStreamUploaderLoadTest.Infrastructure;

/// <summary>
/// A stream that discards all written data. Used to isolate XLSX generation from I/O.
/// </summary>
public class NullStream : Stream
{
    private long _position;

    public override bool CanRead => false;
    public override bool CanSeek => true;
    public override bool CanWrite => true;
    public override long Length => _position;
    public override long Position
    {
        get => _position;
        set => _position = value;
    }

    public override void Flush() { }
    public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin)
    {
        return origin switch
        {
            SeekOrigin.Begin => _position = offset,
            SeekOrigin.Current => _position += offset,
            SeekOrigin.End => _position = _position + offset,
            _ => throw new ArgumentOutOfRangeException(nameof(origin))
        };
    }

    public override void SetLength(long value) => _position = value;

    public override void Write(byte[] buffer, int offset, int count) => _position += count;

    public override void Write(ReadOnlySpan<byte> buffer) => _position += buffer.Length;

    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        _position += count;
        return Task.CompletedTask;
    }

    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
    {
        _position += buffer.Length;
        return ValueTask.CompletedTask;
    }
}
