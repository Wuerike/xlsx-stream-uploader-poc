namespace XlsxStreamUploaderLoadTest.Writers;

/// <summary>
/// Interface for XLSX writers to be benchmarked.
/// </summary>
public interface IXlsxWriter
{
    /// <summary>
    /// Writes CSV data as XLSX to the provided stream.
    /// </summary>
    Task WriteAsync(Stream outputStream, int rowCount, int columnCount);
}
