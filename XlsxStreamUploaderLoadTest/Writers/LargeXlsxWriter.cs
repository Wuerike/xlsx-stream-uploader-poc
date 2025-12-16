using LargeXlsx;
using XlsxStreamUploaderLoadTest.Generators;

namespace XlsxStreamUploaderLoadTest.Writers;

/// <summary>
/// LargeXlsx: synchronous API, writes directly to stream without intermediate collections.
/// </summary>
public class LargeXlsxWriter : IXlsxWriter
{
    public Task WriteAsync(Stream outputStream, int rowCount, int columnCount)
    {
        using var xlsxWriter = new XlsxWriter(outputStream, XlsxCompressionLevel.Fastest);

        xlsxWriter.BeginWorksheet("Data");

        foreach (var line in CsvLineGenerator.EnumerateLines(rowCount, columnCount))
        {
            var columns = CsvLineGenerator.ParseLine(line);

            xlsxWriter.BeginRow();
            foreach (var col in columns)
            {
                xlsxWriter.Write(col);
            }
        }

        return Task.CompletedTask;
    }
}
