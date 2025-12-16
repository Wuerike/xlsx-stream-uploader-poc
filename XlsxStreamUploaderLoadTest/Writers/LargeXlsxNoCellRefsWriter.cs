using LargeXlsx;
using XlsxStreamUploaderLoadTest.Generators;

namespace XlsxStreamUploaderLoadTest.Writers;

/// <summary>
/// LargeXlsx with cell references disabled.
/// This skips generating A1, B1, etc. references which can reduce string allocations.
/// </summary>
public class LargeXlsxNoCellRefsWriter : IXlsxWriter
{
    public Task WriteAsync(Stream outputStream, int rowCount, int columnCount)
    {
        // requireCellReferences: false - skips generating cell references like "A1", "B1"
        using var xlsxWriter = new XlsxWriter(
            outputStream, 
            XlsxCompressionLevel.Fastest, 
            requireCellReferences: false);

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
