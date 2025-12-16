using MiniExcelLibs;
using XlsxStreamUploaderLoadTest.Generators;

namespace XlsxStreamUploaderLoadTest.Writers;

/// <summary>
/// MiniExcel: high-level API, allocates Dictionary per row.
/// </summary>
public class MiniExcelWriter : IXlsxWriter
{
    public Task WriteAsync(Stream outputStream, int rowCount, int columnCount)
    {
        var values = EnumerateRowsAsDictionaries(rowCount, columnCount);
        outputStream.SaveAs(values);

        return Task.CompletedTask;
    }

    private static IEnumerable<IDictionary<string, object?>> EnumerateRowsAsDictionaries(int rowCount, int columnCount)
    {
        foreach (var line in CsvLineGenerator.EnumerateLines(rowCount, columnCount))
        {
            var columns = CsvLineGenerator.ParseLine(line);

            var dict = new Dictionary<string, object?>(columns.Length);
            for (int colIndex = 0; colIndex < columns.Length; colIndex++)
            {
                dict[$"Column{colIndex + 1}"] = columns[colIndex];
            }

            yield return dict;
        }
    }
}
