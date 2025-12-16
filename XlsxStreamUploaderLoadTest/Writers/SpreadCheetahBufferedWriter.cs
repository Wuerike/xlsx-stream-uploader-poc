using SpreadCheetah;
using XlsxStreamUploaderLoadTest.Generators;

namespace XlsxStreamUploaderLoadTest.Writers;

/// <summary>
/// SpreadCheetah with larger buffer size to reduce flush frequency.
/// </summary>
public class SpreadCheetahBufferedWriter : IXlsxWriter
{
    public async Task WriteAsync(Stream outputStream, int rowCount, int columnCount)
    {
        var options = new SpreadCheetahOptions
        {
            BufferSize = 4 * 1024 * 1024 // 4MB buffer
        };

        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, options);

        await spreadsheet.StartWorksheetAsync("Data");

        var cellBuffer = new DataCell[columnCount];

        await foreach (var line in CsvLineGenerator.EnumerateLinesAsync(rowCount, columnCount))
        {
            var columns = CsvLineGenerator.ParseLine(line);

            for (int i = 0; i < columns.Length; i++)
            {
                cellBuffer[i] = new DataCell(columns[i]);
            }

            await spreadsheet.AddRowAsync(cellBuffer.AsMemory(0, columns.Length));
        }

        await spreadsheet.FinishAsync();
    }
}
