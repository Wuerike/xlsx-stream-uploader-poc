using SpreadCheetah;
using XlsxStreamUploaderLoadTest.Generators;

namespace XlsxStreamUploaderLoadTest.Writers;

/// <summary>
/// SpreadCheetah: async API, requires List of Cell per row.
/// </summary>
public class SpreadCheetahWriter : IXlsxWriter
{
    public async Task WriteAsync(Stream outputStream, int rowCount, int columnCount)
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);

        await spreadsheet.StartWorksheetAsync("Data");

        await foreach (var line in CsvLineGenerator.EnumerateLinesAsync(rowCount, columnCount))
        {
            var columns = CsvLineGenerator.ParseLine(line);

            var row = new List<Cell>(columns.Length);
            foreach (var col in columns)
            {
                row.Add(new Cell(col));
            }

            await spreadsheet.AddRowAsync(row);
        }

        await spreadsheet.FinishAsync();
    }
}
