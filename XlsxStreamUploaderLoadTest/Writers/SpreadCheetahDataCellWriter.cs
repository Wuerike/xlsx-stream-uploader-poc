using SpreadCheetah;
using XlsxStreamUploaderLoadTest.Generators;

namespace XlsxStreamUploaderLoadTest.Writers;

/// <summary>
/// SpreadCheetah using DataCell (struct) instead of Cell (class).
/// DataCell is a lightweight struct for simple string/number values.
/// </summary>
public class SpreadCheetahDataCellWriter : IXlsxWriter
{
    public async Task WriteAsync(Stream outputStream, int rowCount, int columnCount)
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);

        await spreadsheet.StartWorksheetAsync("Data");

        // Reusable array with DataCell (struct - no heap allocation per cell)
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
