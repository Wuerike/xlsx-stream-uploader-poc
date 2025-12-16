using SpreadCheetah;
using XlsxStreamUploaderLoadTest.Generators;

namespace XlsxStreamUploaderLoadTest.Writers;

/// <summary>
/// SpreadCheetah with ReadOnlyMemory of Cell to avoid List allocation.
/// Uses a reusable array.
/// </summary>
public class SpreadCheetahReusableArrayWriter : IXlsxWriter
{
    public async Task WriteAsync(Stream outputStream, int rowCount, int columnCount)
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);

        await spreadsheet.StartWorksheetAsync("Data");

        // Reusable array - allocated once with exact size
        var cellBuffer = new Cell[columnCount];

        await foreach (var line in CsvLineGenerator.EnumerateLinesAsync(rowCount, columnCount))
        {
            var columns = CsvLineGenerator.ParseLine(line);

            for (int i = 0; i < columns.Length; i++)
            {
                cellBuffer[i] = new Cell(columns[i]);
            }

            await spreadsheet.AddRowAsync(cellBuffer.AsMemory(0, columns.Length));
        }

        await spreadsheet.FinishAsync();
    }
}
