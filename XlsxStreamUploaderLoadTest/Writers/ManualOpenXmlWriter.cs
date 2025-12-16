using System.IO.Compression;
using System.Text;
using System.Xml;
using XlsxStreamUploaderLoadTest.Generators;

namespace XlsxStreamUploaderLoadTest.Writers;

/// <summary>
/// Manual OpenXML: ZipArchive + XmlWriter, no external XLSX library.
/// </summary>
public class ManualOpenXmlWriter : IXlsxWriter
{
    public async Task WriteAsync(Stream outputStream, int rowCount, int columnCount)
    {
        using var zipArchive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);

        CreateContentTypesXml(zipArchive);
        CreateRelsXml(zipArchive);
        CreateWorkbookRelsXml(zipArchive);
        CreateWorkbookXml(zipArchive);
        await CreateWorksheetXmlFromCsv(zipArchive, rowCount, columnCount);
    }

    private static void CreateContentTypesXml(ZipArchive zip)
    {
        var entry = zip.CreateEntry("[Content_Types].xml", CompressionLevel.Fastest);
        using var entryStream = entry.Open();
        using var writer = new StreamWriter(entryStream, Encoding.UTF8);

        writer.Write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            </Types>
            """);
    }

    private static void CreateRelsXml(ZipArchive zip)
    {
        var entry = zip.CreateEntry("_rels/.rels", CompressionLevel.Fastest);
        using var entryStream = entry.Open();
        using var writer = new StreamWriter(entryStream, Encoding.UTF8);

        writer.Write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>
            """);
    }

    private static void CreateWorkbookRelsXml(ZipArchive zip)
    {
        var entry = zip.CreateEntry("xl/_rels/workbook.xml.rels", CompressionLevel.Fastest);
        using var entryStream = entry.Open();
        using var writer = new StreamWriter(entryStream, Encoding.UTF8);

        writer.Write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
            </Relationships>
            """);
    }

    private static void CreateWorkbookXml(ZipArchive zip)
    {
        var entry = zip.CreateEntry("xl/workbook.xml", CompressionLevel.Fastest);
        using var entryStream = entry.Open();
        using var writer = new StreamWriter(entryStream, Encoding.UTF8);

        writer.Write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets>
                    <sheet name="Data" sheetId="1" r:id="rId1"/>
                </sheets>
            </workbook>
            """);
    }

    private static async Task CreateWorksheetXmlFromCsv(ZipArchive zip, int rowCount, int columnCount)
    {
        var entry = zip.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
        await using var entryStream = entry.Open();
        await using var xmlWriter = XmlWriter.Create(entryStream, new XmlWriterSettings
        {
            Encoding = Encoding.UTF8,
            Async = true,
            CloseOutput = false
        });

        await xmlWriter.WriteStartDocumentAsync();
        await xmlWriter.WriteStartElementAsync(null, "worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        await xmlWriter.WriteStartElementAsync(null, "sheetData", null);

        uint rowIndex = 1;

        foreach (var line in CsvLineGenerator.EnumerateLines(rowCount, columnCount))
        {
            var columns = CsvLineGenerator.ParseLine(line);

            await xmlWriter.WriteStartElementAsync(null, "row", null);
            await xmlWriter.WriteAttributeStringAsync(null, "r", null, rowIndex.ToString());

            for (int colIndex = 0; colIndex < columns.Length; colIndex++)
            {
                string cellValue = columns[colIndex];
                string cellReference = GetCellReference(rowIndex, colIndex);

                await xmlWriter.WriteStartElementAsync(null, "c", null);
                await xmlWriter.WriteAttributeStringAsync(null, "r", null, cellReference);
                await xmlWriter.WriteAttributeStringAsync(null, "t", null, "inlineStr");

                await xmlWriter.WriteStartElementAsync(null, "is", null);
                await xmlWriter.WriteStartElementAsync(null, "t", null);
                await xmlWriter.WriteStringAsync(cellValue);
                await xmlWriter.WriteEndElementAsync(); // t
                await xmlWriter.WriteEndElementAsync(); // is

                await xmlWriter.WriteEndElementAsync(); // c
            }

            await xmlWriter.WriteEndElementAsync(); // row
            rowIndex++;
        }

        await xmlWriter.WriteEndElementAsync(); // sheetData
        await xmlWriter.WriteEndElementAsync(); // worksheet
        await xmlWriter.WriteEndDocumentAsync();
        await xmlWriter.FlushAsync();
    }

    private static string GetCellReference(uint row, int col)
    {
        string columnName = GetColumnName(col);
        return $"{columnName}{row}";
    }

    private static string GetColumnName(int colIndex)
    {
        int dividend = colIndex + 1;
        string columnName = string.Empty;

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }
}
