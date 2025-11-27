using Amazon.S3;
using Amazon.S3.Model;
using System.IO.Pipelines;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace XlsxStreamUploader.Services;

public class UploadService(IAmazonS3 s3Client) : BaseStreamingXlsxUploadService(s3Client)
{
    protected override string ProviderName => "UploadService";

    protected override async Task WriteCsvAsXlsxToPipe(string csvS3Key, PipeWriter writer)
    {
        var pipeStream = writer.AsStream();
        
        using (var zipArchive = new ZipArchive(pipeStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            // 1. Create [Content_Types].xml
            CreateContentTypesXml(zipArchive);
            
            // 2. Create _rels/.rels
            CreateRelsXml(zipArchive);
            
            // 3. Create xl/_rels/workbook.xml.rels
            CreateWorkbookRelsXml(zipArchive);
            
            // 4. Create xl/workbook.xml
            CreateWorkbookXml(zipArchive);
            
            // 5. Create xl/worksheets/sheet1.xml from CSV stream
            int totalRows = await CreateWorksheetXmlFromCsv(zipArchive, csvS3Key);
            
            Console.WriteLine($"Finished processing {totalRows} rows from CSV");
        }
        
        await writer.CompleteAsync();
        Console.WriteLine("XLSX streaming completed");
    }
    
    private static void CreateContentTypesXml(ZipArchive zip)
    {
        var entry = zip.CreateEntry("[Content_Types].xml", CompressionLevel.Fastest);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, Encoding.UTF8);
        
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
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, Encoding.UTF8);
        
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
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, Encoding.UTF8);
        
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
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, Encoding.UTF8);
        
        writer.Write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets>
                    <sheet name="Data" sheetId="1" r:id="rId1"/>
                </sheets>
            </workbook>
            """);
    }
    
    private async Task<int> CreateWorksheetXmlFromCsv(ZipArchive zip, string csvS3Key)
    {
        var entry = zip.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
        using var stream = entry.Open();
        using var xmlWriter = XmlWriter.Create(stream, new XmlWriterSettings 
        { 
            Encoding = Encoding.UTF8,
            Async = true,
            CloseOutput = false
        });
        
        // Write worksheet header
        await xmlWriter.WriteStartDocumentAsync();
        await xmlWriter.WriteStartElementAsync(null, "worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        await xmlWriter.WriteStartElementAsync(null, "sheetData", null);
        
        int totalRows = 0;
        uint rowIndex = 1;
        
        // Read CSV from S3 as stream
        var request = new GetObjectRequest
        {
            BucketName = BucketName,
            Key = csvS3Key
        };
        
        using (var response = await S3Client.GetObjectAsync(request))
        using (var responseStream = response.ResponseStream)
        using (var reader = new StreamReader(responseStream))
        {
            string? line;
            while ((line = await reader.ReadLineAsync()) != null)
            {
                // Parse CSV line
                var columns = ParseCsvLine(line, reader);
                
                // Write row
                await xmlWriter.WriteStartElementAsync(null, "row", null);
                await xmlWriter.WriteAttributeStringAsync(null, "r", null, rowIndex.ToString());
                
                // Write cells
                for (int colIndex = 0; colIndex < columns.Count; colIndex++)
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
                totalRows++;
                
                if (totalRows % 10000 == 0)
                {
                    Console.WriteLine($"Processed {totalRows} rows");
                    await xmlWriter.FlushAsync(); // Flush to ZIP stream
                }
            }
        }
        
        // Close elements
        await xmlWriter.WriteEndElementAsync(); // sheetData
        await xmlWriter.WriteEndElementAsync(); // worksheet
        await xmlWriter.WriteEndDocumentAsync();
        await xmlWriter.FlushAsync();
        
        return totalRows;
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