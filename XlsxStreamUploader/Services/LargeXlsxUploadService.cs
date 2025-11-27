using Amazon.S3;
using Amazon.S3.Model;
using System.IO.Pipelines;
using LargeXlsx;

namespace XlsxStreamUploader.Services;

public class LargeXlsxUploadService(IAmazonS3 s3Client) : BaseStreamingXlsxUploadService(s3Client)
{
    protected override string ProviderName => "LargeXlsx";

    protected override async Task WriteCsvAsXlsxToPipe(string csvS3Key, PipeWriter writer)
    {
        var pipeStream = writer.AsStream();

        await using (var xlsxWriter = new XlsxWriter(pipeStream, XlsxCompressionLevel.Fastest, requireCellReferences: true, skipInvalidCharacters: true))
        {
            xlsxWriter.BeginWorksheet("Data");

            int totalRows = await WriteWorksheetRowsFromCsvAsync(xlsxWriter, csvS3Key);

            Console.WriteLine($"Finished processing {totalRows} rows from CSV using LargeXlsx");
        }

        await writer.CompleteAsync();
        Console.WriteLine("XLSX streaming with LargeXlsx completed");
    }

    private async Task<int> WriteWorksheetRowsFromCsvAsync(XlsxWriter xlsxWriter, string csvS3Key)
    {
        int totalRows = 0;

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
                var columns = ParseCsvLine(line, reader);

                xlsxWriter.BeginRow();

                for (int colIndex = 0; colIndex < columns.Count; colIndex++)
                {
                    string cellValue = columns[colIndex];
                    xlsxWriter.Write(cellValue);
                }

                totalRows++;

                if (totalRows % 10000 == 0)
                {
                    Console.WriteLine($"Processed {totalRows} rows with LargeXlsx...");
                }
            }
        }

        return totalRows;
    }
}
