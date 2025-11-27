using Amazon.S3;
using Amazon.S3.Model;
using System.IO.Pipelines;
using SpreadCheetah;

namespace XlsxStreamUploader.Services;

public class SpreadCheetahUploadService(IAmazonS3 s3Client) : BaseStreamingXlsxUploadService(s3Client)
{
    protected override string ProviderName => "SpreadCheetah";

    protected override async Task WriteCsvAsXlsxToPipe(string csvS3Key, PipeWriter writer)
    {
        var pipeStream = writer.AsStream();

        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(pipeStream))
        {
            await spreadsheet.StartWorksheetAsync("Data");

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

                    var row = new List<Cell>(columns.Count);
                    for (int colIndex = 0; colIndex < columns.Count; colIndex++)
                    {
                        string cellValue = columns[colIndex];
                        row.Add(new Cell(cellValue));
                    }

                    await spreadsheet.AddRowAsync(row);

                    totalRows++;

                    if (totalRows % 10000 == 0)
                    {
                        Console.WriteLine($"Processed {totalRows} rows with SpreadCheetah...");
                    }
                }
            }

            await spreadsheet.FinishAsync();

            Console.WriteLine($"Finished processing {totalRows} rows from CSV using SpreadCheetah");
        }

        await writer.CompleteAsync();
        Console.WriteLine("XLSX streaming with SpreadCheetah completed");
    }

}
