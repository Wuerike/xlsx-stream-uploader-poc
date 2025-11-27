using Amazon.S3;
using Amazon.S3.Model;
using System.IO.Pipelines;
using MiniExcelLibs;

namespace XlsxStreamUploader.Services;

public class MiniExcelUploadService(IAmazonS3 s3Client) : BaseStreamingXlsxUploadService(s3Client)
{
    protected override string ProviderName => "MiniExcel";

    protected override async Task WriteCsvAsXlsxToPipe(string csvS3Key, PipeWriter writer)
    {
        var pipeStream = writer.AsStream();

        var values = EnumerateCsvRowsAsDictionaries(csvS3Key);

        pipeStream.SaveAs(values);

        await writer.CompleteAsync();
        Console.WriteLine("XLSX streaming with MiniExcel completed");
    }

    private IEnumerable<IDictionary<string, object?>> EnumerateCsvRowsAsDictionaries(string csvS3Key)
    {
        var request = new GetObjectRequest
        {
            BucketName = BucketName,
            Key = csvS3Key
        };

        using var response = S3Client.GetObjectAsync(request).GetAwaiter().GetResult();
        using var responseStream = response.ResponseStream;
        using var reader = new StreamReader(responseStream);
        string? line;
        while ((line = reader.ReadLine()) != null)
        {
            var columns = ParseCsvLine(line, reader);

            var dict = new Dictionary<string, object?>(columns.Count);
            for (int colIndex = 0; colIndex < columns.Count; colIndex++)
            {
                dict[$"Column{colIndex + 1}"] = columns[colIndex];
            }

            yield return dict;
        }
    }

}
