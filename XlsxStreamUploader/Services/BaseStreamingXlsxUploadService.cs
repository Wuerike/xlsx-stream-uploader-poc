using System.IO.Pipelines;
using System.Text;
using Amazon.S3;
using Amazon.S3.Model;

namespace XlsxStreamUploader.Services;

public interface IXlsxStreamingUploadService
{
    Task StreamingS3CsvToXlsx(string sourceCsvS3Key, string targetXlsxS3Key);
}

public abstract class BaseStreamingXlsxUploadService(IAmazonS3 s3Client) : IXlsxStreamingUploadService
{
    protected readonly IAmazonS3 S3Client = s3Client;
    protected readonly string BucketName = Environment.GetEnvironmentVariable("S3_BUCKET_NAME")
            ?? throw new InvalidOperationException("S3_BUCKET_NAME environment variable is not set");

    protected abstract string ProviderName { get; }

    public async Task StreamingS3CsvToXlsx(string sourceCsvS3Key, string targetXlsxS3Key)
    {
        try
        {
            Console.WriteLine($"Starting {ProviderName} streaming conversion: s3://{BucketName}/{sourceCsvS3Key} to s3://{BucketName}/{targetXlsxS3Key}");

            var pipe = new Pipe();

            var uploadId = await InitiateMultipartUpload(targetXlsxS3Key);
            var partETags = new List<PartETag>();

            try
            {
                var writerTask = Task.Run(async () =>
                {
                    try
                    {
                        await WriteCsvAsXlsxToPipe(sourceCsvS3Key, pipe.Writer);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error in {ProviderName} writer task: {ex.Message}");
                        throw;
                    }
                });

                var uploaderTask = Task.Run(async () =>
                {
                    try
                    {
                        await UploadFromPipeToS3(pipe.Reader, uploadId, targetXlsxS3Key, partETags);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error in uploader task: {ex.Message}");
                        throw;
                    }
                });

                await Task.WhenAll(writerTask, uploaderTask);

                await CompleteMultipartUpload(uploadId, partETags, targetXlsxS3Key);

                Console.WriteLine($"Streaming conversion with {ProviderName} completed successfully!");
            }
            catch (Exception)
            {
                await AbortMultipartUpload(uploadId, targetXlsxS3Key);
                throw;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in {ProviderName} streaming S3 CSV to XLSX: {ex.Message}");
            throw;
        }
    }

    protected abstract Task WriteCsvAsXlsxToPipe(string csvS3Key, PipeWriter writer);

    protected async Task UploadFromPipeToS3(PipeReader reader, string uploadId, string s3Key, List<PartETag> partETags)
    {
        const int chunkSize = 20 * 1024 * 1024;
        int partNumber = 1;
        var buffer = new MemoryStream();

        while (true)
        {
            var result = await reader.ReadAsync();
            var readBuffer = result.Buffer;

            foreach (var segment in readBuffer)
            {
                await buffer.WriteAsync(segment);

                if (buffer.Length >= chunkSize)
                {
                    buffer.Position = 0;
                    await UploadPart(buffer, uploadId, partNumber, partETags, s3Key);
                    partNumber++;
                    buffer = new MemoryStream();
                }
            }

            reader.AdvanceTo(readBuffer.End);

            if (result.IsCompleted)
            {
                break;
            }
        }

        if (buffer.Length > 0)
        {
            buffer.Position = 0;
            await UploadPart(buffer, uploadId, partNumber, partETags, s3Key);
        }

        await reader.CompleteAsync();
        Console.WriteLine($"Uploaded {partNumber} parts to S3 using {ProviderName} pipeline");
    }

    protected static List<string> ParseCsvLine(string line, StreamReader reader)
    {
        var columns = new List<string>(20);
        var currentField = new StringBuilder(256);
        bool insideQuotes = false;

        ReadOnlySpan<char> span = line.AsSpan();
        int i = 0;

        while (i < span.Length)
        {
            char c = span[i];

            if (c == '"')
            {
                if (insideQuotes)
                {
                    if (i + 1 < span.Length && span[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i += 2;
                        continue;
                    }
                    else
                    {
                        insideQuotes = false;
                        i++;
                        continue;
                    }
                }
                else
                {
                    insideQuotes = true;
                    i++;
                    continue;
                }
            }
            else if (c == ',' && !insideQuotes)
            {
                columns.Add(currentField.ToString());
                currentField.Clear();
                i++;
                continue;
            }
            else
            {
                currentField.Append(c);
                i++;
            }
        }

        while (insideQuotes && !reader.EndOfStream)
        {
            string? nextLine = reader.ReadLine();
            if (nextLine == null) break;

            currentField.Append('\n');
            ReadOnlySpan<char> nextSpan = nextLine.AsSpan();

            for (i = 0; i < nextSpan.Length; i++)
            {
                char c = nextSpan[i];

                if (c == '"')
                {
                    if (i + 1 < nextSpan.Length && nextSpan[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i++;
                        continue;
                    }
                    else
                    {
                        insideQuotes = false;
                        i++;
                        if (i < nextSpan.Length && nextSpan[i] == ',')
                        {
                            columns.Add(currentField.ToString());
                            currentField.Clear();
                            i++;
                        }

                        while (i < nextSpan.Length)
                        {
                            c = nextSpan[i];
                            if (c == ',')
                            {
                                columns.Add(currentField.ToString());
                                currentField.Clear();
                            }
                            else if (c == '"')
                            {
                                insideQuotes = true;
                                i++;
                                break;
                            }
                            else
                            {
                                currentField.Append(c);
                            }
                            i++;
                        }
                        break;
                    }
                }
                else
                {
                    currentField.Append(c);
                }
            }
        }

        columns.Add(currentField.ToString());

        return columns;
    }

    protected async Task<string> InitiateMultipartUpload(string s3Key)
    {
        var request = new InitiateMultipartUploadRequest
        {
            BucketName = BucketName,
            Key = s3Key
        };

        var response = await S3Client.InitiateMultipartUploadAsync(request);
        Console.WriteLine($"Initiated multipart upload with ID: {response.UploadId} ({ProviderName})");
        return response.UploadId;
    }

    protected async Task UploadPart(MemoryStream stream, string uploadId, int partNumber, List<PartETag> partETags, string s3Key)
    {
        stream.Position = 0;

        var request = new UploadPartRequest
        {
            BucketName = BucketName,
            Key = s3Key,
            UploadId = uploadId,
            PartNumber = partNumber,
            InputStream = stream,
            PartSize = stream.Length
        };

        var response = await S3Client.UploadPartAsync(request);
        partETags.Add(new PartETag { PartNumber = partNumber, ETag = response.ETag });

        Console.WriteLine($"Uploaded part {partNumber} ({ProviderName}), Size: {stream.Length / (1024 * 1024):F2}MB, ETag: {response.ETag}");
    }

    protected async Task CompleteMultipartUpload(string uploadId, List<PartETag> partETags, string s3Key)
    {
        var request = new CompleteMultipartUploadRequest
        {
            BucketName = BucketName,
            Key = s3Key,
            UploadId = uploadId,
            PartETags = partETags
        };

        await S3Client.CompleteMultipartUploadAsync(request);
        Console.WriteLine($"Completed multipart upload with {partETags.Count} parts ({ProviderName})");
    }

    protected async Task AbortMultipartUpload(string uploadId, string s3Key)
    {
        var request = new AbortMultipartUploadRequest
        {
            BucketName = BucketName,
            Key = s3Key,
            UploadId = uploadId
        };

        await S3Client.AbortMultipartUploadAsync(request);
        Console.WriteLine($"Aborted multipart upload with ID: {uploadId} ({ProviderName})");
    }
}
