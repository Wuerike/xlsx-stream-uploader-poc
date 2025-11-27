using Amazon.Runtime;
using Amazon.S3;
using Amazon;
using XlsxStreamUploader.Services;
using DotNetEnv;


Console.WriteLine("Starting XlsxStreamUploader...");

Env.Load();

var accessKey = Environment.GetEnvironmentVariable("AWS_ACCESS_KEY_ID");
var secretKey = Environment.GetEnvironmentVariable("AWS_SECRET_ACCESS_KEY");

if (string.IsNullOrEmpty(accessKey) || string.IsNullOrEmpty(secretKey))
{
    Console.WriteLine("Error: AWS credentials not found in environment variables.");
    return;
}

RegionEndpoint BucketRegion = RegionEndpoint.USEast1;
var credentials = new BasicAWSCredentials(accessKey, secretKey);
var s3Client = new AmazonS3Client(credentials, BucketRegion);

var startTime = DateTime.Now;
var uploadService = new UploadService(s3Client);
await uploadService.StreamingS3CsvToXlsx(
    sourceCsvS3Key: "xlsx_poc/data.csv",
    targetXlsxS3Key: "xlsx_poc/UploadService.xlsx"
);
var endTime = DateTime.Now;
Console.WriteLine($"------------------------------------------------------------------");
Console.WriteLine($"---------------> UploadService completed in {endTime - startTime}");
Console.WriteLine($"------------------------------------------------------------------");


startTime = DateTime.Now;
var uploadServiceLargeXlsx = new LargeXlsxUploadService(s3Client);
await uploadServiceLargeXlsx.StreamingS3CsvToXlsx(
    sourceCsvS3Key: "xlsx_poc/data.csv",
    targetXlsxS3Key: "xlsx_poc/LargeXlsxUploadService.xlsx"
);
endTime = DateTime.Now;
Console.WriteLine($"------------------------------------------------------------------");
Console.WriteLine($"---------------> LargeXlsxUploadService completed in {endTime - startTime}");
Console.WriteLine($"------------------------------------------------------------------");


startTime = DateTime.Now;
var uploadServiceMiniExcel = new MiniExcelUploadService(s3Client);
await uploadServiceMiniExcel.StreamingS3CsvToXlsx(
    sourceCsvS3Key: "xlsx_poc/data.csv",
    targetXlsxS3Key: "xlsx_poc/MiniExcelUploadService.xlsx"
);
endTime = DateTime.Now;
Console.WriteLine($"------------------------------------------------------------------");
Console.WriteLine($"---------------> MiniExcelUploadService completed in {endTime - startTime}");
Console.WriteLine($"------------------------------------------------------------------");


startTime = DateTime.Now;
var uploadServiceSpreadCheetah = new SpreadCheetahUploadService(s3Client);
await uploadServiceSpreadCheetah.StreamingS3CsvToXlsx(
    sourceCsvS3Key: "xlsx_poc/data.csv",
    targetXlsxS3Key: "xlsx_poc/SpreadCheetahUploadService.xlsx"
);
endTime = DateTime.Now;
Console.WriteLine($"------------------------------------------------------------------");
Console.WriteLine($"---------------> SpreadCheetahUploadService completed in {endTime - startTime}");
Console.WriteLine($"------------------------------------------------------------------");

