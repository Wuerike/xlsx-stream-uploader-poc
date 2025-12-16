using BenchmarkDotNet.Running;
using XlsxStreamUploaderLoadTest.Benchmarks;
using XlsxStreamUploaderLoadTest.Infrastructure;
using XlsxStreamUploaderLoadTest.Writers;

// For a quick test without full benchmark (Debug mode):
// dotnet run
//
// Run benchmarks in Release mode for accurate results:
// dotnet run -c Release

#if DEBUG
Console.WriteLine("⚠️  Running in DEBUG mode - results will not be accurate!");
Console.WriteLine("For proper benchmarks, run: dotnet run -c Release");
Console.WriteLine();
Console.WriteLine("Running quick validation test...");

const int rowCount = 100_000;
const int columnCount = 100;

Console.WriteLine($"Testing with {rowCount} rows x {columnCount} columns");
Console.WriteLine();

var writers = new (string Name, IXlsxWriter Writer)[]
{
    ("LargeXlsx", new LargeXlsxWriter()),
    ("LargeXlsx (NoCellRefs)", new LargeXlsxNoCellRefsWriter()),
    ("SpreadCheetah", new SpreadCheetahWriter()),
    ("SpreadCheetah (ReusableArray)", new SpreadCheetahReusableArrayWriter()),
    ("SpreadCheetah (DataCell)", new SpreadCheetahDataCellWriter()),
    ("SpreadCheetah (Buffered)", new SpreadCheetahBufferedWriter()),
    ("MiniExcel", new MiniExcelWriter()),
    ("ManualOpenXml", new ManualOpenXmlWriter())
};

var sw = new System.Diagnostics.Stopwatch();

foreach (var (name, writer) in writers)
{
    sw.Restart();
    await using var stream = new NullStream();
    await writer.WriteAsync(stream, rowCount, columnCount);
    Console.WriteLine($"{name}: {sw.ElapsedMilliseconds}ms");
}

Console.WriteLine();
Console.WriteLine("✅ Quick test completed. Run with -c Release for real benchmarks.");
#else
BenchmarkRunner.Run<XlsxWriterBenchmark>();
#endif
