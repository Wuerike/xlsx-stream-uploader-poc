using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;
using XlsxStreamUploaderLoadTest.Infrastructure;
using XlsxStreamUploaderLoadTest.Writers;

namespace XlsxStreamUploaderLoadTest.Benchmarks;

/// <summary>
/// Benchmark comparing all XLSX streaming libraries: LargeXlsx, SpreadCheetah, MiniExcel, and Manual OpenXML.
/// Focus: memory allocations per operation.
/// </summary>
[Config(typeof(Config))]
[MemoryDiagnoser]
public class XlsxWriterBenchmark
{
    private class Config : ManualConfig
    {
        public Config()
        {
            AddJob(Job.ShortRun
                .WithWarmupCount(2)
                .WithIterationCount(5));

            AddColumn(StatisticColumn.AllStatistics);
        }
    }

    // Writers - instantiated once per benchmark class
    private readonly LargeXlsxWriter _largeXlsxWriter = new();
    private readonly LargeXlsxNoCellRefsWriter _largeXlsxNoCellRefsWriter = new();
    private readonly SpreadCheetahWriter _spreadCheetahWriter = new();
    private readonly SpreadCheetahReusableArrayWriter _spreadCheetahReusableArrayWriter = new();
    private readonly SpreadCheetahDataCellWriter _spreadCheetahDataCellWriter = new();
    private readonly SpreadCheetahBufferedWriter _spreadCheetahBufferedWriter = new();
    private readonly MiniExcelWriter _miniExcelWriter = new();
    private readonly ManualOpenXmlWriter _manualOpenXmlWriter = new();

    /// <summary>
    /// Number of rows to write (excluding header).
    /// </summary>
    [Params(10_000, 100_000)]
    public int RowCount { get; set; }

    /// <summary>
    /// Number of columns per row.
    /// </summary>
    [Params(10, 50)]
    public int ColumnCount { get; set; }

    [Benchmark(Baseline = true)]
    public async Task LargeXlsx_Write()
    {
        using var stream = new NullStream();
        await _largeXlsxWriter.WriteAsync(stream, RowCount, ColumnCount);
    }

    [Benchmark]
    public async Task LargeXlsx_NoCellRefs()
    {
        using var stream = new NullStream();
        await _largeXlsxNoCellRefsWriter.WriteAsync(stream, RowCount, ColumnCount);
    }

    [Benchmark]
    public async Task SpreadCheetah_Write()
    {
        await using var stream = new NullStream();
        await _spreadCheetahWriter.WriteAsync(stream, RowCount, ColumnCount);
    }

    [Benchmark]
    public async Task SpreadCheetah_ReusableArray()
    {
        await using var stream = new NullStream();
        await _spreadCheetahReusableArrayWriter.WriteAsync(stream, RowCount, ColumnCount);
    }

    [Benchmark]
    public async Task SpreadCheetah_DataCell()
    {
        await using var stream = new NullStream();
        await _spreadCheetahDataCellWriter.WriteAsync(stream, RowCount, ColumnCount);
    }

    [Benchmark]
    public async Task SpreadCheetah_Buffered()
    {
        await using var stream = new NullStream();
        await _spreadCheetahBufferedWriter.WriteAsync(stream, RowCount, ColumnCount);
    }

    [Benchmark]
    public async Task MiniExcel_Write()
    {
        using var stream = new NullStream();
        await _miniExcelWriter.WriteAsync(stream, RowCount, ColumnCount);
    }

    [Benchmark]
    public async Task ManualOpenXml_Write()
    {
        await using var stream = new NullStream();
        await _manualOpenXmlWriter.WriteAsync(stream, RowCount, ColumnCount);
    }
}
