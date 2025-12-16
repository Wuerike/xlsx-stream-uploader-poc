# CSV to XLSX streaming services

This folder contains several streaming implementations that convert a CSV file stored in Amazon S3 into an XLSX file also stored in S3.

All implementations share the same high-level idea:

- Read the CSV object from S3 as a stream.
- Convert rows into XLSX content while streaming.
- Stream the XLSX bytes into a pipe.
- Upload the XLSX to S3 using multipart upload, reading directly from the pipe.

The common behavior is in `BaseStreamingXlsxUploadService` and each concrete service picks a different XLSX writer strategy.

---

## Shared base: `BaseStreamingXlsxUploadService`

Key types:

- `IXlsxStreamingUploadService`
  - Defines `Task StreamingS3CsvToXlsx(string sourceCsvS3Key, string targetXlsxS3Key)`.
- `BaseStreamingXlsxUploadService`
  - Depends on `IAmazonS3` and reads bucket name from `S3_BUCKET_NAME`.
  - Exposes a `ProviderName` used only for logging.

### Orchestration

`StreamingS3CsvToXlsx` does the following:

1. Logs the start of the conversion with source and target S3 keys.
2. Creates a `Pipe` with a `PipeWriter` and `PipeReader`.
3. Initiates a multipart upload on S3 for the XLSX target object.
4. Starts two tasks in parallel:
   - **Writer task**
     - Calls `WriteCsvAsXlsxToPipe(sourceCsvS3Key, pipe.Writer)`.
     - Concrete services implement this method and are responsible for:
       - Reading CSV from S3.
       - Converting rows into XLSX bytes.
       - Writing those bytes into the `PipeWriter`.
   - **Uploader task**
     - Calls `UploadFromPipeToS3(pipe.Reader, uploadId, targetXlsxS3Key, partETags)`.
     - Reads from the `PipeReader`, accumulates bytes into chunks and uploads each chunk as an S3 multipart part.
5. Waits for both tasks to finish, then completes the multipart upload.
6. On error, aborts the multipart upload and rethrows the exception.

### Multipart upload helpers

The base class also provides:

- `UploadFromPipeToS3`
  - Reads from the pipe in a loop.
  - Buffers data until a chunk reaches 20 MB, then calls `UploadPart`.
  - Tracks `PartETag` values for the multipart upload.
  - At the end, uploads any remaining data and completes the pipe.

- `InitiateMultipartUpload`
  - Creates the multipart upload for the XLSX target object and returns the `uploadId`.

- `UploadPart`
  - Uploads a single multipart part from a `MemoryStream`.

- `CompleteMultipartUpload`
  - Finalizes the multipart upload with the collected part ETags.

- `AbortMultipartUpload`
  - Aborts an in-progress multipart upload if something fails.

- `ParseCsvLine`
  - A CSV parser that supports quoted fields and delimiters following RFC 4180 style rules.
  - Used by most of the implementations to turn each CSV line into a list of columns.

Concrete services only need to implement `WriteCsvAsXlsxToPipe` and can reuse all streaming and S3 logic from the base class.

---

## Implementation: `UploadService` (custom XLSX writer)

`UploadService` is a custom implementation that writes XLSX files without any third-party XLSX library.

### How it writes XLSX

1. Creates a `ZipArchive` on top of the pipe stream.
2. Adds the minimal set of Open XML parts required by Excel:
   - `[Content_Types].xml`
   - `_rels/.rels`
   - `xl/_rels/workbook.xml.rels`
   - `xl/workbook.xml`
   - `xl/worksheets/sheet1.xml`
3. Uses `XmlWriter` to write `xl/worksheets/sheet1.xml` in a streaming way:
   - Starts `<worksheet>` and `<sheetData>` elements.
   - Reads CSV from S3 line by line.
   - Uses `ParseCsvLine` to split the line into columns.
   - For each row:
     - Writes a `<row>` element with a row index.
     - Writes one `<c>` (cell) element per column with `t="inlineStr"` and an `<is><t>` value.
   - Flushes periodically when processing many rows.
   - Closes the XML elements at the end.

### Trade-offs

- **Pros**
  - No XLSX library dependency.
  - Full control over the generated XML and structure.
- **Cons**
  - More verbose and low-level code.
  - You need to understand the Open XML schema parts being created.

---

## Implementation: `LargeXlsxUploadService`

`LargeXlsxUploadService` uses the **LargeXlsx** library to create the XLSX workbook.

### How it writes XLSX

1. Creates an `XlsxWriter` on top of the pipe stream:
   - Compression level: fastest.
   - `requireCellReferences: true`.
   - `skipInvalidCharacters: true`.
2. Starts a worksheet named `Data`.
3. Calls a helper method `WriteWorksheetRowsFromCsvAsync`:
   - Reads the CSV object from S3 asynchronously.
   - For each line:
     - Uses `ParseCsvLine` to get the columns.
     - Calls `BeginRow()`.
     - Writes each column as a cell with `xlsxWriter.Write(cellValue)`.
   - Logs progress every 10,000 rows.

### Trade-offs

- **Pros**
  - Library is designed for large XLSX files and streaming scenarios.
  - Less manual XML handling.
- **Cons**
  - Requires the LargeXlsx dependency.

---

## Implementation: `MiniExcelUploadService`

`MiniExcelUploadService` uses the **MiniExcel** library and its `SaveAs` API on top of the pipe stream.

### How it writes XLSX

1. Builds an `IEnumerable<IDictionary<string, object?>>` with one dictionary per CSV row:
   - Reads the CSV object from S3.
   - Uses `ParseCsvLine` to obtain columns.
   - Creates a dictionary with keys `Column1`, `Column2`, etc. and the corresponding cell values.
   - Yields each dictionary.
2. Calls `pipeStream.SaveAs(values)`:
   - MiniExcel iterates the enumerable and writes an XLSX workbook to the stream.
3. Completes the writer and logs completion.

### Trade-offs

- **Pros**
  - Very simple high-level API.
  - Easy to understand: each row is just a dictionary.
- **Cons**
  - Allocates a dictionary per row.
  - The S3 read path uses `GetObjectAsync(...).GetAwaiter().GetResult()`, which is synchronous.

---

## Implementation: `SpreadCheetahUploadService`

`SpreadCheetahUploadService` uses the **SpreadCheetah** library to build the XLSX workbook.

### How it writes XLSX

1. Creates a `Spreadsheet` on top of the pipe stream.
2. Starts a worksheet named `Data`.
3. Reads the CSV object from S3 asynchronously.
4. For each line:
   - Uses `ParseCsvLine` to split into columns.
   - Builds a `List<Cell>` with one `Cell` per column.
   - Calls `AddRowAsync(row)` to append the row to the worksheet.
   - Logs progress every 10,000 rows.
5. Calls `FinishAsync()` to finalize the workbook.
6. Completes the writer and logs completion.

### Trade-offs

- **Pros**
  - Asynchronous API with good streaming support.
  - Uses a typed `Cell` model, leaving Open XML details to the library.
- **Cons**
  - Allocates a list of cells per row.
  - Requires the SpreadCheetah dependency.

---

## Sample execution

The following results come from running all four implementations against a ~250 MB CSV file stored in S3.

- **Input**
  - S3 object: `s3://{bucket}/xlsx_poc/data.csv`
  - Approximate size: 250 MB

- **Output XLSX files** (same bucket/prefix)
  - `s3://{bucket}/xlsx_poc/UploadService.xlsx`
    - Size: ~169.6 MB
    - Elapsed time: ~00:01:24
  - `s3://{bucket}/xlsx_poc/LargeXlsxUploadService.xlsx`
    - Size: ~170.9 MB
    - Elapsed time: ~00:01:09
  - `s3://{bucket}/xlsx_poc/MiniExcelUploadService.xlsx`
    - Size: ~171.8 MB
    - Elapsed time: ~00:01:10
  - `s3://{bucket}/xlsx_poc/SpreadCheetahUploadService.xlsx`
    - Size: ~57.5 MB
    - Elapsed time: ~00:00:57

---

## Load Test (Benchmark)

The `XlsxStreamUploaderLoadTest` project uses **BenchmarkDotNet** to measure performance and memory allocations of different XLSX writing strategies.

### How to run

```bash
# Quick validation test (Debug mode, ~30 seconds)
make loadtest-quick

# Full benchmark (Release mode, ~5 minutes)
make loadtest
```

### What is tested

The benchmark isolates **XLSX generation performance** by:
- Using a `NullStream` that discards all written bytes (no disk I/O)
- Generating deterministic CSV data in memory (no S3 dependency)
- Measuring only the XLSX writing time and memory allocations

### Writers tested

| Writer | Description |
|--------|-------------|
| **LargeXlsx_Write** | LargeXlsx library with default settings |
| **LargeXlsx_NoCellRefs** | LargeXlsx with cell references disabled |
| **SpreadCheetah_Write** | SpreadCheetah with `List<Cell>` per row |
| **SpreadCheetah_ReusableArray** | SpreadCheetah with reusable `Cell[]` array |
| **SpreadCheetah_DataCell** | SpreadCheetah with reusable `DataCell[]` (struct) |
| **SpreadCheetah_Buffered** | SpreadCheetah with 4MB buffer + reusable `DataCell[]` |
| **MiniExcel_Write** | MiniExcel with `Dictionary` per row |
| **ManualOpenXml_Write** | Manual ZipArchive + XmlWriter (no library) |

### Test scenarios

| RowCount | ColumnCount | Description |
|----------|-------------|-------------|
| 10,000 | 10 | Small dataset |
| 10,000 | 50 | Small dataset, wide rows |
| 100,000 | 10 | Large dataset |
| 100,000 | 50 | Large dataset, wide rows |

### Benchmark results

Environment:
- Ubuntu 22.04.5 LTS
- 12th Gen Intel Core i7-1265U, 12 logical cores
- .NET 8.0.21, X64 RyuJIT AVX2

#### Summary (100K rows × 50 columns - heaviest scenario)

| Writer | Time | vs Baseline | Memory | Alloc Ratio |
|--------|------|-------------|--------|-------------|
| LargeXlsx_Write | 1,256 ms | - | 736 MB | 1.00 |
| **SpreadCheetah_DataCell** | **957 ms** | **24% faster** | **736 MB** | **1.00** |
| **SpreadCheetah_Buffered** | **961 ms** | **24% faster** | **736 MB** | **1.00** |
| SpreadCheetah_ReusableArray | 1,018 ms | 19% faster | 736 MB | 1.00 |
| SpreadCheetah_Write | 1,037 ms | 17% faster | 970 MB | 1.32 |
| LargeXlsx_NoCellRefs | 1,212 ms | 4% faster | 736 MB | 1.00 |
| MiniExcel_Write | 3,147 ms | 2.5x slower | 3,340 MB | 4.54 |
| ManualOpenXml_Write | 3,222 ms | 2.6x slower | 1,170 MB | 1.59 |

#### Full results

| Method | RowCount | ColumnCount | Mean | Allocated | Alloc Ratio |
|--------|----------|-------------|------|-----------|-------------|
| LargeXlsx_Write | 10000 | 10 | 25.02 ms | 13.89 MB | 1.00 |
| LargeXlsx_NoCellRefs | 10000 | 10 | 19.74 ms | 13.89 MB | 1.00 |
| SpreadCheetah_Write | 10000 | 10 | 20.81 ms | 18.91 MB | 1.36 |
| SpreadCheetah_ReusableArray | 10000 | 10 | 21.73 ms | 13.8 MB | 0.99 |
| SpreadCheetah_DataCell | 10000 | 10 | 20.97 ms | 13.8 MB | 0.99 |
| SpreadCheetah_Buffered | 10000 | 10 | 21.40 ms | 13.8 MB | 0.99 |
| MiniExcel_Write | 10000 | 10 | 68.93 ms | 78.84 MB | 5.67 |
| ManualOpenXml_Write | 10000 | 10 | 62.97 ms | 19.51 MB | 1.40 |
| | | | | | |
| LargeXlsx_Write | 10000 | 50 | 128.16 ms | 67.22 MB | 1.00 |
| LargeXlsx_NoCellRefs | 10000 | 50 | 112.03 ms | 67.22 MB | 1.00 |
| SpreadCheetah_Write | 10000 | 50 | 113.29 ms | 90.55 MB | 1.35 |
| SpreadCheetah_ReusableArray | 10000 | 50 | 97.64 ms | 67.15 MB | 1.00 |
| SpreadCheetah_DataCell | 10000 | 50 | 99.47 ms | 67.13 MB | 1.00 |
| SpreadCheetah_Buffered | 10000 | 50 | 106.28 ms | 67.13 MB | 1.00 |
| MiniExcel_Write | 10000 | 50 | 322.92 ms | 333.62 MB | 4.96 |
| ManualOpenXml_Write | 10000 | 50 | 353.74 ms | 108.67 MB | 1.62 |
| | | | | | |
| LargeXlsx_Write | 100000 | 10 | 255.56 ms | 140.24 MB | 1.00 |
| LargeXlsx_NoCellRefs | 100000 | 10 | 209.05 ms | 140.24 MB | 1.00 |
| SpreadCheetah_Write | 100000 | 10 | 248.22 ms | 191.26 MB | 1.36 |
| SpreadCheetah_ReusableArray | 100000 | 10 | 211.43 ms | 140.14 MB | 1.00 |
| SpreadCheetah_DataCell | 100000 | 10 | 196.90 ms | 140.14 MB | 1.00 |
| SpreadCheetah_Buffered | 100000 | 10 | 199.07 ms | 140.15 MB | 1.00 |
| MiniExcel_Write | 100000 | 10 | 607.76 ms | 669.4 MB | 4.77 |
| ManualOpenXml_Write | 100000 | 10 | 635.05 ms | 203.53 MB | 1.45 |
| | | | | | |
| LargeXlsx_Write | 100000 | 50 | 1,256.21 ms | 736.01 MB | 1.00 |
| LargeXlsx_NoCellRefs | 100000 | 50 | 1,211.81 ms | 736.01 MB | 1.00 |
| SpreadCheetah_Write | 100000 | 50 | 1,036.68 ms | 970.22 MB | 1.32 |
| SpreadCheetah_ReusableArray | 100000 | 50 | 1,017.92 ms | 735.92 MB | 1.00 |
| SpreadCheetah_DataCell | 100000 | 50 | 957.22 ms | 735.92 MB | 1.00 |
| SpreadCheetah_Buffered | 100000 | 50 | 960.51 ms | 735.92 MB | 1.00 |
| MiniExcel_Write | 100000 | 50 | 3,147.13 ms | 3340.02 MB | 4.54 |
| ManualOpenXml_Write | 100000 | 50 | 3,221.53 ms | 1170.22 MB | 1.59 |

### Conclusion

**Recommended: SpreadCheetah with `DataCell[]` reusable array**

- **24% faster** than LargeXlsx
- **Same memory allocation** as LargeXlsx
- Async-friendly API

```csharp
var cellBuffer = new DataCell[columnCount];

foreach (var line in csvLines)
{
    var columns = ParseLine(line);
    for (int i = 0; i < columns.Length; i++)
        cellBuffer[i] = new DataCell(columns[i]);
    
    await spreadsheet.AddRowAsync(cellBuffer.AsMemory(0, columns.Length));
}
```
