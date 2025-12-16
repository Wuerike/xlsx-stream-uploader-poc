namespace XlsxStreamUploaderLoadTest.Generators;

/// <summary>
/// Generates fake CSV lines for benchmarking without I/O dependencies.
/// </summary>
public static class CsvLineGenerator
{
    /// <summary>
    /// Generates a deterministic CSV line with the specified number of columns.
    /// Each column contains "Value_rowIndex_colIndex" to simulate realistic string lengths.
    /// </summary>
    public static string GenerateLine(int rowIndex, int columnCount)
    {
        var columns = new string[columnCount];
        for (int col = 0; col < columnCount; col++)
        {
            columns[col] = $"Value_{rowIndex}_{col}";
        }
        return string.Join(",", columns);
    }

    /// <summary>
    /// Generates a header line with column names.
    /// </summary>
    public static string GenerateHeader(int columnCount)
    {
        var columns = new string[columnCount];
        for (int col = 0; col < columnCount; col++)
        {
            columns[col] = $"Column{col}";
        }
        return string.Join(",", columns);
    }

    /// <summary>
    /// Parses a CSV line into columns. Simple implementation for benchmarking.
    /// Does not handle quoted fields with commas (not needed for our generated data).
    /// </summary>
    public static string[] ParseLine(string line)
    {
        return line.Split(',');
    }

    /// <summary>
    /// Enumerates CSV lines as an async stream, simulating reading from a source.
    /// </summary>
    public static async IAsyncEnumerable<string> EnumerateLinesAsync(int rowCount, int columnCount)
    {
        yield return GenerateHeader(columnCount);

        for (int row = 0; row < rowCount; row++)
        {
            yield return GenerateLine(row, columnCount);

            // Yield control occasionally to simulate async I/O
            if (row % 1000 == 0)
            {
                await Task.Yield();
            }
        }
    }

    /// <summary>
    /// Enumerates CSV lines synchronously.
    /// </summary>
    public static IEnumerable<string> EnumerateLines(int rowCount, int columnCount)
    {
        yield return GenerateHeader(columnCount);

        for (int row = 0; row < rowCount; row++)
        {
            yield return GenerateLine(row, columnCount);
        }
    }
}
