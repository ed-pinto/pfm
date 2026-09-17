using System.Globalization;
using System.Text;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pfm;

/// <summary>
/// One table of the structured configuration a run records: what the run calls it, and where the workbook states it.
/// </summary>
/// <param name="Name">The name the table is recorded and read back under.</param>
/// <param name="Worksheet">The worksheet that hosts the table.</param>
/// <param name="Table">The name of the table on that worksheet.</param>
internal sealed record TableSource(string Name, string Worksheet, string Table);

/// <summary>
/// One table of a run's configuration, as it was copied out of the workbook or read back from the file it was written
/// to.
/// </summary>
/// <remarks>
/// The values are the ones the workbook held, not text: a number is a double, a flag is a bool, and a cell that was
/// empty is null.  That is what lets a coalesced workbook put the table back the way the plan stated it rather than as
/// a sheet of strings.
/// </remarks>
public sealed class ConfigurationTable
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ConfigurationTable"/> class.
    /// </summary>
    /// <param name="name">The name the table is recorded under.</param>
    /// <param name="header">The column headings, in column order.</param>
    /// <param name="rows">The data rows, each holding one value per heading.</param>
    public ConfigurationTable(string name, IReadOnlyList<string> header, IReadOnlyList<object?[]> rows)
    {
        Name = name;
        Header = header;
        Rows = rows;
    }

    /// <summary>
    /// Gets the name the table is recorded under, ex. Expenditures.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Gets the column headings, in column order.
    /// </summary>
    public IReadOnlyList<string> Header { get; }

    /// <summary>
    /// Gets the data rows, in table order.  Each row holds one value per heading.
    /// </summary>
    public IReadOnlyList<object?[]> Rows { get; }

    /// <summary>
    /// Describes the table as lines of text: its name, then its headings, then its rows.
    /// </summary>
    /// <returns>The lines.</returns>
    /// <remarks>
    /// This is what a sweep's identifier is computed over, so it states the whole content of the table and states it
    /// the same way for a given table however that table reached this object.
    /// </remarks>
    public IEnumerable<string> Describe()
    {
        yield return Name;
        yield return string.Join(',', Header);

        foreach (object?[] row in Rows)
        {
            yield return string.Join(',', row.Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)));
        }
    }
}

/// <summary>
/// The configuration of the workbook a simulation exercised, recorded beside that simulation's results: a synopsis
/// written for a reader, and the tables of the plan itself written as data.
/// </summary>
/// <remarks>
/// A sweep drives a throwaway copy of the workbook and deletes it when the job ends, so the results file would
/// otherwise be all that survives of a run whose numbers only mean anything read against the plan that produced them.
/// The synopsis is maintained on 00_Overview as one line of plain text per configured parameter, which is what lets a
/// run record its configuration by copying it rather than by describing the model a second time in this tool, where it
/// would have to be kept in step with the workbook.
/// <para>
/// The synopsis is prose, which is what makes it readable and also what makes it useless to sort, filter or compare a
/// column of.  The tables beside it answer that: each is one table of the workbook copied out whole, as values rather
/// than as sentences, and <see cref="ParametersTableName"/> is the same thing for the scalar parameters, which the
/// workbook states as defined names rather than as a table.  Both are recorded, because they are for different
/// readers.
/// </para>
/// <para>
/// Every operation that drives a copy of the workbook records it the same way, so a back test and a Monte Carlo run
/// leave the same files, named the same, in their run directories.
/// </para>
/// </remarks>
public static class RunConfiguration
{
    /// <summary>
    /// The name of the file the synopsis is written to.
    /// </summary>
    public const string FileName = "RunConfiguration.txt";

    /// <summary>
    /// The name of the table of scalar parameters.  It is assembled from defined names rather than copied from a
    /// table of the workbook, because that is how 10_Parameters states them.
    /// </summary>
    public const string ParametersTableName = "Parameters";

    private const string SynopsisWorksheet = "00_Overview";
    private const string SynopsisTable = "Synopsis";
    private const string SynopsisColumn = "Line";

    private const string TableFilePrefix = "RunConfiguration.";
    private const string TableFileExtension = ".csv";

    private const string ParameterNameColumn = "Name";
    private const string ParameterValueColumn = "Value";

    /// <summary>
    /// The tables of the workbook a run copies, in the order they are recorded and read back.
    /// </summary>
    private static readonly TableSource[] TableSources =
    [
        new TableSource("Expenditures", "52_Expenditures", "Expenditures"),
        new TableSource("TargetNetIncomeEras", "51_Targets", "TargetNetIncomeEras"),
        new TableSource("AllocationTargets", "13_Allocations", "AllocationTargets"),
        new TableSource("AllocationFloors", "13_Allocations", "AllocationFloors"),
        new TableSource("DirectedDraws", "13_Allocations", "DirectedDraws")
    ];

    /// <summary>
    /// The defined names of 10_Parameters that make up the <see cref="ParametersTableName"/> table, in the order they
    /// are recorded.
    /// </summary>
    /// <remarks>
    /// These are the parameters a sweep is read against that the tables above do not carry.  A name that the workbook
    /// does not define fails the run rather than being recorded empty, which is what makes an incomplete record of a
    /// multi hour sweep impossible.
    /// </remarks>
    private static readonly string[] ParameterDefinedNames =
    [
        "CdnResidencyYear",
        "YearlyMaxDisposal",
        "EquityGainsRebalanceThreshold",
        "ConsumptionStressStart",
        "ConsumptionStressFull",
        "BootstrapBlockYears",
        "BootstrapMatchPool"
    ];

    /// <summary>
    /// Gets the names of the structured tables a run records, in the order they are written and read back.
    /// </summary>
    /// <remarks>
    /// The writer and the reader work from this one list rather than from what happens to be in a run directory, so
    /// the tables of a coalesced workbook are stacked in a stated order and a file that is missing is noticed.
    /// </remarks>
    public static IReadOnlyList<string> TableNames { get; } =
        [.. TableSources.Select(source => source.Name), ParametersTableName];

    /// <summary>
    /// Gets the name of the file a structured table is written to.
    /// </summary>
    /// <param name="tableName">The name of the table, which is one of <see cref="TableNames"/>.</param>
    /// <returns>The file name, ex. RunConfiguration.Expenditures.csv.</returns>
    /// <remarks>
    /// The prefix keeps these beside <see cref="FileName"/> in a listing of a run directory, and keeps them clear of
    /// the results file, which a sweep names after the command that ran.
    /// </remarks>
    public static string TableFileName(string tableName)
    {
        ArgumentNullException.ThrowIfNull(tableName);

        return TableFilePrefix + tableName + TableFileExtension;
    }

    /// <summary>
    /// Reads the synopsis of a workbook, one entry per row of the table in table order.
    /// </summary>
    /// <param name="workbook">The workbook whose configuration is wanted.</param>
    /// <returns>The lines of the synopsis.</returns>
    /// <remarks>
    /// Every row of the table is a formula, so what this returns is whatever the workbook last calculated.  A caller
    /// that has suspended calculation must therefore recalculate before reading, or it records the values the workbook
    /// was saved with rather than the ones the run will be driven from.
    /// </remarks>
    public static string[] Read(Excel.Workbook workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.ListObject synopsis = ExcelUtils.GetListObject(workbook, SynopsisWorksheet, SynopsisTable);
        try
        {
            object?[] values = ExcelUtils.ReadColumn(synopsis, SynopsisColumn);
            return Array.ConvertAll(values, FormatLine);
        }
        finally
        {
            ExcelUtils.ReleaseComObject(synopsis);
        }
    }

    /// <summary>
    /// Reads the synopsis of a workbook and writes it into a directory as <see cref="FileName"/>.
    /// </summary>
    /// <param name="workbook">The workbook whose configuration is being recorded.</param>
    /// <param name="directory">The directory to write the file into, ex. a job's run directory.</param>
    /// <returns>The path of the file that was written.</returns>
    public static string Write(Excel.Workbook workbook, string directory)
    {
        ArgumentNullException.ThrowIfNull(directory);

        string[] lines = Read(workbook);
        string path = Path.Combine(directory, FileName);

        // A byte order mark, unlike the results file, which has none: the synopsis carries whatever punctuation the
        // workbook wrote into it, en dashes included, and the mark is what makes a plain text viewer read those bytes
        // as UTF-8 rather than as the machine's ANSI code page.  WriteAllLines ends each line the way the platform
        // does, which on Windows is the CRLF such a viewer expects.
        File.WriteAllLines(path, lines, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

        Log.Logger.Information("PFM_RUN_CONFIGURATION: path=" + path + " lines=" + Format(lines.Length));

        return path;
    }

    /// <summary>
    /// Copies the structured configuration out of a workbook and writes it into a directory, one CSV file per table.
    /// </summary>
    /// <param name="workbook">The workbook whose configuration is being recorded.</param>
    /// <param name="directory">The directory to write the files into, ex. a job's run directory.</param>
    /// <returns>The paths of the files that were written, in <see cref="TableNames"/> order.</returns>
    /// <remarks>
    /// Like the synopsis, this reads what the workbook last calculated, so a caller that has suspended calculation
    /// must recalculate first.  It matters more here than there: several of these columns are formulas over the
    /// anchors beside them, ex. the inflated amount of an expenditure.
    /// </remarks>
    public static IReadOnlyList<string> WriteTables(Excel.Workbook workbook, string directory)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(directory);

        var paths = new List<string>(TableNames.Count);

        foreach (TableSource source in TableSources)
        {
            paths.Add(WriteTable(directory, ReadTable(workbook, source)));
        }

        paths.Add(WriteTable(directory, ReadParameters(workbook)));

        return paths;
    }

    /// <summary>
    /// Copies one table of the workbook.
    /// </summary>
    /// <param name="workbook">The workbook the table is read from.</param>
    /// <param name="source">Where the table is, and what it is recorded as.</param>
    /// <returns>The table.</returns>
    private static ConfigurationTable ReadTable(Excel.Workbook workbook, TableSource source)
    {
        Excel.ListObject table = ExcelUtils.GetListObject(workbook, source.Worksheet, source.Table);
        try
        {
            string[] header = ExcelUtils.GetColumnNames(table);
            object?[,] grid = ExcelUtils.ReadTableRows(table);

            int rowCount = grid.GetLength(0);
            int columnCount = Math.Min(header.Length, grid.GetLength(1));

            var rows = new List<object?[]>(rowCount);
            for (int row = 0; row < rowCount; row++)
            {
                var values = new object?[header.Length];
                for (int column = 0; column < columnCount; column++)
                {
                    values[column] = grid[row, column];
                }

                rows.Add(values);
            }

            return new ConfigurationTable(source.Name, header, rows);
        }
        finally
        {
            ExcelUtils.ReleaseComObject(table);
        }
    }

    /// <summary>
    /// Assembles the table of scalar parameters from the defined names of the workbook.
    /// </summary>
    /// <param name="workbook">The workbook the names are read from.</param>
    /// <returns>The table, one row per name.</returns>
    private static ConfigurationTable ReadParameters(Excel.Workbook workbook)
    {
        var rows = new List<object?[]>(ParameterDefinedNames.Length);

        foreach (string name in ParameterDefinedNames)
        {
            rows.Add([name, ExcelUtils.EvaluateName(workbook, name)]);
        }

        return new ConfigurationTable(ParametersTableName, [ParameterNameColumn, ParameterValueColumn], rows);
    }

    /// <summary>
    /// Writes one table into a directory as CSV.
    /// </summary>
    /// <param name="directory">The directory to write the file into.</param>
    /// <param name="table">The table to write.</param>
    /// <returns>The path of the file that was written.</returns>
    /// <remarks>
    /// A table with no rows is written as its heading row alone rather than not at all.  A plan that directs no draws
    /// is a plan, and the file saying so is what distinguishes it from a sweep whose configuration was never recorded.
    /// </remarks>
    private static string WriteTable(string directory, ConfigurationTable table)
    {
        string path = Path.Combine(directory, TableFileName(table.Name));

        // A byte order mark, as the synopsis has and the results file has not: these carry whatever the workbook holds
        // in a description or a note, and the mark is what makes a viewer read those bytes as UTF-8.  The reader
        // detects it and does not hand it back as part of the first heading.
        using (var writer = new StreamWriter(path, append: false,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)) { NewLine = "\r\n" })
        {
            writer.WriteLine(string.Join(',', table.Header.Select(FormatField)));

            foreach (object?[] row in table.Rows)
            {
                writer.WriteLine(string.Join(',', row.Select(FormatField)));
            }
        }

        Log.Logger.Information("PFM_RUN_CONFIGURATION_TABLE: path=" + path + " rows=" + Format(table.Rows.Count)
            + " columns=" + Format(table.Header.Count));

        return path;
    }

    /// <summary>
    /// Formats one cell of the synopsis as a line of the file.
    /// </summary>
    /// <param name="value">The value the cell holds.  Null when the cell is empty.</param>
    /// <returns>The line.</returns>
    /// <remarks>
    /// A synopsis cell is text, because every row of the table is a formula that concatenates one.  A cell that holds
    /// something else, ex. one whose formula returned an error, is written as Excel returned it rather than failing a
    /// run over one line of a report.
    /// </remarks>
    private static string FormatLine(object? value)
    {
        return value switch
        {
            null => string.Empty,
            string text => text,
            bool flag => flag ? "TRUE" : "FALSE",
            double number => number.ToString("R", CultureInfo.InvariantCulture),
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
        };
    }

    /// <summary>
    /// Formats one cell of a table as a field of the CSV file.
    /// </summary>
    /// <param name="value">The value the cell holds.  Null when the cell is empty.</param>
    /// <returns>The field, quoted when it has to be.</returns>
    /// <remarks>
    /// The same form the results are written in, and for the same reason: these files are read back into Excel, so a
    /// number is written to the shortest representation that reads back as the same double, a flag as TRUE or FALSE so
    /// that it imports as a boolean, and an empty cell as an empty field so that it imports as an empty cell.
    /// </remarks>
    private static string FormatField(object? value)
    {
        string text = value switch
        {
            null => string.Empty,
            string field => field,
            bool flag => flag ? "TRUE" : "FALSE",
            double number => number.ToString(CultureInfo.InvariantCulture),
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
        };

        return Quote(text);
    }

    /// <summary>
    /// Renders a field so that it reads back as the one value it is.
    /// </summary>
    /// <param name="text">The text of the field.</param>
    /// <returns>The field as it is written.</returns>
    /// <remarks>
    /// A line break inside a cell, ex. a note typed with Alt+Enter, becomes a space.  A field spanning a line break is
    /// not something the reader of these files supports, and one line of a note is worth more than a file that cannot
    /// be put back together.
    /// </remarks>
    private static string Quote(string text)
    {
        string field = text.ReplaceLineEndings(" ");

        return field.Contains(',', StringComparison.Ordinal) || field.Contains('"', StringComparison.Ordinal)
            ? "\"" + field.Replace("\"", "\"\"", StringComparison.Ordinal) + "\""
            : field;
    }

    /// <summary>
    /// Formats a count for a log message.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <returns>The formatted value.</returns>
    private static string Format(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }
}
