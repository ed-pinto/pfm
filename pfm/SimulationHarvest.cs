using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pfm;

/// <summary>
/// What one simulation produced: the per year outcome metrics harvested from 56_Summary, and how the tax provision
/// behind them settled.
/// </summary>
public sealed class SimulationRecord
{
    /// <summary>
    /// Gets the zero based index of the simulation within the whole sweep, not within the job that ran it.  It is what
    /// puts the slices of a parallel sweep back in order.
    /// </summary>
    public required int Index { get; init; }

    /// <summary>
    /// Gets the year of 21_EconomyHistorical the projection was driven from.
    /// </summary>
    public required int StartYear { get; init; }

    /// <summary>
    /// Gets the semester of <see cref="StartYear"/> the projection was driven from.
    /// </summary>
    public required int StartSemester { get; init; }

    /// <summary>
    /// Gets the harvested values, indexed by data element and then by simulated year.  A value that was not a number,
    /// ex. a cell that resolved to an error, is null.
    /// </summary>
    public required IReadOnlyList<double?[]> Elements { get; init; }

    /// <summary>
    /// Gets how the tax provision of this simulation settled.
    /// </summary>
    public required TaxProvisionOutcome TaxProvision { get; init; }
}

/// <summary>
/// Harvests the outcome metrics of one simulation from the SummaryAnnual table on 56_Summary.
/// </summary>
/// <remarks>
/// 56_Summary is the harvest surface the workbook is designed around: one row per calendar year, carrying portfolio
/// value, income, tax and year end sleeve composition. Only its Projected years are harvested. A Closed year reports
/// what actually happened and a Mixed year is only half modelled, so neither is a simulation outcome, and the workbook
/// itself directs a harvester to filter on Regime before reading.
/// <para>
/// The harvester resolves the table and its column positions once and then reads the whole data body per simulation, so
/// harvesting ten scattered columns costs one cross process call rather than ten.
/// </para>
/// </remarks>
public sealed class SummaryHarvester : IDisposable
{
    private const string SummaryWorksheet = "56_Summary";
    private const string SummaryTable = "SummaryAnnual";
    private const string YearColumn = "Year";
    private const string RegimeColumn = "Regime";
    private const string ProjectedStatusName = "StatusProjected";

    /// <summary>
    /// The columns of SummaryAnnual a simulation harvests, in the order they appear in the results.
    /// </summary>
    /// <remarks>
    /// Back testing and Monte Carlo harvest the same elements deliberately: the two sweeps answer the same questions of
    /// different market paths, and their results are only comparable if they are measured the same way.
    /// </remarks>
    public static readonly IReadOnlyList<string> DataElements =
    [
        "InflationIndex",
        "NetIncome",
        "EndingPortfolioValue",
        "TotalTax",
        "PortfolioConsumedRatio",
        "RatioBondCash",
        "RatioBondLadder",
        "RatioEquityConcentrated",
        "RatioEquityCore",
        "RatioEquityInternational"
    ];

    private readonly int[] _elementColumns;
    private readonly int _yearColumn;
    private readonly int _regimeColumn;
    private readonly string _projectedStatus;
    private readonly int[] _projectedRows;
    private readonly int[] _years;

    private Excel.ListObject? _summary;
    private bool _disposed;

    /// <summary>
    /// Initializes a new instance of the <see cref="SummaryHarvester"/> class and reads the shape of the harvest: which
    /// rows are projected, and therefore which years every simulation reports.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <remarks>
    /// The caller must have settled the workbook first, because the regime of a year is itself calculated.
    /// </remarks>
    public SummaryHarvester(ExcelSession session)
    {
        ArgumentNullException.ThrowIfNull(session);

        _summary = ExcelUtils.GetListObject(session.Workbook, SummaryWorksheet, SummaryTable);
        _projectedStatus = ExcelUtils.GetNameConstantText(session.Workbook, ProjectedStatusName);

        Dictionary<string, int> columns = ExcelUtils.GetColumnIndexes(_summary);

        _yearColumn = Column(columns, YearColumn);
        _regimeColumn = Column(columns, RegimeColumn);
        _elementColumns = [.. DataElements.Select(element => Column(columns, element))];

        object?[,] grid = ExcelUtils.ReadTable(_summary);
        (_projectedRows, _years) = FindProjectedRows(grid);

        if (_projectedRows.Length == 0)
        {
            throw new InvalidOperationException("Table " + SummaryTable + " on worksheet " + SummaryWorksheet
                + " has no " + _projectedStatus + " years, so there is nothing to simulate.");
        }

        Log.Logger.Information("PFM_HARVEST_SHAPE: projectionYears=" + Format(_projectedRows.Length) + " first="
            + Format(_years[0]) + " last=" + Format(_years[^1]));
    }

    /// <summary>
    /// Gets the calendar years the harvest covers, in table order.  These are the simulated years, and they are the
    /// same for every simulation of a sweep: shifting the market data a projection is driven from changes what happens
    /// in each year, not which years are projected.
    /// </summary>
    public IReadOnlyList<int> Years => _years;

    /// <summary>
    /// Gets the number of projected years in the workbook, which is how far ahead a simulation looks.
    /// </summary>
    public int ProjectionYears => _years.Length;

    /// <summary>
    /// Harvests the data elements of the simulation the workbook currently holds.
    /// </summary>
    /// <returns>The harvested values, indexed by data element and then by simulated year.</returns>
    /// <exception cref="InvalidOperationException">
    /// The set of projected years changed since the harvester was created, so the results would no longer line up
    /// column for column with the ones already recorded.
    /// </exception>
    public IReadOnlyList<double?[]> Harvest()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        object?[,] grid = ExcelUtils.ReadTable(_summary!);

        (int[] projectedRows, int[] years) = FindProjectedRows(grid);
        if (!years.SequenceEqual(_years))
        {
            throw new InvalidOperationException("The " + _projectedStatus + " years of table " + SummaryTable
                + " changed during the sweep, so the harvested results no longer describe the same years.");
        }

        var harvested = new double?[_elementColumns.Length][];

        for (int element = 0; element < _elementColumns.Length; element++)
        {
            var values = new double?[projectedRows.Length];
            for (int year = 0; year < projectedRows.Length; year++)
            {
                // A cell that holds an error marshals back as its integer error code rather than as a number.  It is
                // recorded as a gap instead of being coerced, so a broken year cannot masquerade as a result.
                values[year] = grid[projectedRows[year], _elementColumns[element]] as double?;
            }

            harvested[element] = values;
        }

        return harvested;
    }

    /// <summary>
    /// Releases the table the harvester holds.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        ExcelUtils.ReleaseComObject(_summary);
        _summary = null;
    }

    /// <summary>
    /// Finds the rows of the summary that report a projected year.
    /// </summary>
    /// <param name="grid">The data body of the summary table.</param>
    /// <returns>The zero based row positions of the projected years, and the years themselves.</returns>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "The grid comes from Excel, which marshals a multi cell range as a rectangular variant array.")]
    private (int[] Rows, int[] Years) FindProjectedRows(object?[,] grid)
    {
        var rows = new List<int>();
        var years = new List<int>();

        for (int row = 0; row < grid.GetLength(0); row++)
        {
            if (!string.Equals(grid[row, _regimeColumn] as string, _projectedStatus, StringComparison.Ordinal))
            {
                continue;
            }

            if (grid[row, _yearColumn] is not double year)
            {
                throw new InvalidOperationException("Row " + Format(row + 1) + " of table " + SummaryTable
                    + " has a non numeric " + YearColumn + " value.");
            }

            rows.Add(row);
            years.Add((int)year);
        }

        return ([.. rows], [.. years]);
    }

    /// <summary>
    /// Resolves the position of a column the harvest needs.
    /// </summary>
    /// <param name="columns">The column positions of the summary table.</param>
    /// <param name="name">The header text of the column.</param>
    /// <returns>The zero based position of the column.</returns>
    private static int Column(Dictionary<string, int> columns, string name)
    {
        ArgumentNullException.ThrowIfNull(columns);

        return columns.TryGetValue(name, out int index)
            ? index
            : throw new InvalidOperationException("Column " + name + " was not found in table " + SummaryTable
                + " on worksheet " + SummaryWorksheet + ".");
    }

    /// <summary>
    /// Formats a count for a log or exception message.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <returns>The formatted value.</returns>
    private static string Format(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }
}

/// <summary>
/// Writes the results of a sweep as CSV, one row per simulation.
/// </summary>
/// <remarks>
/// The layout puts one column per simulated year and repeats that run of columns for each data element, so an element
/// reads across as a contiguous block of years.  The columns describing how the tax provision settled follow the year
/// blocks, and the columns identifying the simulation lead them.
/// <para>
/// The results exist to be read back into Excel, so every value is written in a form that survives the round trip: the
/// invariant culture throughout, no grouping separators, no currency or percent formatting, the shortest representation
/// that reads back as the same double, and TRUE or FALSE for a flag so that Excel imports it as a boolean rather than
/// as text.  A value that was not a number is written as an empty field, which Excel imports as a blank cell.
/// </para>
/// <para>
/// Rows are written and flushed as each simulation finishes rather than buffered to the end.  A sweep runs for hours,
/// and a run that is interrupted at hour three should leave three hours of results behind rather than nothing.
/// </para>
/// </remarks>
public sealed class SimulationResultWriter : IDisposable
{
    private readonly StreamWriter _writer;
    private readonly int _elementCount;
    private readonly int _yearCount;

    private bool _disposed;

    /// <summary>
    /// Initializes a new instance of the <see cref="SimulationResultWriter"/> class and writes the header row.
    /// </summary>
    /// <param name="path">The path of the file to write.</param>
    /// <param name="elements">The data elements each simulation reports, in column order.</param>
    /// <param name="years">The simulated years each data element reports, in column order.</param>
    public SimulationResultWriter(string path, IReadOnlyList<string> elements, IReadOnlyList<int> years)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentNullException.ThrowIfNull(elements);
        ArgumentNullException.ThrowIfNull(years);

        _elementCount = elements.Count;
        _yearCount = years.Count;

        // No byte order mark: the content is entirely ASCII, and a mark only risks being read as data by an importer
        // that does not expect one.  CRLF is what Excel writes and expects on this platform.
        _writer = new StreamWriter(path, append: false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false))
        {
            AutoFlush = true,
            NewLine = "\r\n"
        };

        WriteHeader(elements, years);
    }

    /// <summary>
    /// Writes one simulation's results.
    /// </summary>
    /// <param name="record">The simulation to write.</param>
    public void Write(SimulationRecord record)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(record);

        if (record.Elements.Count != _elementCount)
        {
            throw new ArgumentException("The simulation reported " + Format(record.Elements.Count)
                + " data elements rather than " + Format(_elementCount) + ".", nameof(record));
        }

        var fields = new List<string>((_elementCount * _yearCount) + 6)
        {
            Format(record.Index),
            Format(record.StartYear),
            Format(record.StartSemester)
        };

        foreach (double?[] values in record.Elements)
        {
            if (values.Length != _yearCount)
            {
                throw new ArgumentException("The simulation reported " + Format(values.Length)
                    + " years rather than " + Format(_yearCount) + ".", nameof(record));
            }

            fields.AddRange(values.Select(Format));
        }

        fields.Add(Format(record.TaxProvision.Passes));
        fields.Add(Format(record.TaxProvision.Drift));
        fields.Add(Format(record.TaxProvision.WithinTolerance));

        _writer.WriteLine(string.Join(',', fields));
    }

    /// <summary>
    /// Closes the file.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _writer.Dispose();
    }

    /// <summary>
    /// Writes the header row: the columns identifying the simulation, then one block of years per data element, then
    /// the columns describing how the tax provision settled.
    /// </summary>
    /// <param name="elements">The data elements each simulation reports.</param>
    /// <param name="years">The simulated years each data element reports.</param>
    private void WriteHeader(IReadOnlyList<string> elements, IReadOnlyList<int> years)
    {
        var headers = new List<string>((elements.Count * years.Count) + 6)
        {
            "SimulationIndex",
            "StartYear",
            "StartSemester"
        };

        foreach (string element in elements)
        {
            headers.AddRange(years.Select(year => element + "_" + Format(year)));
        }

        headers.Add("ConvergencePasses");
        headers.Add("FinalTaxDrift");
        headers.Add("DriftWithinTolerance");

        _writer.WriteLine(string.Join(',', headers));
    }

    /// <summary>
    /// Formats a whole number for CSV.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <returns>The formatted value.</returns>
    private static string Format(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Formats a number for CSV.
    /// </summary>
    /// <param name="value">The value to format, or null when the workbook did not report a number.</param>
    /// <returns>The formatted value, or an empty field.</returns>
    /// <remarks>
    /// The default double format has produced the shortest string that reads back as the same value since .NET Core
    /// 3.0, which is exactly what a round trip through CSV needs: no precision is lost, and no digits are written that
    /// only look like precision.
    /// </remarks>
    private static string Format(double? value)
    {
        return value?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
    }

    /// <summary>
    /// Formats a flag for CSV.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <returns>TRUE or FALSE, which Excel imports as a boolean.</returns>
    private static string Format(bool value)
    {
        return value ? "TRUE" : "FALSE";
    }
}
