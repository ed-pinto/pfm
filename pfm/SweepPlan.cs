using System.Globalization;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pfm;

/// <summary>
/// What one sweep drives and how each of its simulations is identified.
/// </summary>
/// <remarks>
/// A back test and a Monte Carlo campaign are the same run with different drivers: both stamp a parameter or two on
/// 10_Parameters, recalculate, settle the materialized series and harvest 56_Summary.  What differs is which
/// parameters a simulation writes, and therefore what identifies it in the results.  Stating that difference here is
/// what lets one loop drive both, which is also what makes the two results files readable the same way.
/// <para>
/// A simulation is addressed by its zero based offset within the whole sweep rather than within the job running it,
/// so the slices of a parallel sweep reassemble by concatenation.
/// </para>
/// </remarks>
public abstract class SweepPlan
{
    /// <summary>
    /// Gets the tag the sweep's log messages carry, ex. BACKTEST in PFM_BACKTEST_COMPLETE.
    /// </summary>
    public abstract string LogName { get; }

    /// <summary>
    /// Gets the columns that say which simulation a results row is, beyond the sweep wide index leading every row.
    /// </summary>
    public abstract IReadOnlyList<string> IdentityColumns { get; }

    /// <summary>
    /// Describes one simulation for a log or console message, ex. simulation 12 (1945 S1).
    /// </summary>
    /// <param name="offset">The zero based index of the simulation within the whole sweep.</param>
    /// <returns>The description.</returns>
    public abstract string Describe(int offset);

    /// <summary>
    /// Writes the parameters that make the workbook hold one simulation.  The caller recalculates afterwards.
    /// </summary>
    /// <param name="workbook">The job's own copy of the workbook.</param>
    /// <param name="offset">The zero based index of the simulation within the whole sweep.</param>
    public abstract void Drive(Excel.Workbook workbook, int offset);

    /// <summary>
    /// Yields the values of <see cref="IdentityColumns"/> for one simulation.
    /// </summary>
    /// <param name="offset">The zero based index of the simulation within the whole sweep.</param>
    /// <returns>The values, in the order the columns are stated in.</returns>
    public abstract IReadOnlyList<int> Identify(int offset);

    /// <summary>
    /// Formats a count for a message.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <returns>The formatted value.</returns>
    protected static string Format(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }
}

/// <summary>
/// A back test: one simulation per period of 21_EconomyHistorical the projection is long enough to finish in.
/// </summary>
/// <remarks>
/// A simulation sets BackTestingStartYear and BackTestingStartSemester, which moves 30_MarketSimulation onto a
/// different run of historical semesters and reprojects the whole model from it.  The starting point advances one
/// semester per simulation, so the sweep asks what this plan would have done had it begun in every period history is
/// long enough to have finished in.
/// </remarks>
public sealed class BackTestSweepPlan : SweepPlan
{
    private const string EconomyWorksheet = "21_EconomyHistorical";
    private const string EconomyTable = "EconomyHistorical";
    private const string EconomyYearColumn = "Year";
    private const string EconomySemesterColumn = "Semester";

    private const string StartYearName = "BackTestingStartYear";
    private const string StartSemesterName = "BackTestingStartSemester";

    private readonly int[] _years;
    private readonly int[] _semesters;

    /// <summary>
    /// Initializes a new instance of the <see cref="BackTestSweepPlan"/> class.
    /// </summary>
    /// <param name="years">The year of each period of 21_EconomyHistorical, in table order.</param>
    /// <param name="semesters">The semester of each period of 21_EconomyHistorical, in table order.</param>
    private BackTestSweepPlan(int[] years, int[] semesters)
    {
        _years = years;
        _semesters = semesters;
    }

    /// <inheritdoc/>
    public override string LogName => "BACKTEST";

    /// <inheritdoc/>
    public override IReadOnlyList<string> IdentityColumns { get; } = ["StartYear", "StartSemester"];

    /// <summary>
    /// Gets the number of historical periods a simulation could be driven from, before the projection horizon is
    /// taken into account.
    /// </summary>
    public int PeriodCount => _years.Length;

    /// <summary>
    /// Reads the periods of the historical economy from the workbook.
    /// </summary>
    /// <param name="workbook">The job's own copy of the workbook.</param>
    /// <returns>The plan.</returns>
    /// <remarks>
    /// The periods are read from the table rather than counted forward from a first year, so a sweep starts where the
    /// data starts and steps the way the data steps.  Extending 21_EconomyHistorical backwards or forwards therefore
    /// changes what the sweep covers without changing this tool.
    /// </remarks>
    public static BackTestSweepPlan Create(Excel.Workbook workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.ListObject economy = ExcelUtils.GetListObject(workbook, EconomyWorksheet, EconomyTable);
        try
        {
            int[] years = ReadWholeNumbers(economy, EconomyYearColumn);
            int[] semesters = ReadWholeNumbers(economy, EconomySemesterColumn);

            if (years.Length != semesters.Length)
            {
                throw new InvalidOperationException("The columns of table " + EconomyTable
                    + " do not have a consistent number of rows.");
            }

            Log.Logger.Information("PFM_BACKTEST_PERIODS: periods=" + Format(years.Length) + " first="
                + Format(years[0]) + " S" + Format(semesters[0]) + " last=" + Format(years[^1]) + " S"
                + Format(semesters[^1]));

            return new BackTestSweepPlan(years, semesters);
        }
        finally
        {
            ExcelUtils.ReleaseComObject(economy);
        }
    }

    /// <inheritdoc/>
    public override string Describe(int offset)
    {
        return "simulation " + Format(offset) + " (" + Format(_years[offset]) + " S" + Format(_semesters[offset])
            + ")";
    }

    /// <inheritdoc/>
    public override void Drive(Excel.Workbook workbook, int offset)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        // Writing the start moves 30_MarketSimulation onto a different run of historical semesters; the recalculation
        // that follows reprojects the whole model from it.
        ExcelUtils.SetNameValue(workbook, StartYearName, (double)_years[offset]);
        ExcelUtils.SetNameValue(workbook, StartSemesterName, (double)_semesters[offset]);
    }

    /// <inheritdoc/>
    public override IReadOnlyList<int> Identify(int offset)
    {
        return [_years[offset], _semesters[offset]];
    }

    /// <summary>
    /// Reads a column of whole numbers from a table.
    /// </summary>
    /// <param name="table">The table that contains the column.</param>
    /// <param name="columnName">The header text of the column.</param>
    /// <returns>The column's values in row order.</returns>
    private static int[] ReadWholeNumbers(Excel.ListObject table, string columnName)
    {
        object?[] values = ExcelUtils.ReadColumn(table, columnName);
        var numbers = new int[values.Length];

        for (int row = 0; row < values.Length; row++)
        {
            if (values[row] is not double value)
            {
                throw new InvalidOperationException("Row " + Format(row + 1) + " of column " + columnName
                    + " in table " + table.Name + " is not a number.");
            }

            numbers[row] = (int)value;
        }

        return numbers;
    }
}

/// <summary>
/// A Monte Carlo campaign: one simulation per seed column of 22_MCSeeds.
/// </summary>
/// <remarks>
/// A simulation sets MCIteration alone, which is the whole of the contract 22_MCSeeds and 30_MarketSimulation state
/// for an external runner.  It selects the column of seeds the block bootstrap draws its block start rows from, so
/// writing it moves the projection onto a different spliced run of real history.  Nothing else about the driver set
/// is stamped per iteration and none of it is random, so an iteration is exactly reproducible.
/// </remarks>
public sealed class MonteCarloSweepPlan : SweepPlan
{
    private const string SeedsWorksheet = "22_MCSeeds";
    private const string SeedsTable = "MCSeeds";
    private const string SeedsYearColumn = "Year";

    private const string IterationName = "MCIteration";

    /// <summary>
    /// Initializes a new instance of the <see cref="MonteCarloSweepPlan"/> class.
    /// </summary>
    /// <param name="iterationCount">The number of iterations the workbook holds seeds for.</param>
    private MonteCarloSweepPlan(int iterationCount)
    {
        IterationCount = iterationCount;
    }

    /// <inheritdoc/>
    public override string LogName => "MONTECARLO";

    /// <inheritdoc/>
    public override IReadOnlyList<string> IdentityColumns { get; } = [IterationName];

    /// <summary>
    /// Gets the number of iterations the sweep runs, which is the number of seed columns 22_MCSeeds holds.  The
    /// iterations themselves are numbered from one.
    /// </summary>
    public int IterationCount { get; }

    /// <summary>
    /// Reads the number of iterations the workbook holds seeds for.
    /// </summary>
    /// <param name="workbook">The job's own copy of the workbook.</param>
    /// <returns>The plan.</returns>
    /// <remarks>
    /// GetMCSeed reads the seed of a year from column MCIteration + 1 of MCSeeds, column one being Year, so the seed
    /// columns are exactly the iterations the workbook can be driven through: one to a thousand as the table stands
    /// today.  Counting them rather than stating that range here keeps a widened seed table a change to the workbook
    /// alone, which is what a back test does with the periods of 21_EconomyHistorical.
    /// </remarks>
    public static MonteCarloSweepPlan Create(Excel.Workbook workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.ListObject seeds = ExcelUtils.GetListObject(workbook, SeedsWorksheet, SeedsTable);
        try
        {
            int yearPosition = ExcelUtils.GetColumnPosition(seeds, SeedsYearColumn);

            // The seed columns are counted as everything after Year, which only describes the table while Year leads
            // it.  GetMCSeed assumes the same thing when it indexes by MCIteration + 1, so a table whose Year column
            // has moved is one this sweep would drive to the wrong seeds rather than one it can count.
            if (yearPosition != 1)
            {
                throw new InvalidOperationException("Column " + SeedsYearColumn + " is at position "
                    + Format(yearPosition) + " of table " + SeedsTable + " rather than the first, so the columns "
                    + "after it are not the seeds of iterations one and up.");
            }

            int iterationCount = ExcelUtils.GetColumnCount(seeds) - 1;

            if (iterationCount < 1)
            {
                throw new InvalidOperationException("Table " + SeedsTable + " on worksheet " + SeedsWorksheet
                    + " holds no seed columns beside " + SeedsYearColumn + ", so there is nothing to simulate.");
            }

            Log.Logger.Information("PFM_MONTECARLO_ITERATIONS: iterations=" + Format(iterationCount));

            return new MonteCarloSweepPlan(iterationCount);
        }
        finally
        {
            ExcelUtils.ReleaseComObject(seeds);
        }
    }

    /// <summary>
    /// Yields the iteration one simulation of the sweep runs.
    /// </summary>
    /// <param name="offset">The zero based index of the simulation within the whole sweep.</param>
    /// <returns>The iteration, numbered from one as the seed columns are.</returns>
    public int Iteration(int offset)
    {
        return offset + 1;
    }

    /// <inheritdoc/>
    public override string Describe(int offset)
    {
        return "simulation " + Format(offset) + " (iteration " + Format(Iteration(offset)) + ")";
    }

    /// <inheritdoc/>
    public override void Drive(Excel.Workbook workbook, int offset)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        // The one cell an iteration stamps.  Everything the path is made of follows from it: GetMCSeed reads this
        // column of 22_MCSeeds, and the MC_ columns of 30_MarketSimulation draw their block start rows with it.
        ExcelUtils.SetNameValue(workbook, IterationName, (double)Iteration(offset));
    }

    /// <inheritdoc/>
    public override IReadOnlyList<int> Identify(int offset)
    {
        return [Iteration(offset)];
    }
}
