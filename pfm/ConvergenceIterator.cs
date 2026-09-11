using System.Globalization;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pfm;

/// <summary>
/// How a run of the convergence iteration ended.
/// </summary>
public enum ConvergenceStatus
{
    /// <summary>
    /// Every materialized series agrees with the live value it stands for: every guard check reported OK.
    /// </summary>
    Converged,

    /// <summary>
    /// At least one series was still moving when the pass limit was reached.
    /// </summary>
    NotConverged,

    /// <summary>
    /// No series would be moved by another pass and a guard check still disagrees, so no further pass can settle it.
    /// </summary>
    Stalled
}

/// <summary>
/// Describes one materialized series the iteration refreshes: a static column that stands in for a live one the
/// workbook cannot link to, and the check that guards it.
/// </summary>
/// <remarks>
/// The workbook materializes a value wherever a live link would be circular at range level, which Excel does not report
/// as an error: it silently freezes the column instead (D009).  Every such column is refreshed the same way, from a
/// live column beside it, over the rows a status column marks, within a tolerance the workbook states, until the check
/// that compares the two reports OK.  Stating that shape once is what lets one loop settle every one of them together,
/// which is what the workbook's own close runbook asks for.
/// </remarks>
public sealed class ConvergenceSeriesDefinition
{
    /// <summary>
    /// Gets the name the series is reported under, ex. TaxProvision.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// Gets the worksheet that hosts the table, ex. 15_TaxPayments.
    /// </summary>
    public required string Worksheet { get; init; }

    /// <summary>
    /// Gets the table that holds both columns, ex. TaxProvision.
    /// </summary>
    public required string Table { get; init; }

    /// <summary>
    /// Gets the column that says which rows the iteration owns, ex. ProvisionSource.
    /// </summary>
    public required string SelectorColumn { get; init; }

    /// <summary>
    /// Gets the defined name holding the text constant a row of <see cref="SelectorColumn"/> must equal for the
    /// iteration to own it, ex. ProvisionModelled.
    /// </summary>
    public required string SelectorName { get; init; }

    /// <summary>
    /// Gets the column holding the live value the materialized column is refreshed from, ex. TotalTaxLive.
    /// </summary>
    public required string LiveColumn { get; init; }

    /// <summary>
    /// Gets the static column the iteration writes, ex. ProvisionAmount.
    /// </summary>
    public required string MaterializedColumn { get; init; }

    /// <summary>
    /// Gets the defined name holding the dollar tolerance the guard check judges the difference against, ex.
    /// TaxProvisionTolerance.
    /// </summary>
    public required string ToleranceName { get; init; }

    /// <summary>
    /// Gets the id of the check on 90_Checks that reports whether the series is refreshed, ex. K39.
    /// </summary>
    public required string CheckId { get; init; }
}

/// <summary>
/// How one materialized series stood when a run ended.
/// </summary>
public sealed class ConvergenceSeriesOutcome
{
    /// <summary>
    /// Gets the name of the series, ex. TaxProvision.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// Gets the id of the check that guards the series, ex. K39.
    /// </summary>
    public required string CheckId { get; init; }

    /// <summary>
    /// Gets the largest absolute difference, in dollars, between the materialized column and the live column across
    /// the rows the iteration owns, as the run left them.
    /// </summary>
    public required double Drift { get; init; }

    /// <summary>
    /// Gets the dollar tolerance the drift was judged against, read from the series' tolerance name.
    /// </summary>
    public required double Tolerance { get; init; }

    /// <summary>
    /// Indicates whether <see cref="Drift"/> is within <see cref="Tolerance"/>.  This is the condition the guard check
    /// tests.
    /// </summary>
    public required bool WithinTolerance { get; init; }

    /// <summary>
    /// Gets the result the guard check reported when the run ended, ex. OK or FAIL.
    /// </summary>
    public required string CheckResult { get; init; }
}

/// <summary>
/// What one run of the convergence iteration produced.  A simulation records this alongside the outcome metrics it
/// harvests, because a year whose materialized values never settled is not comparable with one whose values did.
/// </summary>
public sealed class ConvergenceOutcome
{
    /// <summary>
    /// Gets how the run ended.
    /// </summary>
    public required ConvergenceStatus Status { get; init; }

    /// <summary>
    /// Gets the number of refresh and recalculate passes the run performed.  A workbook that was already settled costs
    /// zero passes.
    /// </summary>
    public required int Passes { get; init; }

    /// <summary>
    /// Gets how each materialized series stood when the run ended, in the order
    /// <see cref="ConvergenceIterator.Definitions"/> states them.
    /// </summary>
    public required IReadOnlyList<ConvergenceSeriesOutcome> Series { get; init; }

    /// <summary>
    /// Gets a value indicating whether every series is within its own tolerance.  A run that converged is within
    /// tolerance by definition; one that did not may still be, when a check disagreed by less than a pass could
    /// resolve.
    /// </summary>
    public bool WithinTolerance => Series.All(series => series.WithinTolerance);

    /// <summary>
    /// Renders the drift of every series for a log or console message, ex. "TaxProvision 0.00, PortfolioState 12.34".
    /// </summary>
    /// <returns>The description.</returns>
    public string DescribeDrift()
    {
        return string.Join(", ", Series.Select(series => series.Name + " "
            + series.Drift.ToString("F2", CultureInfo.InvariantCulture)));
    }

    /// <summary>
    /// Renders the series that a run left disagreeing with their guard checks, ex. "PortfolioState (K53 reports FAIL
    /// with a drift of 12.34 against a tolerance of 1.00)".
    /// </summary>
    /// <returns>The description, or an empty string when every check agreed.</returns>
    public string DescribeFailingSeries()
    {
        return string.Join("; ", Series
            .Where(series => !series.WithinTolerance)
            .Select(series => series.Name + " (" + series.CheckId + " reports " + series.CheckResult
                + " with a drift of " + series.Drift.ToString("F2", CultureInfo.InvariantCulture)
                + " against a tolerance of " + series.Tolerance.ToString("F2", CultureInfo.InvariantCulture) + ")"));
    }
}

/// <summary>
/// Iterates every materialized series of the workbook until each agrees with the live value it stands for.
/// </summary>
/// <remarks>
/// A materialized column feeds the very model that produces the live column it is refreshed from: 15_TaxPayments
/// TaxProvision.ProvisionAmount feeds the tax funding that 55_Tax draws on, and 16_PortfolioState
/// PortfolioState.OpeningPortfolioValue feeds the income throttle that 51_Targets draws on, which 50_Assets and so the
/// opening portfolio itself respond to.  A single pass therefore cannot settle either; each pass copies every live
/// column into the static column beside it for the rows the iteration owns and recalculates, until every guard check
/// on 90_Checks reports OK.
/// <para>
/// The series are settled together in one loop rather than one after another, because they are coupled: refreshing the
/// provision moves the portfolio the throttle reads, and refreshing that portfolio moves the tax the provision holds.
/// Settling one and then the other would leave the first stale again, so a pass writes every stale series and pays for
/// one recalculation rather than one per series.  It is also what the workbook's own close runbook asks for, which
/// re-converges both materialized series together.
/// </para>
/// <para>
/// A run may instead write part of the difference, which is what <see cref="DampedGain"/> is for.  The undamped loop
/// settles only while the feedback from a materialized column back to its live column is weaker than one for one; where
/// it is not, the value oscillates between two values that a pass swaps rather than narrows.  Writing half the
/// difference halves that feedback as the iteration sees it, which turns the oscillation into a convergence.  It also
/// slows down a series that was converging anyway, so it belongs to a retry of a run that failed rather than to every
/// run.
/// </para>
/// <para>
/// The iterator holds the COM references and the workbook constants a run needs, so a sweep that settles hundreds of
/// simulations resolves them once rather than once per simulation.  It is therefore shared by the iterate command,
/// which runs it exactly once, and by the simulation commands, which run it once per simulated period.
/// </para>
/// </remarks>
public sealed class ConvergenceIterator : IDisposable
{
    private const string ChecksWorksheet = "90_Checks";
    private const string ChecksTable = "Checks";
    private const string CheckResultOk = "OK";

    /// <summary>
    /// The number of passes after which the workbook is treated as failing to converge.
    /// </summary>
    public const int DefaultMaxPasses = 100;

    /// <summary>
    /// The gain of an undamped run: each pass writes the whole difference, so the materialized value becomes the live
    /// one.
    /// </summary>
    public const double FullGain = 1.0;

    /// <summary>
    /// The gain of a damped run: each pass writes half the difference.  Halving the step halves the effective loop
    /// gain, which settles the period two oscillation an undamped run cannot leave and widens the range of feedback
    /// the iteration tolerates at all, at the cost of roughly twice the passes on a series that would have converged
    /// anyway.  That cost is why it is what a retry uses rather than what every run uses.
    /// </summary>
    public const double DampedGain = 0.5;

    /// <summary>
    /// The materialized series the iteration refreshes, in the order a pass reads and writes them.
    /// </summary>
    /// <remarks>
    /// This is the list to extend when the workbook materializes another series.  A new entry needs no code beyond
    /// itself: the loop, the outcome and the results file are all stated in terms of this list.
    /// </remarks>
    public static readonly IReadOnlyList<ConvergenceSeriesDefinition> Definitions =
    [
        new ConvergenceSeriesDefinition
        {
            Name = "TaxProvision",
            Worksheet = "15_TaxPayments",
            Table = "TaxProvision",
            SelectorColumn = "ProvisionSource",
            SelectorName = "ProvisionModelled",
            LiveColumn = "TotalTaxLive",
            MaterializedColumn = "ProvisionAmount",
            ToleranceName = "TaxProvisionTolerance",
            CheckId = "K39"
        },
        new ConvergenceSeriesDefinition
        {
            Name = "PortfolioState",
            Worksheet = "16_PortfolioState",
            Table = "PortfolioState",
            SelectorColumn = "Regime",
            SelectorName = "StatusProjected",
            LiveColumn = "OpeningPortfolioLive",
            MaterializedColumn = "OpeningPortfolioValue",
            ToleranceName = "PortfolioStateTolerance",
            CheckId = "K53"
        }
    ];

    private readonly ExcelSession _session;
    private readonly DiagnosticTimer _timer;
    private readonly int _maxPasses;
    private readonly SeriesBinding[] _series;

    private Excel.ListObject? _checks;
    private bool _disposed;

    /// <summary>
    /// Initializes a new instance of the <see cref="ConvergenceIterator"/> class, resolving the tables and the
    /// workbook constants that every run needs.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="maxPasses">The number of passes after which a run is treated as failing to converge.</param>
    public ConvergenceIterator(ExcelSession session, DiagnosticTimer timer, int maxPasses = DefaultMaxPasses)
    {
        ArgumentNullException.ThrowIfNull(session);
        ArgumentNullException.ThrowIfNull(timer);
        ArgumentOutOfRangeException.ThrowIfNegative(maxPasses);

        _session = session;
        _timer = timer;
        _maxPasses = maxPasses;

        _checks = ExcelUtils.GetListObject(session.Workbook, ChecksWorksheet, ChecksTable);
        _series = [.. Definitions.Select(definition => new SeriesBinding(session.Workbook, definition))];
    }

    /// <summary>
    /// Gets the name of every materialized series the iteration refreshes, in the order an outcome reports them.
    /// </summary>
    public static IReadOnlyList<string> SeriesNames { get; } = [.. Definitions.Select(definition => definition.Name)];

    /// <summary>
    /// Runs the refresh and recalculate loop until every guard check reports OK or the pass limit is reached.
    /// </summary>
    /// <param name="label">
    /// What the measurements of this run are reported under, ex. the simulation it belongs to.  Null names no run,
    /// which suits the iterate command because it only ever performs one.
    /// </param>
    /// <param name="gain">
    /// The fraction of the difference between a live column and its materialized column that a pass writes, ex.
    /// <see cref="FullGain"/> or <see cref="DampedGain"/>.
    /// </param>
    /// <returns>How the run ended.</returns>
    /// <remarks>
    /// The caller is responsible for settling the workbook before the first run.  A check cell left dirty by an earlier
    /// edit would otherwise report FAIL for a series that is already in agreement.
    /// </remarks>
    public ConvergenceOutcome Run(string? label = null, double gain = FullGain)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(gain);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(gain, FullGain);

        int passes = 0;

        while (true)
        {
            // The columns are read before the checks rather than only when a refresh is due, so that the drift the
            // checks just judged can be reported whatever the outcome.
            SeriesPass[] measured = [.. _series.Select(series => Measure(series, gain, label, passes + 1))];
            bool converged = true;
            int stale = 0;

            foreach (SeriesPass pass in measured)
            {
                pass.CheckResult = ReadCheckResult(pass.Series, label, passes + 1);
                converged &= string.Equals(pass.CheckResult, CheckResultOk, StringComparison.OrdinalIgnoreCase);
                stale += pass.Stale;
            }

            if (converged)
            {
                return Outcome(ConvergenceStatus.Converged, passes, measured);
            }

            if (passes >= _maxPasses)
            {
                return Outcome(ConvergenceStatus.NotConverged, passes, measured);
            }

            if (stale == 0)
            {
                // No owned row of any series would be moved by another pass, so no further pass can move a check
                // either.  At any gain above zero that means every owned row already holds its live value.
                // Returning before the write is what leaves the workbook settled on every path out of this loop, which
                // a caller about to harvest its values depends on.
                return Outcome(ConvergenceStatus.Stalled, passes, measured);
            }

            foreach (SeriesPass pass in measured.Where(pass => pass.Stale > 0))
            {
                WriteMaterialized(pass, label, passes + 1);
            }

            passes++;

            using (IDisposable scope = _timer.Measure(Phase("calculate", label, passes)))
            {
                _session.Calculate();
            }

            // The drift is the one measured entering this pass, so a run of these lines is the trajectory the
            // iteration took.  It is what distinguishes a series that is converging slowly from one that is
            // oscillating or diverging, which the final drift alone cannot say.  Every series is reported on the one
            // line, because what a coupled loop does to one of them is only readable beside what it did to the others.
            Log.Logger.Information("PFM_ITERATE_PASS: " + Describe(label) + "pass=" + Format(passes) + " gain="
                + Format(gain) + " " + string.Join(" ", measured.Select(pass => pass.Series.Definition.Name
                    + "(refreshed=" + Format(pass.Stale) + " drift=" + Format(pass.Drift) + ")")));
        }
    }

    /// <summary>
    /// Releases the tables the iterator holds.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        foreach (SeriesBinding series in _series)
        {
            series.Dispose();
        }

        ExcelUtils.ReleaseComObject(_checks);
        _checks = null;
    }

    /// <summary>
    /// Builds the outcome of a run and logs it.
    /// </summary>
    /// <param name="status">How the run ended.</param>
    /// <param name="passes">The number of passes the run performed.</param>
    /// <param name="measured">What the last pass measured of each series.</param>
    /// <returns>The outcome.</returns>
    private static ConvergenceOutcome Outcome(ConvergenceStatus status, int passes, SeriesPass[] measured)
    {
        var outcome = new ConvergenceOutcome
        {
            Status = status,
            Passes = passes,
            Series =
            [
                .. measured.Select(pass => new ConvergenceSeriesOutcome
                {
                    Name = pass.Series.Definition.Name,
                    CheckId = pass.Series.Definition.CheckId,
                    Drift = pass.Drift,
                    Tolerance = pass.Series.Tolerance,
                    WithinTolerance = pass.Drift <= pass.Series.Tolerance,
                    CheckResult = pass.CheckResult
                })
            ]
        };

        Log.Logger.Information("PFM_ITERATE_OUTCOME: status=" + status + " passes=" + Format(passes) + " "
            + string.Join(" ", outcome.Series.Select(series => series.Name + "(drift=" + Format(series.Drift) + " "
                + series.CheckId + "=" + series.CheckResult + ")")));

        return outcome;
    }

    /// <summary>
    /// Reads one series and measures how far its materialized column is from its live column.
    /// </summary>
    /// <param name="series">The series to measure.</param>
    /// <param name="gain">The fraction of the difference a pass writes.</param>
    /// <param name="label">What the measurement is reported under.</param>
    /// <param name="pass">The one based number of the pass being run.</param>
    /// <returns>What the pass measured, and what it would write.</returns>
    private SeriesPass Measure(SeriesBinding series, double gain, string? label, int pass)
    {
        (object?[] selectors, object?[] live, object?[] materialized) = ReadSeriesColumns(series, label, pass);
        (double drift, int stale, object?[] targets) = Evaluate(series, selectors, live, materialized, gain);

        return new SeriesPass(series, selectors, targets, drift, stale);
    }

    /// <summary>
    /// Measures how far a materialized column is from the live column it stands for, across the rows the iteration
    /// owns.
    /// </summary>
    /// <param name="series">The series being measured.</param>
    /// <param name="selectors">The selector column, ex. ProvisionSource.</param>
    /// <param name="live">The live column, ex. TotalTaxLive.</param>
    /// <param name="materialized">The static column, ex. ProvisionAmount.</param>
    /// <param name="gain">The fraction of the difference a pass writes.</param>
    /// <returns>
    /// The largest absolute difference in dollars, the number of rows a pass would move, and the value a pass would
    /// write for each row.  The difference is zero, no row is stale and no value is a target when the iteration owns
    /// no row.
    /// </returns>
    /// <remarks>
    /// The difference is the quantity the guard check compares against the series' tolerance.  It is measured here as
    /// well so that a simulation can record how far from agreement its values settled, which a pass or fail alone does
    /// not say.  The count is what tells a pass whether there is anything left to write: a check that still fails when
    /// no owned row would move cannot be moved by another pass.
    /// <para>
    /// The target is what makes that count hold at any gain.  A damped pass never lands exactly on the live value, so
    /// counting the rows that do not already hold it would count every row forever and no damped run could ever report
    /// Stalled.  Comparing the value the pass would write instead says the same thing at full gain and the true thing
    /// at any other: at a gain above zero a target equals its current value only when that value is already the live
    /// one, or when the step has become too small for it to represent, which is a row no further pass can move either.
    /// </para>
    /// <para>
    /// Rows the selector does not mark, ex. a ProvisionExogenous tax year or a Closed semester, are excluded by design:
    /// their values are authored or historical rather than modelled, and the guard checks do not compare them either.
    /// </para>
    /// </remarks>
    private static (double Drift, int Stale, object?[] Targets) Evaluate(SeriesBinding series, object?[] selectors,
        object?[] live, object?[] materialized, double gain)
    {
        double drift = 0;
        int stale = 0;
        object?[] targets = new object?[selectors.Length];

        for (int row = 0; row < selectors.Length; row++)
        {
            if (!series.Owns(selectors, row))
            {
                continue;
            }

            if (live[row] is not double liveValue)
            {
                throw new InvalidOperationException("Row " + Format(row + 1) + " of table "
                    + series.Definition.Table + " has a non numeric " + series.Definition.LiveColumn + " value.");
            }

            // An empty materialized cell is a drift of the whole live value, not a row to skip.
            double value = materialized[row] as double? ?? 0;
            targets[row] = value + (gain * (liveValue - value));

            if (!Equals(materialized[row], targets[row]))
            {
                stale++;
            }

            drift = Math.Max(drift, Math.Abs(liveValue - value));
        }

        return (drift, stale, targets);
    }

    /// <summary>
    /// Writes the value a pass moves the materialized column to for every row the iteration owns.
    /// </summary>
    /// <param name="pass">What the pass measured of the series, which is what it writes.</param>
    /// <param name="label">What the measurement is reported under.</param>
    /// <param name="number">The one based number of the pass being run.</param>
    /// <remarks>
    /// Only runs of consecutive owned rows are written, so a row the iteration does not own is never covered by a
    /// write.  That is also what keeps the targets of those rows, which are never measured, out of every write.
    /// </remarks>
    private void WriteMaterialized(SeriesPass pass, string? label, int number)
    {
        SeriesBinding series = pass.Series;

        using IDisposable scope = _timer.Measure(Phase("write " + series.Definition.MaterializedColumn, label,
            number));

        int runStart = -1;

        // One extra step past the end so that a run reaching the last row is written.
        for (int row = 0; row <= pass.Selectors.Length; row++)
        {
            if (series.Owns(pass.Selectors, row))
            {
                if (runStart < 0)
                {
                    runStart = row;
                }
            }
            else if (runStart >= 0)
            {
                ExcelUtils.WriteColumnBlock(series.Table, series.Definition.MaterializedColumn, runStart,
                    pass.Targets[runStart..row]);
                runStart = -1;
            }
        }
    }

    /// <summary>
    /// Reads the columns of a series that a pass needs, measuring how long it takes.
    /// </summary>
    /// <param name="series">The series to read.</param>
    /// <param name="label">What the measurement is reported under.</param>
    /// <param name="pass">The one based number of the pass being run.</param>
    /// <returns>The selector, live and materialized columns, in row order.</returns>
    private (object?[] Selectors, object?[] Live, object?[] Materialized) ReadSeriesColumns(SeriesBinding series,
        string? label, int pass)
    {
        using IDisposable scope = _timer.Measure(Phase("read " + series.Definition.Table + " columns", label, pass));

        object?[] selectors = ExcelUtils.ReadColumn(series.Table, series.Definition.SelectorColumn);
        object?[] live = ExcelUtils.ReadColumn(series.Table, series.Definition.LiveColumn);
        object?[] materialized = ExcelUtils.ReadColumn(series.Table, series.Definition.MaterializedColumn);

        if (selectors.Length != live.Length || selectors.Length != materialized.Length)
        {
            throw new InvalidOperationException("The columns of table " + series.Definition.Table
                + " do not have a consistent number of rows.");
        }

        return (selectors, live, materialized);
    }

    /// <summary>
    /// Reads the result of a series' guard check from the 90_Checks worksheet.
    /// </summary>
    /// <param name="series">The series whose check to read.</param>
    /// <param name="label">What the measurement is reported under.</param>
    /// <param name="pass">The one based number of the pass being run.</param>
    /// <returns>The check result, ex. OK or FAIL.</returns>
    /// <remarks>
    /// The check results occupy a single row table whose headers are the check ids, so a check is located by its id
    /// rather than by its address.  K39 currently resolves to 90_Checks!AM5 and K53 to 90_Checks!BA5.
    /// </remarks>
    private string ReadCheckResult(SeriesBinding series, string? label, int pass)
    {
        string checkId = series.Definition.CheckId;

        using IDisposable scope = _timer.Measure(Phase("read check " + checkId, label, pass));

        object?[] values = ExcelUtils.ReadColumn(_checks!, checkId);

        if (values.Length != 1)
        {
            throw new InvalidOperationException("Check " + checkId + " on worksheet " + ChecksWorksheet
                + " resolved to " + Format(values.Length) + " rows rather than one.");
        }

        return values[0] as string ?? string.Empty;
    }

    /// <summary>
    /// Names a measurement that belongs to one pass of one run.
    /// </summary>
    /// <param name="phase">The name of the phase.</param>
    /// <param name="label">The run the pass belongs to, or null when there is only one run.</param>
    /// <param name="pass">The one based number of the pass.</param>
    /// <returns>The phase name qualified by the run and the pass number.</returns>
    private static string Phase(string phase, string? label, int pass)
    {
        return phase + " (" + Describe(label) + "pass " + Format(pass) + ")";
    }

    /// <summary>
    /// Renders a run label as a prefix for a message, or nothing when there is no label.
    /// </summary>
    /// <param name="label">The run label.</param>
    /// <returns>The label followed by a space, or an empty string.</returns>
    private static string Describe(string? label)
    {
        return string.IsNullOrEmpty(label) ? string.Empty : label + " ";
    }

    /// <summary>
    /// Formats a count for a log or console message.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <returns>The formatted value.</returns>
    private static string Format(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Formats a dollar amount for a log or console message.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <returns>The formatted value.</returns>
    private static string Format(double value)
    {
        return value.ToString("F2", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// One materialized series as the iterator holds it: its definition, the table it lives in, and the workbook
    /// constants that decide which of its rows the iteration owns and how close is close enough.
    /// </summary>
    /// <remarks>
    /// These are resolved once per iterator rather than once per run, because a sweep runs the iteration hundreds of
    /// times against one workbook and none of them can change.
    /// </remarks>
    private sealed class SeriesBinding : IDisposable
    {
        private Excel.ListObject? _table;
        private bool _disposed;

        /// <summary>
        /// Initializes a new instance of the <see cref="SeriesBinding"/> class.
        /// </summary>
        /// <param name="workbook">The workbook that holds the series.</param>
        /// <param name="definition">The series to bind.</param>
        public SeriesBinding(Excel.Workbook workbook, ConvergenceSeriesDefinition definition)
        {
            Definition = definition;

            _table = ExcelUtils.GetListObject(workbook, definition.Worksheet, definition.Table);
            Selector = ExcelUtils.GetNameConstantText(workbook, definition.SelectorName);
            Tolerance = ExcelUtils.GetNameNumber(workbook, definition.ToleranceName);
        }

        /// <summary>
        /// Gets what the series is.
        /// </summary>
        public ConvergenceSeriesDefinition Definition { get; }

        /// <summary>
        /// Gets the text a selector cell must equal for the iteration to own its row, ex. ProvisionModelled.
        /// </summary>
        public string Selector { get; }

        /// <summary>
        /// Gets the dollar tolerance within which the materialized column is treated as agreeing with the live one.
        /// </summary>
        public double Tolerance { get; }

        /// <summary>
        /// Gets the table that holds the series.
        /// </summary>
        public Excel.ListObject Table => _table!;

        /// <summary>
        /// Determines whether the iteration owns a row of the series.
        /// </summary>
        /// <param name="selectors">The selector column.</param>
        /// <param name="row">
        /// The zero based row to test.  A row past the end is not owned, which closes an open run.
        /// </param>
        /// <returns>True when the iteration owns the row.</returns>
        public bool Owns(object?[] selectors, int row)
        {
            return row < selectors.Length
                && string.Equals(selectors[row] as string, Selector, StringComparison.Ordinal);
        }

        /// <summary>
        /// Releases the table the binding holds.
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            ExcelUtils.ReleaseComObject(_table);
            _table = null;
        }
    }

    /// <summary>
    /// What one pass measured of one series, and what it would write for it.
    /// </summary>
    /// <param name="series">The series the pass read.</param>
    /// <param name="selectors">The selector column as the pass read it.</param>
    /// <param name="targets">The value the pass would write for each row.</param>
    /// <param name="drift">The largest absolute difference the pass measured.</param>
    /// <param name="stale">The number of rows the pass would move.</param>
    private sealed class SeriesPass(SeriesBinding series, object?[] selectors, object?[] targets, double drift,
        int stale)
    {
        /// <summary>
        /// Gets the series the pass read.
        /// </summary>
        public SeriesBinding Series { get; } = series;

        /// <summary>
        /// Gets the selector column as the pass read it.
        /// </summary>
        public object?[] Selectors { get; } = selectors;

        /// <summary>
        /// Gets the value the pass would write for each row.
        /// </summary>
        public object?[] Targets { get; } = targets;

        /// <summary>
        /// Gets the largest absolute difference the pass measured, in dollars.
        /// </summary>
        public double Drift { get; } = drift;

        /// <summary>
        /// Gets the number of rows the pass would move.
        /// </summary>
        public int Stale { get; } = stale;

        /// <summary>
        /// Gets or sets the result the series' guard check reported for this pass.
        /// </summary>
        public string CheckResult { get; set; } = string.Empty;
    }
}
