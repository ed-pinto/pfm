using System.Globalization;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pfm;

/// <summary>
/// How a run of the tax provision iteration ended.
/// </summary>
public enum TaxProvisionStatus
{
    /// <summary>
    /// The provision agrees with the modelled total tax: check K39 reported OK.
    /// </summary>
    Converged,

    /// <summary>
    /// The provision was still moving when the pass limit was reached.
    /// </summary>
    NotConverged,

    /// <summary>
    /// The provision stopped moving without the check agreeing, so no further pass can settle it.
    /// </summary>
    Stalled
}

/// <summary>
/// What one run of the tax provision iteration produced.  A simulation records this alongside the outcome metrics it
/// harvests, because a year whose tax never settled is not comparable with one whose tax did.
/// </summary>
public sealed class TaxProvisionOutcome
{
    /// <summary>
    /// Gets how the run ended.
    /// </summary>
    public required TaxProvisionStatus Status { get; init; }

    /// <summary>
    /// Gets the number of refresh and recalculate passes the run performed.  A provision that was already settled costs
    /// zero passes.
    /// </summary>
    public required int Passes { get; init; }

    /// <summary>
    /// Gets the largest absolute difference, in dollars, between ProvisionAmount and TotalTaxLive across the modelled
    /// tax years as the run left them.
    /// </summary>
    public required double Drift { get; init; }

    /// <summary>
    /// Gets the dollar tolerance the drift was judged against, read from TaxProvisionTolerance.
    /// </summary>
    public required double Tolerance { get; init; }

    /// <summary>
    /// Indicates whether <see cref="Drift"/> is within <see cref="Tolerance"/>.  This is the condition check K39 tests.
    /// </summary>
    public required bool WithinTolerance { get; init; }

    /// <summary>
    /// Gets the result check K39 reported when the run ended, ex. OK or FAIL.
    /// </summary>
    public required string CheckResult { get; init; }
}

/// <summary>
/// Iterates 15_TaxPayments.TaxProvision until it agrees with the modelled total tax on 55_Tax.
/// </summary>
/// <remarks>
/// TaxProvision.ProvisionAmount feeds the tax funding that 55_Tax draws on, so changing it changes the very
/// TaxAnnual.TotalTax it is meant to match.  A single pass therefore cannot settle the two; each pass copies
/// TaxProvision.TotalTaxLive (the modelled total tax looked up by year) into TaxProvision.ProvisionAmount for the
/// ProvisionModelled years and recalculates, until check K39 on 90_Checks reports OK.
/// <para>
/// A run may instead write part of that difference, which is what <see cref="DampedGain"/> is for.  The undamped loop
/// settles only while the feedback from ProvisionAmount back to TotalTaxLive is weaker than one for one; where it is
/// not, the provision oscillates between two values that a pass swaps rather than narrows.  Writing half the
/// difference halves that feedback as the iteration sees it, which turns the oscillation into a convergence.  It also
/// slows down a provision that was converging anyway, so it belongs to a retry of a run that failed rather than to
/// every run.
/// </para>
/// <para>
/// The iterator holds the COM references and the workbook constants a run needs, so a sweep that settles the tax of
/// hundreds of simulations resolves them once rather than once per simulation.  It is therefore shared by the iterate
/// command, which runs it exactly once, and by the simulation commands, which run it once per simulated period.
/// </para>
/// </remarks>
public sealed class TaxProvisionIterator : IDisposable
{
    private const string TaxPaymentsWorksheet = "15_TaxPayments";
    private const string TaxProvisionTable = "TaxProvision";
    private const string TotalTaxLiveColumn = "TotalTaxLive";
    private const string ProvisionSourceColumn = "ProvisionSource";
    private const string ProvisionAmountColumn = "ProvisionAmount";
    private const string ProvisionModelledName = "ProvisionModelled";
    private const string TaxProvisionToleranceName = "TaxProvisionTolerance";

    private const string ChecksWorksheet = "90_Checks";
    private const string ChecksTable = "Checks";
    private const string CheckResultOk = "OK";

    /// <summary>
    /// The number of passes after which the tax provision is treated as failing to converge.
    /// </summary>
    public const int DefaultMaxPasses = 100;

    /// <summary>
    /// The gain of an undamped run: each pass writes the whole difference, so ProvisionAmount becomes TotalTaxLive.
    /// </summary>
    public const double FullGain = 1.0;

    /// <summary>
    /// The gain of a damped run: each pass writes half the difference.  Halving the step halves the effective loop
    /// gain, which settles the period two oscillation an undamped run cannot leave and widens the range of feedback
    /// the iteration tolerates at all, at the cost of roughly twice the passes on a provision that would have
    /// converged anyway.  That cost is why it is what a retry uses rather than what every run uses.
    /// </summary>
    public const double DampedGain = 0.5;

    /// <summary>
    /// The id of the check that reports whether the provision agrees with the modelled total tax.
    /// </summary>
    public const string TaxProvisionCheckId = "K39";

    private readonly ExcelSession _session;
    private readonly DiagnosticTimer _timer;
    private readonly int _maxPasses;
    private readonly string _modelledSource;
    private readonly double _tolerance;

    private Excel.ListObject? _provisions;
    private Excel.ListObject? _checks;
    private bool _disposed;

    /// <summary>
    /// Initializes a new instance of the <see cref="TaxProvisionIterator"/> class, resolving the tables and the
    /// workbook constants that every run needs.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="maxPasses">The number of passes after which a run is treated as failing to converge.</param>
    public TaxProvisionIterator(ExcelSession session, DiagnosticTimer timer, int maxPasses = DefaultMaxPasses)
    {
        ArgumentNullException.ThrowIfNull(session);
        ArgumentNullException.ThrowIfNull(timer);
        ArgumentOutOfRangeException.ThrowIfNegative(maxPasses);

        _session = session;
        _timer = timer;
        _maxPasses = maxPasses;

        _provisions = ExcelUtils.GetListObject(session.Workbook, TaxPaymentsWorksheet, TaxProvisionTable);
        _checks = ExcelUtils.GetListObject(session.Workbook, ChecksWorksheet, ChecksTable);
        _modelledSource = ExcelUtils.GetNameConstantText(session.Workbook, ProvisionModelledName);
        _tolerance = ExcelUtils.GetNameNumber(session.Workbook, TaxProvisionToleranceName);
    }

    /// <summary>
    /// Gets the dollar tolerance within which the provision is treated as agreeing with the modelled total tax.
    /// </summary>
    public double Tolerance => _tolerance;

    /// <summary>
    /// Runs the refresh and recalculate loop until check K39 reports OK or the pass limit is reached.
    /// </summary>
    /// <param name="label">
    /// What the measurements of this run are reported under, ex. the simulation it belongs to.  Null names no run,
    /// which suits the iterate command because it only ever performs one.
    /// </param>
    /// <param name="gain">
    /// The fraction of the difference between TotalTaxLive and ProvisionAmount that a pass writes, ex.
    /// <see cref="FullGain"/> or <see cref="DampedGain"/>.
    /// </param>
    /// <returns>How the run ended.</returns>
    /// <remarks>
    /// The caller is responsible for settling the workbook before the first run.  A check cell left dirty by an earlier
    /// edit would otherwise report FAIL for a provision that is already in agreement.
    /// </remarks>
    public TaxProvisionOutcome Run(string? label = null, double gain = FullGain)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(gain);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(gain, FullGain);

        int passes = 0;

        while (true)
        {
            // The columns are read before the check rather than only when a refresh is due, so that the drift the
            // check just judged can be reported whatever the outcome.
            (object?[] sources, object?[] liveTax, object?[] amounts) = ReadProvisionColumns(label, passes + 1);
            (double drift, int stale, object?[] targets) = EvaluateProvision(sources, liveTax, amounts, gain);
            string result = ReadCheckResult(label, passes + 1);

            if (string.Equals(result, CheckResultOk, StringComparison.OrdinalIgnoreCase))
            {
                return Outcome(TaxProvisionStatus.Converged, passes, drift, result);
            }

            if (passes >= _maxPasses)
            {
                return Outcome(TaxProvisionStatus.NotConverged, passes, drift, result);
            }

            if (stale == 0)
            {
                // No modelled year would be moved by another pass, so no further pass can move the check either.  At
                // any gain above zero that means every modelled year already holds its TotalTaxLive value.
                // Returning before the write is what leaves the workbook settled on every path out of this loop, which
                // a caller about to harvest its values depends on.
                return Outcome(TaxProvisionStatus.Stalled, passes, drift, result);
            }

            WriteProvisionAmounts(sources, targets, label, passes + 1);
            passes++;

            using (var _measure = _timer.Measure(Phase("calculate", label, passes)))
                _session.Calculate();

            // The drift is the one measured entering this pass, so a run of these lines is the trajectory the
            // iteration took.  It is what distinguishes a provision that is converging slowly from one that is
            // oscillating or diverging, which the final drift alone cannot say.
            Log.Logger.Information("PFM_ITERATE_PASS: " + Describe(label) + "pass=" + Format(passes) + " gain="
                + Format(gain) + " refreshed=" + Format(stale) + " drift=" + Format(drift));
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

        ExcelUtils.ReleaseComObject(_checks);
        ExcelUtils.ReleaseComObject(_provisions);
        _checks = null;
        _provisions = null;
    }

    /// <summary>
    /// Builds the outcome of a run and logs it.
    /// </summary>
    /// <param name="status">How the run ended.</param>
    /// <param name="passes">The number of passes the run performed.</param>
    /// <param name="drift">The largest absolute provision drift the run left behind.</param>
    /// <param name="result">The result check K39 reported.</param>
    /// <returns>The outcome.</returns>
    private TaxProvisionOutcome Outcome(TaxProvisionStatus status, int passes, double drift, string result)
    {
        var outcome = new TaxProvisionOutcome
        {
            Status = status,
            Passes = passes,
            Drift = drift,
            Tolerance = _tolerance,
            WithinTolerance = drift <= _tolerance,
            CheckResult = result
        };

        Log.Logger.Information("PFM_ITERATE_OUTCOME: status=" + status + " passes=" + Format(passes) + " drift="
            + Format(drift) + " check=" + result);

        return outcome;
    }

    /// <summary>
    /// Measures how far the provision is from the modelled total tax across the modelled tax years.
    /// </summary>
    /// <param name="sources">The ProvisionSource column.</param>
    /// <param name="liveTax">The TotalTaxLive column.</param>
    /// <param name="amounts">The ProvisionAmount column.</param>
    /// <param name="gain">The fraction of the difference a pass writes.</param>
    /// <returns>
    /// The largest absolute difference in dollars, the number of years a pass would move, and the value a pass would
    /// write for each row.  The difference is zero, no year is stale and no value is a target when no year is modelled.
    /// </returns>
    /// <remarks>
    /// The difference is the quantity check K39 compares against TaxProvisionTolerance.  It is measured here as well so
    /// that a simulation can record how far from agreement its tax settled, which a pass or fail alone does not say.
    /// The count is what tells a pass whether there is anything left to write: a check that still fails when no
    /// modelled year would move cannot be moved by another pass.
    /// <para>
    /// The target is what makes that count hold at any gain.  A damped pass never lands exactly on TotalTaxLive, so
    /// counting the years that do not already hold their live value would count every year forever and no damped run
    /// could ever report Stalled.  Comparing the value the pass would write instead says the same thing at full gain
    /// and the true thing at any other: at a gain above zero a target equals its amount only when the amount is
    /// already the live value, or when the step has become too small for the amount to represent, which is a year no
    /// further pass can move either.
    /// </para>
    /// <para>
    /// Years with any other provision source, ex. ProvisionExogenous, are excluded by design: their amounts are
    /// authored rather than modelled, and K39 does not compare them either.
    /// </para>
    /// </remarks>
    private (double Drift, int Stale, object?[] Targets) EvaluateProvision(object?[] sources, object?[] liveTax,
        object?[] amounts, double gain)
    {
        double drift = 0;
        int stale = 0;
        object?[] targets = new object?[sources.Length];

        for (int row = 0; row < sources.Length; row++)
        {
            if (!IsModelled(sources, row))
            {
                continue;
            }

            if (liveTax[row] is not double live)
            {
                throw new InvalidOperationException("Row " + Format(row + 1) + " of table " + TaxProvisionTable
                    + " has a non numeric " + TotalTaxLiveColumn + " value.");
            }

            // An empty ProvisionAmount is a drift of the whole modelled tax, not a row to skip.
            double amount = amounts[row] as double? ?? 0;
            targets[row] = amount + (gain * (live - amount));

            if (!Equals(amounts[row], targets[row]))
            {
                stale++;
            }

            drift = Math.Max(drift, Math.Abs(live - amount));
        }

        return (drift, stale, targets);
    }

    /// <summary>
    /// Writes the value a pass moves ProvisionAmount to for every year whose ProvisionSource is ProvisionModelled.
    /// </summary>
    /// <param name="sources">The ProvisionSource column.</param>
    /// <param name="targets">The value to write for each row, as measured by <see cref="EvaluateProvision"/>.</param>
    /// <param name="label">What the measurement is reported under.</param>
    /// <param name="pass">The one based number of the pass being run.</param>
    /// <remarks>
    /// Only runs of consecutive modelled rows are written, so a non modelled row is never covered by a write.  That is
    /// also what keeps the targets of the non modelled rows, which are never measured, out of every write.
    /// </remarks>
    private void WriteProvisionAmounts(object?[] sources, object?[] targets, string? label, int pass)
    {
        using IDisposable scope = _timer.Measure(Phase("write provision amounts", label, pass));

        int runStart = -1;

        // One extra step past the end so that a run reaching the last row is written.
        for (int row = 0; row <= sources.Length; row++)
        {
            if (IsModelled(sources, row))
            {
                if (runStart < 0)
                {
                    runStart = row;
                }
            }
            else if (runStart >= 0)
            {
                ExcelUtils.WriteColumnBlock(_provisions!, ProvisionAmountColumn, runStart, targets[runStart..row]);
                runStart = -1;
            }
        }
    }

    /// <summary>
    /// Determines whether a row of the TaxProvision table is a modelled tax year.
    /// </summary>
    /// <param name="sources">The ProvisionSource column.</param>
    /// <param name="row">The zero based row to test.  A row past the end is not modelled, which closes an open run.</param>
    /// <returns>True when the row's provision is modelled.</returns>
    private bool IsModelled(object?[] sources, int row)
    {
        return row < sources.Length
            && string.Equals(sources[row] as string, _modelledSource, StringComparison.Ordinal);
    }

    /// <summary>
    /// Reads the columns of the TaxProvision table that a pass needs, measuring how long it takes.
    /// </summary>
    /// <param name="label">What the measurement is reported under.</param>
    /// <param name="pass">The one based number of the pass being run.</param>
    /// <returns>The ProvisionSource, TotalTaxLive and ProvisionAmount columns, in row order.</returns>
    private (object?[] Sources, object?[] LiveTax, object?[] Amounts) ReadProvisionColumns(string? label, int pass)
    {
        using IDisposable scope = _timer.Measure(Phase("read provision columns", label, pass));

        object?[] sources = ExcelUtils.ReadColumn(_provisions!, ProvisionSourceColumn);
        object?[] liveTax = ExcelUtils.ReadColumn(_provisions!, TotalTaxLiveColumn);
        object?[] amounts = ExcelUtils.ReadColumn(_provisions!, ProvisionAmountColumn);

        if (sources.Length != liveTax.Length || sources.Length != amounts.Length)
        {
            throw new InvalidOperationException("The columns of table " + TaxProvisionTable
                + " do not have a consistent number of rows.");
        }

        return (sources, liveTax, amounts);
    }

    /// <summary>
    /// Reads the result of check K39 from the 90_Checks worksheet.
    /// </summary>
    /// <param name="label">What the measurement is reported under.</param>
    /// <param name="pass">The one based number of the pass being run.</param>
    /// <returns>The check result, ex. OK or FAIL.</returns>
    /// <remarks>
    /// The check results occupy a single row table whose headers are the check ids, so a check is located by its id
    /// rather than by its address.  K39 currently resolves to 90_Checks!AM5.
    /// </remarks>
    private string ReadCheckResult(string? label, int pass)
    {
        using IDisposable scope = _timer.Measure(Phase("read check " + TaxProvisionCheckId, label, pass));

        object?[] values = ExcelUtils.ReadColumn(_checks!, TaxProvisionCheckId);

        if (values.Length != 1)
        {
            throw new InvalidOperationException("Check " + TaxProvisionCheckId + " on worksheet " + ChecksWorksheet
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
}
