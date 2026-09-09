using System.Globalization;
using System.Runtime.InteropServices;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pfm;

/// <summary>
/// The high level operations of the pfm CLI.  Each public method implements one member of <see cref="Commands"/> and
/// returns the process exit code for it.
/// </summary>
public static class Operations
{
    private const string EconomyWorksheet = "21_EconomyHistorical";
    private const string EconomyTable = "EconomyHistorical";
    private const string EconomyYearColumn = "Year";
    private const string EconomySemesterColumn = "Semester";

    private const string SimulationModeName = "SimulationMode";
    private const string HistoricalModeName = "Hist";
    private const string BackTestStartYearName = "BackTestingStartYear";
    private const string BackTestStartSemesterName = "BackTestingStartSemester";
    private const string SemestersPerYearName = "SemestersPerYear";

    private const string BackTestResultsName = "backtest";

    /// <summary>
    /// The number of times a simulation whose tax provision did not converge is settled again, with damping, before
    /// its result is recorded unsettled.
    /// </summary>
    private const int MaxRetries = 3;

    /// <summary>
    /// Iterates the tax provision until it agrees with the modelled total tax.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>Zero when the provision converged; otherwise one.</returns>
    /// <remarks>
    /// The iteration itself lives in <see cref="TaxProvisionIterator"/>, because a simulation has to settle the tax the
    /// same way once per simulated period.  What belongs to this command alone is acquiring the workbook the user
    /// pointed at, and saving it afterwards.
    /// </remarks>
    public static int Iterate(Arguments arguments, TextWriter output, TextWriter error)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(error);

        var timer = new DiagnosticTimer(arguments.Timing, output);

        try
        {
            return RunIterate(arguments, timer, output, error);
        }
        catch (COMException ex)
        {
            return Fail("PFM_ITERATE_EXCEL_ERROR: " + ex.Message, "Excel reported an error: " + ex.Message, error);
        }
        catch (InvalidOperationException ex)
        {
            return Fail("PFM_ITERATE_WORKBOOK_ERROR: " + ex.Message, ex.Message, error);
        }
        catch (IOException ex)
        {
            return Fail("PFM_ITERATE_IO_ERROR: " + ex.Message, ex.Message, error);
        }
        catch (UnauthorizedAccessException ex)
        {
            return Fail("PFM_ITERATE_ACCESS_ERROR: " + ex.Message, ex.Message, error);
        }
        finally
        {
            timer.ReportTotal();
        }
    }

    /// <summary>
    /// Back tests the workbook against every historical period long enough to project over.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>Zero when the sweep completed; otherwise one.</returns>
    /// <remarks>
    /// Each simulation drives the projection from one starting point in 21_EconomyHistorical: setting the back testing
    /// start moves 30_MarketSimulation onto a different run of historical semesters, and the whole model reprojects
    /// from it.  The starting point then advances one semester at a time, so the sweep asks what this plan would have
    /// done had it begun in every period history is long enough to have finished in.
    /// </remarks>
    public static int BackTest(Arguments arguments, TextWriter output, TextWriter error)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(error);

        var timer = new DiagnosticTimer(arguments.Timing, output);

        try
        {
            return RunBackTest(arguments, timer, output, error);
        }
        catch (COMException ex)
        {
            return Fail("PFM_BACKTEST_EXCEL_ERROR: " + ex.Message, "Excel reported an error: " + ex.Message, error);
        }
        catch (InvalidOperationException ex)
        {
            return Fail("PFM_BACKTEST_WORKBOOK_ERROR: " + ex.Message, ex.Message, error);
        }
        catch (IOException ex)
        {
            return Fail("PFM_BACKTEST_IO_ERROR: " + ex.Message, ex.Message, error);
        }
        catch (UnauthorizedAccessException ex)
        {
            return Fail("PFM_BACKTEST_ACCESS_ERROR: " + ex.Message, ex.Message, error);
        }
        finally
        {
            timer.ReportTotal();
        }
    }

    /// <summary>
    /// Acquires the workbook and runs the iteration to completion.  The exceptions this may raise are handled by
    /// <see cref="Iterate"/>.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>Zero when the provision converged; otherwise one.</returns>
    private static int RunIterate(Arguments arguments, DiagnosticTimer timer, TextWriter output, TextWriter error)
    {
        ExcelSession session = OpenWorkbook(arguments.FilePath, timer);

        try
        {
            session.SuspendCalculation();

            // Settle the workbook before the first read.  A check cell left dirty by an earlier edit would otherwise
            // report FAIL for a provision that is already in agreement.
            Calculate(session, timer, "initial calculate");

            using var iterator = new TaxProvisionIterator(session, timer);
            TaxProvisionOutcome outcome = iterator.Run();

            if (outcome.Status != TaxProvisionStatus.Converged)
            {
                return Fail("PFM_ITERATE_" + outcome.Status.ToString().ToUpperInvariant() + ": passes="
                    + Format(outcome.Passes), DescribeFailure(outcome), error);
            }

            bool saved = SaveUnlessAlreadyOpenInExcel(session, timer);

            output.WriteLine("The tax provision converged.  Check " + TaxProvisionIterator.TaxProvisionCheckId
                + " reported OK after " + Format(outcome.Passes) + " iterations."
                + (saved ? string.Empty : "  The workbook was already open in Excel and has not been saved."));
            return 0;
        }
        finally
        {
            // Disposing a session that started Excel closes the workbook and quits, which is itself slow enough to be
            // worth measuring.
            using IDisposable scope = timer.Measure("close workbook");
            session.Dispose();
        }
    }

    /// <summary>
    /// Prepares the run directory, drives every simulation of this job's slice and records the results.  The exceptions
    /// this may raise are handled by <see cref="BackTest"/>.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>Zero when the sweep completed; otherwise one.</returns>
    private static int RunBackTest(Arguments arguments, DiagnosticTimer timer, TextWriter output, TextWriter error)
    {
        // Declared before the session so that it is disposed after it: Excel holds the copy open until the session
        // that drove it is closed, and the copy is what disposing the run deletes.
        using SimulationRun run = SimulationRun.Create(arguments.FilePath, BackTestResultsName, arguments.JobCount,
            arguments.JobIndex);

        output.WriteLine("Job " + Format(arguments.JobIndex) + " of " + Format(arguments.JobCount) + " is working in "
            + run.Directory + ".");

        ExcelSession session = StartWorkbook(run.WorkbookPath, timer);

        try
        {
            session.SuspendCalculation();

            // Hist is what makes 30_MarketSimulation read the historical semesters rather than the constant or the
            // Monte Carlo drivers, so nothing the sweep does afterwards would reach the model without it.
            ExcelUtils.SetNameValue(session.Workbook, SimulationModeName,
                ExcelUtils.GetNameConstantText(session.Workbook, HistoricalModeName));

            (int[] startYears, int[] startSemesters) = ReadHistoricalPeriods(session, timer);

            Calculate(session, timer, "initial calculate");

            // After the initial calculate, because the synopsis is formulas: read before it, and the file would record
            // the configuration the workbook was last saved with rather than the one this sweep is about to run.
            WriteRunConfiguration(session, run, timer, output);

            using var harvester = new SummaryHarvester(session);

            int semestersPerYear = (int)ExcelUtils.GetNameNumber(session.Workbook, SemestersPerYearName);
            int projectedSemesters = harvester.ProjectionYears * semestersPerYear;

            // A simulation starting at row r reads the projected semesters from rows r through r + projected - 1, so
            // the last one history can carry starts at row (rows - projected + 1) and finishes exactly on the last row.
            // Counting the starts rather than subtracting the horizon is what keeps that final period in the sweep;
            // 30_MarketSimulation agrees, returning NA only once a projection would read past the last row.
            int total = startYears.Length - projectedSemesters + 1;

            if (total <= 0)
            {
                return Fail("PFM_BACKTEST_NO_SIMULATIONS: rows=" + Format(startYears.Length) + " projectionYears="
                    + Format(harvester.ProjectionYears),
                    "Table " + EconomyTable + " holds " + Format(startYears.Length) + " semesters, which is not enough "
                    + "to project " + Format(harvester.ProjectionYears) + " years from even once.", error);
            }

            JobPartition partition = JobPartition.Create(total, arguments.JobCount, arguments.JobIndex);

            output.WriteLine("The sweep is " + Format(total) + " simulations of " + Format(harvester.ProjectionYears)
                + " years each.  This job runs " + Format(partition.Count) + " of them, starting at "
                + Format(partition.StartOffset) + ".");

            return RunSimulations(session, run, harvester, partition, startYears, startSemesters, timer, output);
        }
        finally
        {
            using IDisposable scope = timer.Measure("close workbook");
            session.Dispose();
        }
    }

    /// <summary>
    /// Drives this job's slice of the sweep, recording each simulation as it finishes.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="run">The job's working directory and results file.</param>
    /// <param name="harvester">The harvester that reads the outcome metrics.</param>
    /// <param name="partition">The slice of the sweep this job runs.</param>
    /// <param name="startYears">The year of each period of 21_EconomyHistorical, in table order.</param>
    /// <param name="startSemesters">The semester of each period of 21_EconomyHistorical, in table order.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <returns>Zero when the sweep completed; otherwise one.</returns>
    /// <remarks>
    /// A simulation whose tax never settles is recorded rather than abandoned: DriftWithinTolerance is exactly the
    /// column that says so, and dropping the row would quietly bias the sweep towards the periods that were easy to
    /// settle.  The job still reports at the end how many of its simulations came out that way.  It is the outcome of
    /// the attempt that settled the provision, or of the last one that failed to, which is recorded.
    /// </remarks>
    private static int RunSimulations(ExcelSession session, SimulationRun run, SummaryHarvester harvester,
        JobPartition partition, int[] startYears, int[] startSemesters, DiagnosticTimer timer, TextWriter output)
    {
        using var writer = new SimulationResultWriter(run.ResultsPath, SummaryHarvester.DataElements,
            harvester.Years);
        using var iterator = new TaxProvisionIterator(session, timer);

        int unsettled = 0;

        for (int offset = partition.StartOffset; offset < partition.StartOffset + partition.Count; offset++)
        {
            int startYear = startYears[offset];
            int startSemester = startSemesters[offset];
            string label = "simulation " + Format(offset) + " (" + Format(startYear) + " S" + Format(startSemester)
                + ")";

            // Writing the start moves 30_MarketSimulation onto a different run of historical semesters; the
            // recalculation that follows reprojects the whole model from it.
            ExcelUtils.SetNameValue(session.Workbook, BackTestStartYearName, (double)startYear);
            ExcelUtils.SetNameValue(session.Workbook, BackTestStartSemesterName, (double)startSemester);
            Calculate(session, timer, "calculate " + label);

            TaxProvisionOutcome outcome = SettleTaxProvision(iterator, label, output);

            if (outcome.Status != TaxProvisionStatus.Converged)
            {
                unsettled++;
                Log.Logger.Warning("PFM_BACKTEST_UNSETTLED: " + label + " " + DescribeFailure(outcome));
            }

            writer.Write(new SimulationRecord
            {
                Index = offset,
                StartYear = startYear,
                StartSemester = startSemester,
                Elements = harvester.Harvest(),
                TaxProvision = outcome
            });

            output.WriteLine("Recorded " + label + ": " + Format(outcome.Passes) + " passes, drift "
                + Format(outcome.Drift) + (outcome.WithinTolerance ? string.Empty : " OUT OF TOLERANCE") + ".");
        }

        Log.Logger.Information("PFM_BACKTEST_COMPLETE: simulations=" + Format(partition.Count) + " unsettled="
            + Format(unsettled) + " results=" + run.ResultsPath);

        output.WriteLine("Recorded " + Format(partition.Count) + " simulations to " + run.ResultsPath + "."
            + (unsettled == 0
                ? string.Empty
                : "  " + Format(unsettled) + " of them left the tax provision out of tolerance; their rows report "
                    + "DriftWithinTolerance as FALSE."));

        return 0;
    }

    /// <summary>
    /// Settles the tax provision of one simulation, retrying with damping when it does not converge.
    /// </summary>
    /// <param name="iterator">The iterator that settles the provision.</param>
    /// <param name="label">What the measurements of the run are reported under.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <returns>The outcome of the attempt that converged, or of the last attempt when none did.</returns>
    /// <remarks>
    /// The first attempt is undamped because that is the faster of the two on a provision that converges, which nearly
    /// every simulation of a sweep does.  A retry damps instead, which settles the oscillation between two values that
    /// is what an undamped attempt cannot leave, at the cost of the extra passes that are why the first attempt does
    /// not pay for it.
    /// <para>
    /// A retry starts from wherever the failed attempt left the provision rather than from a reset one.  A damped pass
    /// converges on the same fixed point from any starting value, and the one the failed attempt reached is no worse a
    /// starting value than any other.  It also makes a retry of a stalled provision cost a single read: a provision
    /// that no pass can move is one that no damped pass can move either, so the retry returns at pass zero.
    /// </para>
    /// </remarks>
    private static TaxProvisionOutcome SettleTaxProvision(TaxProvisionIterator iterator, string label,
        TextWriter output)
    {
        TaxProvisionOutcome outcome = iterator.Run(label);

        for (int retry = 1; retry <= MaxRetries && outcome.Status != TaxProvisionStatus.Converged; retry++)
        {
            Log.Logger.Warning("PFM_BACKTEST_RETRY: " + label + " retry=" + Format(retry) + " of "
                + Format(MaxRetries) + " " + DescribeFailure(outcome));

            output.WriteLine("Retrying " + label + " with damping (" + Format(retry) + " of " + Format(MaxRetries)
                + "): the tax provision did not converge in " + Format(outcome.Passes) + " passes.");

            outcome = iterator.Run(label + " retry " + Format(retry), TaxProvisionIterator.DampedGain);
        }

        return outcome;
    }

    /// <summary>
    /// Records the configuration of the workbook the sweep is about to exercise, measuring how long it takes.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="run">The job's working directory and results file.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <remarks>
    /// The copy of the workbook is deleted when the job ends, so the synopsis is what a results file is read against
    /// afterwards.  Writing it before the first simulation also means an interrupted sweep still says what it was
    /// running.
    /// </remarks>
    private static void WriteRunConfiguration(ExcelSession session, SimulationRun run, DiagnosticTimer timer,
        TextWriter output)
    {
        using IDisposable scope = timer.Measure("write run configuration");

        string path = RunConfiguration.Write(session.Workbook, run.Directory);
        output.WriteLine("The configuration of the workbook this job runs is recorded in " + path + ".");
    }

    /// <summary>
    /// Reads the year and semester of every period of the historical economy, in table order.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <returns>The years and the semesters, one entry per period.</returns>
    /// <remarks>
    /// The periods are read from the table rather than counted forward from a first year, so a sweep starts where the
    /// data starts and steps the way the data steps.  Extending 21_EconomyHistorical backwards or forwards therefore
    /// changes what the sweep covers without changing this tool.
    /// </remarks>
    private static (int[] Years, int[] Semesters) ReadHistoricalPeriods(ExcelSession session, DiagnosticTimer timer)
    {
        using IDisposable scope = timer.Measure("read historical periods");

        Excel.ListObject economy = ExcelUtils.GetListObject(session.Workbook, EconomyWorksheet, EconomyTable);
        try
        {
            int[] years = ReadWholeNumbers(economy, EconomyYearColumn);
            int[] semesters = ReadWholeNumbers(economy, EconomySemesterColumn);

            if (years.Length != semesters.Length)
            {
                throw new InvalidOperationException("The columns of table " + EconomyTable
                    + " do not have a consistent number of rows.");
            }

            return (years, semesters);
        }
        finally
        {
            ExcelUtils.ReleaseComObject(economy);
        }
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

    /// <summary>
    /// Acquires the workbook the user pointed at, measuring how long it takes.
    /// </summary>
    /// <param name="filePath">The path to the workbook.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <returns>The session that owns the workbook.</returns>
    /// <remarks>
    /// This is the phase that dominates a run which has to start Excel rather than attach to it, so it is measured on
    /// its own rather than as part of the work that follows.
    /// </remarks>
    private static ExcelSession OpenWorkbook(string filePath, DiagnosticTimer timer)
    {
        using IDisposable scope = timer.Measure("open workbook");
        return ExcelUtils.OpenWorkbook(filePath);
    }

    /// <summary>
    /// Opens a job's own copy of the workbook in an Excel process of its own, measuring how long it takes.
    /// </summary>
    /// <param name="filePath">The path to the copy.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <returns>The session that owns the workbook.</returns>
    private static ExcelSession StartWorkbook(string filePath, DiagnosticTimer timer)
    {
        using IDisposable scope = timer.Measure("open workbook");
        return ExcelUtils.StartIsolatedWorkbook(filePath);
    }

    /// <summary>
    /// Recalculates the workbook, measuring how long it takes.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="phase">The name the measurement is reported under.</param>
    private static void Calculate(ExcelSession session, DiagnosticTimer timer, string phase)
    {
        using IDisposable scope = timer.Measure(phase);
        session.Calculate();
    }

    /// <summary>
    /// Saves the workbook unless it was already open in Excel when the session attached to it.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <returns>True when the workbook was saved.</returns>
    /// <remarks>
    /// Saving a workbook this size is expensive, and one that was already open belongs to whoever opened it: they can
    /// see the iterated values in front of them and save on their own terms.  A run that opened the workbook itself
    /// has no such owner, so it must save or the iteration would be discarded when the session closes the workbook.
    /// </remarks>
    private static bool SaveUnlessAlreadyOpenInExcel(ExcelSession session, DiagnosticTimer timer)
    {
        if (session.WasAlreadyOpen)
        {
            Log.Logger.Information("PFM_ITERATE_SAVE_SKIPPED: the workbook was already open in Excel.");
            return false;
        }

        using IDisposable scope = timer.Measure("save workbook");
        session.Save();
        return true;
    }

    /// <summary>
    /// Describes an iteration that did not converge.
    /// </summary>
    /// <param name="outcome">The outcome to describe.</param>
    /// <returns>The description.</returns>
    private static string DescribeFailure(TaxProvisionOutcome outcome)
    {
        string check = "Check " + TaxProvisionIterator.TaxProvisionCheckId + " reports " + outcome.CheckResult
            + " with a drift of " + Format(outcome.Drift) + " against a tolerance of " + Format(outcome.Tolerance)
            + " after " + Format(outcome.Passes) + " passes.";

        return outcome.Status == TaxProvisionStatus.Stalled
            ? "The tax provision cannot converge.  " + check
                + "  Every modelled year already matches TotalTaxLive, so no further pass can move it."
            : "The tax provision did not converge.  " + check;
    }

    /// <summary>
    /// Logs and writes a failure, and yields the exit code for it.
    /// </summary>
    /// <param name="logMessage">The message to log.</param>
    /// <param name="errorMessage">The message to write to the error writer.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>One.</returns>
    private static int Fail(string logMessage, string errorMessage, TextWriter error)
    {
        Log.Logger.Error(logMessage);
        error.WriteLine(errorMessage);
        return 1;
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
