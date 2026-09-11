using System.Diagnostics.CodeAnalysis;
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
    private const string SimulationModeName = "SimulationMode";
    private const string HistoricalModeName = "Hist";
    private const string MonteCarloModeName = "MonteCarlo";
    private const string SemestersPerYearName = "SemestersPerYear";

    private const string BackTestResultsName = "backtest";
    private const string MonteCarloResultsName = "montecarlo";

    private const string RunConfigurationWorksheet = "RunConfiguration";
    private const string CoalescedWorkbookPrefix = "PortfolioSimAnalysis.";
    private const string CoalescedWorkbookExtension = ".xlsx";

    /// <summary>
    /// The number of result rows imported into the coalesced workbook per cross process call.  A block is one call
    /// whatever its size, so this trades the memory a block occupies against the number of calls a large sweep costs.
    /// </summary>
    private const int ImportBlockRows = 2000;

    /// <summary>
    /// The number of rows a worksheet holds, which is what limits how large a sweep can be coalesced into one sheet.
    /// </summary>
    private const int WorksheetRowLimit = 1048576;

    /// <summary>
    /// The number of times a simulation whose materialized series did not converge is settled again, with damping,
    /// before its result is recorded unsettled.
    /// </summary>
    private const int MaxRetries = 3;

    /// <summary>
    /// Iterates every materialized series of the workbook until each agrees with the live value it stands for.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>Zero when the workbook converged; otherwise one.</returns>
    /// <remarks>
    /// The iteration itself lives in <see cref="ConvergenceIterator"/>, because a simulation has to settle the same
    /// values the same way once per simulated period.  What belongs to this command alone is acquiring the workbook
    /// the user pointed at, and saving it afterwards.
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
    /// Runs a Monte Carlo campaign over the workbook, one simulation per iteration 22_MCSeeds holds seeds for.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>Zero when the sweep completed; otherwise one.</returns>
    /// <remarks>
    /// Each simulation stamps MCIteration, which selects the column of seeds the block bootstrap on
    /// 30_MarketSimulation draws its block start rows from: a Monte Carlo path is four spliced runs of real history
    /// rather than a sequence of independent draws, and the iteration is the whole of what selects one.  The seed
    /// sheet is static and nothing in the driver set is random, so an iteration is exactly reproducible.
    /// </remarks>
    public static int MonteCarlo(Arguments arguments, TextWriter output, TextWriter error)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(error);

        var timer = new DiagnosticTimer(arguments.Timing, output);

        try
        {
            return RunMonteCarlo(arguments, timer, output);
        }
        catch (COMException ex)
        {
            return Fail("PFM_MONTECARLO_EXCEL_ERROR: " + ex.Message, "Excel reported an error: " + ex.Message, error);
        }
        catch (InvalidOperationException ex)
        {
            return Fail("PFM_MONTECARLO_WORKBOOK_ERROR: " + ex.Message, ex.Message, error);
        }
        catch (IOException ex)
        {
            return Fail("PFM_MONTECARLO_IO_ERROR: " + ex.Message, ex.Message, error);
        }
        catch (UnauthorizedAccessException ex)
        {
            return Fail("PFM_MONTECARLO_ACCESS_ERROR: " + ex.Message, ex.Message, error);
        }
        finally
        {
            timer.ReportTotal();
        }
    }

    /// <summary>
    /// Gathers the run directories of one simulation sweep into a single Excel workbook.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>Zero when the workbook was written; otherwise one.</returns>
    /// <remarks>
    /// A sweep leaves one directory per job, each holding that job's slice of the results and a copy of the
    /// configuration the whole sweep was driven from.  This is what turns that back into the one thing it describes: a
    /// workbook stating the configuration on one sheet and the results of every job, concatenated in job order, on
    /// another.  It reads the sweep and never writes to it, so it can be run against a set of run directories as often
    /// as it is useful to.
    /// </remarks>
    public static int Coalesce(Arguments arguments, TextWriter output, TextWriter error)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(error);

        var timer = new DiagnosticTimer(arguments.Timing, output);

        try
        {
            return RunCoalesce(arguments, timer, output, error);
        }
        catch (COMException ex)
        {
            return Fail("PFM_COALESCE_EXCEL_ERROR: " + ex.Message, "Excel reported an error: " + ex.Message, error);
        }
        catch (InvalidOperationException ex)
        {
            return Fail("PFM_COALESCE_INPUT_ERROR: " + ex.Message, ex.Message, error);
        }
        catch (IOException ex)
        {
            return Fail("PFM_COALESCE_IO_ERROR: " + ex.Message, ex.Message, error);
        }
        catch (UnauthorizedAccessException ex)
        {
            return Fail("PFM_COALESCE_ACCESS_ERROR: " + ex.Message, ex.Message, error);
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
        ExcelSession session = OpenWorkbook(RequireFilePath(arguments), timer);

        try
        {
            session.SuspendCalculation();

            // Settle the workbook before the first read.  A check cell left dirty by an earlier edit would otherwise
            // report FAIL for a series that is already in agreement.
            Calculate(session, timer, "initial calculate");

            using var iterator = new ConvergenceIterator(session, timer);
            ConvergenceOutcome outcome = iterator.Run();

            if (outcome.Status != ConvergenceStatus.Converged)
            {
                return Fail("PFM_ITERATE_" + outcome.Status.ToString().ToUpperInvariant() + ": passes="
                    + Format(outcome.Passes), DescribeFailure(outcome), error);
            }

            bool saved = SaveUnlessAlreadyOpenInExcel(session, timer);

            output.WriteLine("The workbook converged.  Checks "
                + string.Join(", ", outcome.Series.Select(series => series.CheckId))
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
        using SimulationRun run = SimulationRun.Create(RequireFilePath(arguments), BackTestResultsName,
            arguments.JobCount, arguments.JobIndex);

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

            BackTestSweepPlan plan = ReadBackTestPlan(session, timer);

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
            int available = plan.PeriodCount - projectedSemesters + 1;

            if (available <= 0)
            {
                return Fail("PFM_BACKTEST_NO_SIMULATIONS: rows=" + Format(plan.PeriodCount) + " projectionYears="
                    + Format(harvester.ProjectionYears),
                    "The historical economy holds " + Format(plan.PeriodCount) + " semesters, which is not enough "
                    + "to project " + Format(harvester.ProjectionYears) + " years from even once.", error);
            }

            int total = CountSimulations(available, arguments, plan);

            JobPartition partition = JobPartition.Create(total, arguments.JobCount, arguments.JobIndex);

            output.WriteLine("The sweep is " + Format(total) + " simulations of " + Format(harvester.ProjectionYears)
                + " years each" + DescribeLimit(total, available) + ".  This job runs " + Format(partition.Count)
                + " of them, starting at " + Format(partition.StartOffset) + ".");

            return RunSimulations(session, run, harvester, partition, plan, timer, output);
        }
        finally
        {
            using IDisposable scope = timer.Measure("close workbook");
            session.Dispose();
        }
    }

    /// <summary>
    /// Prepares the run directory, drives every iteration of this job's slice and records the results.  The exceptions
    /// this may raise are handled by <see cref="MonteCarlo"/>.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <returns>Zero when the sweep completed; otherwise one.</returns>
    /// <remarks>
    /// This is the back test with a different driver: the same copied workbook, the same recorded configuration, the
    /// same convergence and the same harvest, differing only in what a simulation stamps on 10_Parameters, which is
    /// what <see cref="MonteCarloSweepPlan"/> states.  A workbook that holds no seeds is reported by the plan rather
    /// than here, because the count of them is the sweep itself.
    /// </remarks>
    private static int RunMonteCarlo(Arguments arguments, DiagnosticTimer timer, TextWriter output)
    {
        // Declared before the session so that it is disposed after it: Excel holds the copy open until the session
        // that drove it is closed, and the copy is what disposing the run deletes.
        using SimulationRun run = SimulationRun.Create(RequireFilePath(arguments), MonteCarloResultsName,
            arguments.JobCount, arguments.JobIndex);

        output.WriteLine("Job " + Format(arguments.JobIndex) + " of " + Format(arguments.JobCount) + " is working in "
            + run.Directory + ".");

        ExcelSession session = StartWorkbook(run.WorkbookPath, timer);

        try
        {
            session.SuspendCalculation();

            // MonteCarlo is what makes the mode selector on 30_MarketSimulation read the MC_ columns rather than the
            // constant or back testing drivers, so nothing an iteration stamps afterwards would reach the model
            // without it.
            ExcelUtils.SetNameValue(session.Workbook, SimulationModeName,
                ExcelUtils.GetNameConstantText(session.Workbook, MonteCarloModeName));

            MonteCarloSweepPlan plan = ReadMonteCarloPlan(session, timer);

            Calculate(session, timer, "initial calculate");

            // After the initial calculate, because the synopsis is formulas: read before it, and the file would record
            // the configuration the workbook was last saved with rather than the one this sweep is about to run.
            WriteRunConfiguration(session, run, timer, output);

            using var harvester = new SummaryHarvester(session);

            int total = CountSimulations(plan.IterationCount, arguments, plan);

            JobPartition partition = JobPartition.Create(total, arguments.JobCount, arguments.JobIndex);

            output.WriteLine("The sweep is " + Format(total) + " iterations of "
                + Format(harvester.ProjectionYears) + " years each"
                + DescribeLimit(total, plan.IterationCount) + ".  This job runs " + Format(partition.Count)
                + " of them, starting at iteration " + Format(plan.Iteration(partition.StartOffset)) + ".");

            return RunSimulations(session, run, harvester, partition, plan, timer, output);
        }
        finally
        {
            using IDisposable scope = timer.Measure("close workbook");
            session.Dispose();
        }
    }

    /// <summary>
    /// Discovers the sweep, builds the workbook and saves it.  The exceptions this may raise are handled by
    /// <see cref="Coalesce"/>.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>Zero when the workbook was written; otherwise one.</returns>
    /// <remarks>
    /// The sweep is discovered and the target named before Excel is started, so a set of run directories that is not a
    /// complete sweep is reported in the time it takes to list a directory rather than after a workbook has been
    /// created for it.
    /// </remarks>
    private static int RunCoalesce(Arguments arguments, DiagnosticTimer timer, TextWriter output, TextWriter error)
    {
        CoalescedSweep sweep = DiscoverSweep(arguments.InputPath, timer);

        output.WriteLine("The sweep in " + Path.GetFullPath(arguments.InputPath) + " is " + Format(sweep.Jobs.Count)
            + " " + sweep.TestType + " jobs, run " + sweep.RunId + ".");

        string outputDirectory = Path.GetFullPath(arguments.OutputPath);
        Directory.CreateDirectory(outputDirectory);

        string targetPath = Path.Combine(outputDirectory,
            CoalescedWorkbookPrefix + sweep.RunId + CoalescedWorkbookExtension);

        // The run id is derived from the sweep, so coalescing one twice names the same workbook both times.  That is
        // what makes the name meaningful, and it is also why an existing one is reported rather than written over: the
        // file that is already there is this same sweep, and replacing it silently would discard whatever has since
        // been done to it.
        if (File.Exists(targetPath))
        {
            return Fail("PFM_COALESCE_TARGET_EXISTS: " + targetPath,
                "The workbook " + targetPath + " already holds this sweep.  Delete it, or choose another "
                + "--output-path, to coalesce the sweep again.", error);
        }

        ExcelSession session = CreateWorkbook(timer);

        try
        {
            session.SuspendCalculation();

            WriteRunConfigurationWorksheet(session, sweep, timer, output);
            int rows = ImportResults(session, sweep, timer, output);

            using (IDisposable scope = timer.Measure("save workbook"))
            {
                session.SaveAs(targetPath);
            }

            Log.Logger.Information("PFM_COALESCE_COMPLETE: runId=" + sweep.RunId + " jobs=" + Format(sweep.Jobs.Count)
                + " rows=" + Format(rows) + " workbook=" + targetPath);

            output.WriteLine("Coalesced " + Format(rows) + " rows from " + Format(sweep.Jobs.Count) + " jobs into "
                + targetPath + ".");

            return 0;
        }
        finally
        {
            using IDisposable scope = timer.Measure("close workbook");
            session.Dispose();
        }
    }

    /// <summary>
    /// Finds the run directories of the sweep and checks that they are a complete set, measuring how long it takes.
    /// </summary>
    /// <param name="inputPath">The directory holding the run directories.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <returns>The discovered sweep.</returns>
    private static CoalescedSweep DiscoverSweep(string inputPath, DiagnosticTimer timer)
    {
        using IDisposable scope = timer.Measure("discover sweep");
        return CoalescedSweep.Discover(inputPath);
    }

    /// <summary>
    /// Renames the workbook's first worksheet and writes the configuration the sweep was driven from into it, one line
    /// of the file per row.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="sweep">The sweep being coalesced.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="output">The writer for normal output.</param>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    private static void WriteRunConfigurationWorksheet(ExcelSession session, CoalescedSweep sweep,
        DiagnosticTimer timer, TextWriter output)
    {
        using IDisposable scope = timer.Measure("write run configuration");

        Excel.Worksheet worksheet = ExcelUtils.GetWorksheet(session.Workbook, 1);
        try
        {
            worksheet.Name = RunConfigurationWorksheet;

            // The synopsis is a report rather than data: the column is formatted as text before anything is written,
            // so a line reading like a date or beginning with an equals sign is stored as the workbook wrote it.
            ExcelUtils.FormatColumnAsText(worksheet, 1);

            if (sweep.RunConfiguration.Count > 0)
            {
                var lines = new object?[sweep.RunConfiguration.Count, 1];
                for (int row = 0; row < sweep.RunConfiguration.Count; row++)
                {
                    lines[row, 0] = sweep.RunConfiguration[row];
                }

                ExcelUtils.WriteGrid(worksheet, 1, 1, lines);
            }
        }
        finally
        {
            ExcelUtils.ReleaseComObject(worksheet);
        }

        output.WriteLine("Wrote " + Format(sweep.RunConfiguration.Count) + " lines of run configuration to "
            + RunConfigurationWorksheet + ".");
    }

    /// <summary>
    /// Imports the results of every job into one worksheet, in job order.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="sweep">The sweep being coalesced.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <returns>The number of data rows imported, not counting the header.</returns>
    /// <remarks>
    /// The header is written once, from the first job, and every other job's is checked against it.  Jobs of one sweep
    /// drive copies of one workbook, so their headers agree; two that do not are results of different plans, or of
    /// different projection horizons, and concatenating them would produce a sheet whose columns mean different things
    /// in different rows.
    /// </remarks>
    private static int ImportResults(ExcelSession session, CoalescedSweep sweep, DiagnosticTimer timer,
        TextWriter output)
    {
        using IDisposable scope = timer.Measure("import results");

        Excel.Worksheet worksheet = ExcelUtils.AddWorksheet(session.Workbook, sweep.DataWorksheetName);
        try
        {
            IReadOnlyList<string>? header = null;
            int nextRow = 1;

            foreach (SweepJob job in sweep.Jobs)
            {
                using var reader = new SimulationResultReader(job.ResultsPath);

                if (header is null)
                {
                    header = reader.Header;
                    WriteHeaderRow(worksheet, header);
                    nextRow++;
                }
                else
                {
                    CheckHeaderMatches(header, reader.Header, sweep.Jobs[0], job);
                }

                int imported = ImportJob(worksheet, reader, ref nextRow, job);

                output.WriteLine("Imported " + Format(imported) + " rows from " + job.Name + ".");
            }

            // Every row of the sheet but the header, which the first job wrote.
            return nextRow - 2;
        }
        finally
        {
            ExcelUtils.ReleaseComObject(worksheet);
        }
    }

    /// <summary>
    /// Imports one job's results, a block of rows at a time.
    /// </summary>
    /// <param name="worksheet">The worksheet the results are imported into.</param>
    /// <param name="reader">The reader for the job's results file.</param>
    /// <param name="nextRow">The one based row the next block is written at, advanced by what is written.</param>
    /// <param name="job">The job being imported, named by a message about a sheet that has run out of rows.</param>
    /// <returns>The number of rows imported.</returns>
    private static int ImportJob(Excel.Worksheet worksheet, SimulationResultReader reader, ref int nextRow,
        SweepJob job)
    {
        int imported = 0;

        while (reader.ReadBlock(ImportBlockRows) is object?[,] block)
        {
            int rows = block.GetLength(0);

            if (nextRow + rows - 1 > WorksheetRowLimit)
            {
                throw new InvalidOperationException("The results of this sweep do not fit on one worksheet: importing "
                    + job.Name + " would pass row " + Format(WorksheetRowLimit) + ", which is the last row Excel "
                    + "holds.");
            }

            ExcelUtils.WriteGrid(worksheet, nextRow, 1, block);

            nextRow += rows;
            imported += rows;
        }

        return imported;
    }

    /// <summary>
    /// Writes the column headings of the results into the first row of the worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet the results are imported into.</param>
    /// <param name="header">The column headings.</param>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    private static void WriteHeaderRow(Excel.Worksheet worksheet, IReadOnlyList<string> header)
    {
        var headings = new object?[1, header.Count];
        for (int column = 0; column < header.Count; column++)
        {
            headings[0, column] = header[column];
        }

        ExcelUtils.WriteGrid(worksheet, 1, 1, headings);
    }

    /// <summary>
    /// Checks that a job's results describe the same columns as the sweep's first job.
    /// </summary>
    /// <param name="expected">The column headings of the first job.</param>
    /// <param name="actual">The column headings of the job being imported.</param>
    /// <param name="first">The first job of the sweep.</param>
    /// <param name="job">The job being imported.</param>
    private static void CheckHeaderMatches(IReadOnlyList<string> expected, IReadOnlyList<string> actual,
        SweepJob first, SweepJob job)
    {
        if (expected.Count != actual.Count)
        {
            throw new InvalidOperationException("The results of " + job.Name + " have " + Format(actual.Count)
                + " columns, and those of " + first.Name + " have " + Format(expected.Count)
                + ".  These are not jobs of one sweep.");
        }

        for (int column = 0; column < expected.Count; column++)
        {
            if (!string.Equals(expected[column], actual[column], StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Column " + Format(column + 1) + " of the results of " + job.Name
                    + " is " + actual[column] + ", and of " + first.Name + " is " + expected[column]
                    + ".  These are not jobs of one sweep.");
            }
        }
    }

    /// <summary>
    /// Creates the workbook the sweep is coalesced into, measuring how long it takes.
    /// </summary>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <returns>The session that owns the workbook.</returns>
    private static ExcelSession CreateWorkbook(DiagnosticTimer timer)
    {
        using IDisposable scope = timer.Measure("create workbook");
        return ExcelUtils.CreateWorkbook();
    }

    /// <summary>
    /// Drives this job's slice of the sweep, recording each simulation as it finishes.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="run">The job's working directory and results file.</param>
    /// <param name="harvester">The harvester that reads the outcome metrics.</param>
    /// <param name="partition">The slice of the sweep this job runs.</param>
    /// <param name="plan">What the sweep drives, and how each of its simulations is identified.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <returns>Zero when the sweep completed; otherwise one.</returns>
    /// <remarks>
    /// Every sweep runs through here.  A back test and a Monte Carlo campaign differ in what a simulation stamps and
    /// in what the results call it, which is what the plan states; what happens around that, ex. the retries and the
    /// recording of a simulation that never settled, is the same work and is done in one place.
    /// <para>
    /// A simulation whose materialized values never settle is recorded rather than abandoned: DriftWithinTolerance is
    /// exactly the column that says so, and dropping the row would quietly bias the sweep towards the simulations that
    /// were easy to settle.  The job still reports at the end how many of its simulations came out that way.  It is
    /// the outcome of the attempt that converged, or of the last one that failed to, which is recorded.
    /// </para>
    /// </remarks>
    private static int RunSimulations(ExcelSession session, SimulationRun run, SummaryHarvester harvester,
        JobPartition partition, SweepPlan plan, DiagnosticTimer timer, TextWriter output)
    {
        using var writer = new SimulationResultWriter(run.ResultsPath, plan.IdentityColumns,
            SummaryHarvester.DataElements, harvester.Years, ConvergenceIterator.SeriesNames);
        using var iterator = new ConvergenceIterator(session, timer);

        int unsettled = 0;

        for (int offset = partition.StartOffset; offset < partition.StartOffset + partition.Count; offset++)
        {
            string label = plan.Describe(offset);

            // Stamping the parameters of one simulation on 10_Parameters is all it takes to move the whole model onto
            // a different market path; the recalculation that follows reprojects it from there.
            plan.Drive(session.Workbook, offset);
            Calculate(session, timer, "calculate " + label);

            ConvergenceOutcome outcome = Settle(iterator, plan, label, output);

            if (outcome.Status != ConvergenceStatus.Converged)
            {
                unsettled++;
                Log.Logger.Warning("PFM_" + plan.LogName + "_UNSETTLED: " + label + " " + DescribeFailure(outcome));
            }

            writer.Write(new SimulationRecord
            {
                Index = offset,
                Identity = plan.Identify(offset),
                Elements = harvester.Harvest(),
                Convergence = outcome
            });

            output.WriteLine("Recorded " + label + ": " + Format(outcome.Passes) + " passes, drift "
                + outcome.DescribeDrift() + (outcome.WithinTolerance ? string.Empty : " OUT OF TOLERANCE") + ".");
        }

        Log.Logger.Information("PFM_" + plan.LogName + "_COMPLETE: simulations=" + Format(partition.Count)
            + " unsettled=" + Format(unsettled) + " results=" + run.ResultsPath);

        output.WriteLine("Recorded " + Format(partition.Count) + " simulations to " + run.ResultsPath + "."
            + (unsettled == 0
                ? string.Empty
                : "  " + Format(unsettled) + " of them left a materialized series out of tolerance; their rows "
                    + "report DriftWithinTolerance as FALSE."));

        return 0;
    }

    /// <summary>
    /// Settles the materialized series of one simulation, retrying with damping when they do not converge.
    /// </summary>
    /// <param name="iterator">The iterator that settles the series.</param>
    /// <param name="plan">The sweep being run, which names the tag its log messages carry.</param>
    /// <param name="label">What the measurements of the run are reported under.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <returns>The outcome of the attempt that converged, or of the last attempt when none did.</returns>
    /// <remarks>
    /// The first attempt is undamped because that is the faster of the two on a workbook that converges, which nearly
    /// every simulation of a sweep does.  A retry damps instead, which settles the oscillation between two values that
    /// is what an undamped attempt cannot leave, at the cost of the extra passes that are why the first attempt does
    /// not pay for it.
    /// <para>
    /// A retry starts from wherever the failed attempt left the series rather than from reset ones.  A damped pass
    /// converges on the same fixed point from any starting value, and the one the failed attempt reached is no worse a
    /// starting value than any other.  It also makes a retry of a stalled workbook cost a single read: a series that
    /// no pass can move is one that no damped pass can move either, so the retry returns at pass zero.
    /// </para>
    /// </remarks>
    private static ConvergenceOutcome Settle(ConvergenceIterator iterator, SweepPlan plan, string label,
        TextWriter output)
    {
        ConvergenceOutcome outcome = iterator.Run(label);

        for (int retry = 1; retry <= MaxRetries && outcome.Status != ConvergenceStatus.Converged; retry++)
        {
            Log.Logger.Warning("PFM_" + plan.LogName + "_RETRY: " + label + " retry=" + Format(retry) + " of "
                + Format(MaxRetries) + " " + DescribeFailure(outcome));

            output.WriteLine("Retrying " + label + " with damping (" + Format(retry) + " of " + Format(MaxRetries)
                + "): the workbook did not converge in " + Format(outcome.Passes) + " passes.");

            outcome = iterator.Run(label + " retry " + Format(retry), ConvergenceIterator.DampedGain);
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
    /// Yields the number of simulations the whole sweep runs: what the workbook offers, unless the caller asked for
    /// fewer.
    /// </summary>
    /// <param name="available">The number of simulations the workbook offers.</param>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="plan">The sweep being run, which names the tag its log messages carry.</param>
    /// <returns>The number of simulations to divide among the jobs.</returns>
    /// <remarks>
    /// A limit shortens the sweep rather than sampling it, so the simulations that run are the first ones: the
    /// earliest historical periods of a back test, the lowest iterations of a Monte Carlo campaign.  A shortened
    /// sweep is divided among its jobs exactly as a whole one is, which is what makes a trial run of a sweep that
    /// takes hours cost minutes without being a different thing from the sweep it is a trial of.
    /// <para>
    /// A limit larger than the sweep is not an error: it asks for at most that many simulations, and a workbook that
    /// offers fewer has already answered.
    /// </para>
    /// </remarks>
    private static int CountSimulations(int available, Arguments arguments, SweepPlan plan)
    {
        int total = arguments.SimulationCount is int limit && limit < available ? limit : available;

        Log.Logger.Information("PFM_" + plan.LogName + "_SWEEP: available=" + Format(available) + " simulations="
            + Format(total) + " jobs=" + Format(arguments.JobCount));

        return total;
    }

    /// <summary>
    /// Describes a sweep the caller shortened, for the message that says how large it is.
    /// </summary>
    /// <param name="total">The number of simulations the sweep runs.</param>
    /// <param name="available">The number of simulations the workbook offers.</param>
    /// <returns>The clause to append, or an empty string when the sweep runs whole.</returns>
    private static string DescribeLimit(int total, int available)
    {
        return total < available
            ? ", limited from the " + Format(available) + " the workbook offers"
            : string.Empty;
    }

    /// <summary>
    /// Reads the periods a back test sweeps, measuring how long it takes.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <returns>The plan.</returns>
    private static BackTestSweepPlan ReadBackTestPlan(ExcelSession session, DiagnosticTimer timer)
    {
        using IDisposable scope = timer.Measure("read historical periods");
        return BackTestSweepPlan.Create(session.Workbook);
    }

    /// <summary>
    /// Reads the iterations a Monte Carlo campaign sweeps, measuring how long it takes.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <returns>The plan.</returns>
    private static MonteCarloSweepPlan ReadMonteCarloPlan(ExcelSession session, DiagnosticTimer timer)
    {
        using IDisposable scope = timer.Measure("read seed columns");
        return MonteCarloSweepPlan.Create(session.Workbook);
    }

    /// <summary>
    /// Gets the path of the workbook a command was pointed at.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <returns>The path.</returns>
    /// <remarks>
    /// The option is required by every command that reaches this, so the parser has already refused a run without one.
    /// This states that where the path is used, so that a command added later without the option fails saying what is
    /// missing rather than somewhere further in with a null.
    /// </remarks>
    private static string RequireFilePath(Arguments arguments)
    {
        return arguments.FilePath
            ?? throw new InvalidOperationException("The command " + arguments.Command
                + " operates on a workbook, and no --file-path was given.");
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
    private static string DescribeFailure(ConvergenceOutcome outcome)
    {
        string checks = "After " + Format(outcome.Passes) + " passes: " + outcome.DescribeFailingSeries() + ".";

        return outcome.Status == ConvergenceStatus.Stalled
            ? "The workbook cannot converge.  " + checks
                + "  Every row the iteration owns already matches its live value, so no further pass can move it."
            : "The workbook did not converge.  " + checks;
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
