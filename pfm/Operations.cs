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
    private const string TaxPaymentsWorksheet = "15_TaxPayments";
    private const string TaxProvisionTable = "TaxProvision";
    private const string TotalTaxLiveColumn = "TotalTaxLive";
    private const string ProvisionSourceColumn = "ProvisionSource";
    private const string ProvisionAmountColumn = "ProvisionAmount";
    private const string ProvisionModelledName = "ProvisionModelled";

    private const string ChecksWorksheet = "90_Checks";
    private const string ChecksTable = "Checks";
    private const string TaxProvisionCheckId = "K39";
    private const string CheckResultOk = "OK";

    /// <summary>
    /// The number of passes after which the tax provision is treated as failing to converge.
    /// </summary>
    private const int MaxIterations = 24;

    /// <summary>
    /// Iterates the tax provision until it agrees with the modelled total tax.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>Zero when the provision converged; otherwise one.</returns>
    /// <remarks>
    /// TaxProvision.ProvisionAmount feeds the tax funding that 55_Tax draws on, so changing it changes the very
    /// TaxAnnual.TotalTax it is meant to match.  A single pass therefore cannot settle the two; each pass copies
    /// TaxProvision.TotalTaxLive (the modelled total tax looked up by year) into TaxProvision.ProvisionAmount for the
    /// ProvisionModelled years and recalculates, until check K39 on 90_Checks reports OK.
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

            Excel.ListObject provisions = ExcelUtils.GetListObject(session.Workbook, TaxPaymentsWorksheet,
                TaxProvisionTable);
            try
            {
                string modelledSource = ExcelUtils.GetNameConstantText(session.Workbook, ProvisionModelledName);

                // Settle the workbook before the first read.  A check cell left dirty by an earlier edit would
                // otherwise report FAIL for a provision that is already in agreement.
                Calculate(session, timer, "initial calculate");

                return RunIterations(session, provisions, modelledSource, timer, output, error);
            }
            finally
            {
                ExcelUtils.ReleaseComObject(provisions);
            }
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
    /// Acquires the workbook, measuring how long it takes.
    /// </summary>
    /// <param name="filePath">The path to the workbook.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <returns>The session that owns the workbook.</returns>
    /// <remarks>
    /// This is the phase that dominates a run which has to start Excel rather than attach to it, so it is measured on
    /// its own rather than as part of the iteration that follows.
    /// </remarks>
    private static ExcelSession OpenWorkbook(string filePath, DiagnosticTimer timer)
    {
        using IDisposable scope = timer.Measure("open workbook");
        return ExcelUtils.OpenWorkbook(filePath);
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
    /// Runs the refresh and recalculate loop until check K39 reports OK or the iteration limit is reached.
    /// </summary>
    /// <param name="session">The session that owns the workbook.</param>
    /// <param name="provisions">The TaxProvision table.</param>
    /// <param name="modelledSource">The ProvisionSource value that marks a year as modelled.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>Zero when the provision converged; otherwise one.</returns>
    private static int RunIterations(ExcelSession session, Excel.ListObject provisions, string modelledSource,
        DiagnosticTimer timer, TextWriter output, TextWriter error)
    {
        int iterations = 0;

        while (true)
        {
            string result = ReadCheckResult(session.Workbook, TaxProvisionCheckId, timer, iterations + 1);

            if (string.Equals(result, CheckResultOk, StringComparison.OrdinalIgnoreCase))
            {
                bool saved = SaveUnlessAlreadyOpenInExcel(session, timer);

                string message = "The tax provision converged.  Check " + TaxProvisionCheckId + " reported "
                    + CheckResultOk + " after " + Format(iterations) + " iterations."
                    + (saved ? string.Empty : "  The workbook was already open in Excel and has not been saved.");
                Log.Logger.Information("PFM_ITERATE_CONVERGED: iterations=" + Format(iterations));
                output.WriteLine(message);
                return 0;
            }

            if (iterations >= MaxIterations)
            {
                return Fail("PFM_ITERATE_NOT_CONVERGED: iterations=" + Format(iterations),
                    "The tax provision did not converge.  Check " + TaxProvisionCheckId + " still reports " + result
                    + " after " + Format(MaxIterations) + " iterations.", error);
            }

            int updated = RefreshProvisionAmounts(provisions, modelledSource, timer, iterations + 1);
            iterations++;

            if (updated == 0)
            {
                // Every modelled year already holds its TotalTaxLive value, so no further pass can move the check.
                return Fail("PFM_ITERATE_STALLED: iterations=" + Format(iterations),
                    "The tax provision cannot converge.  Check " + TaxProvisionCheckId + " reports " + result
                    + " but every " + ProvisionModelledName + " year already matches " + TotalTaxLiveColumn + ".",
                    error);
            }

            Calculate(session, timer, Pass("calculate", iterations));

            Log.Logger.Information("PFM_ITERATE_PASS: iteration=" + Format(iterations) + " updated="
                + Format(updated));
        }
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
    /// Copies TotalTaxLive into ProvisionAmount for every year whose ProvisionSource is ProvisionModelled.
    /// </summary>
    /// <param name="provisions">The TaxProvision table.</param>
    /// <param name="modelledSource">The ProvisionSource value that marks a year as modelled.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="pass">The one based number of the pass being run.  It labels the measurements.</param>
    /// <returns>The number of rows whose ProvisionAmount differed from TotalTaxLive before the copy.</returns>
    /// <remarks>
    /// Years with any other provision source, ex. ProvisionExogenous, are excluded by design: their amounts are
    /// authored rather than modelled, and K39 does not compare them.  Only runs of consecutive modelled rows are
    /// written, so a non modelled row is never covered by a write.
    /// </remarks>
    private static int RefreshProvisionAmounts(Excel.ListObject provisions, string modelledSource,
        DiagnosticTimer timer, int pass)
    {
        (object?[] sources, object?[] liveTax, object?[] amounts) = ReadProvisionColumns(provisions, timer, pass);

        using IDisposable scope = timer.Measure(Pass("write provision amounts", pass));

        int updated = 0;
        int runStart = -1;

        // One extra step past the end so that a run reaching the last row is written.
        for (int row = 0; row <= sources.Length; row++)
        {
            bool isModelled = row < sources.Length
                && string.Equals(sources[row] as string, modelledSource, StringComparison.Ordinal);

            if (isModelled)
            {
                if (liveTax[row] is not double)
                {
                    throw new InvalidOperationException("Row " + Format(row + 1) + " of table " + TaxProvisionTable
                        + " has a non numeric " + TotalTaxLiveColumn + " value.");
                }

                if (runStart < 0)
                {
                    runStart = row;
                }

                if (!Equals(liveTax[row], amounts[row]))
                {
                    updated++;
                }
            }
            else if (runStart >= 0)
            {
                ExcelUtils.WriteColumnBlock(provisions, ProvisionAmountColumn, runStart, liveTax[runStart..row]);
                runStart = -1;
            }
        }

        return updated;
    }

    /// <summary>
    /// Reads the columns of the TaxProvision table that a pass needs, measuring how long it takes.
    /// </summary>
    /// <param name="provisions">The TaxProvision table.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="pass">The one based number of the pass being run.  It labels the measurement.</param>
    /// <returns>The ProvisionSource, TotalTaxLive and ProvisionAmount columns, in row order.</returns>
    private static (object?[] Sources, object?[] LiveTax, object?[] Amounts) ReadProvisionColumns(
        Excel.ListObject provisions, DiagnosticTimer timer, int pass)
    {
        using IDisposable scope = timer.Measure(Pass("read provision columns", pass));

        object?[] sources = ExcelUtils.ReadColumn(provisions, ProvisionSourceColumn);
        object?[] liveTax = ExcelUtils.ReadColumn(provisions, TotalTaxLiveColumn);
        object?[] amounts = ExcelUtils.ReadColumn(provisions, ProvisionAmountColumn);

        if (sources.Length != liveTax.Length || sources.Length != amounts.Length)
        {
            throw new InvalidOperationException("The columns of table " + TaxProvisionTable
                + " do not have a consistent number of rows.");
        }

        return (sources, liveTax, amounts);
    }

    /// <summary>
    /// Reads the result of a check from the 90_Checks worksheet.
    /// </summary>
    /// <param name="workbook">The workbook to read.</param>
    /// <param name="checkId">The id of the check, ex. K39.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="pass">The one based number of the pass being run.  It labels the measurement.</param>
    /// <returns>The check result, ex. OK or FAIL.</returns>
    /// <remarks>
    /// The check results occupy a single row table whose headers are the check ids, so a check is located by its id
    /// rather than by its address.  K39 currently resolves to 90_Checks!AM5.
    /// </remarks>
    private static string ReadCheckResult(Excel.Workbook workbook, string checkId, DiagnosticTimer timer, int pass)
    {
        using IDisposable scope = timer.Measure(Pass("read check " + checkId, pass));

        Excel.ListObject checks = ExcelUtils.GetListObject(workbook, ChecksWorksheet, ChecksTable);
        try
        {
            object?[] values = ExcelUtils.ReadColumn(checks, checkId);

            if (values.Length != 1)
            {
                throw new InvalidOperationException("Check " + checkId + " on worksheet " + ChecksWorksheet
                    + " resolved to " + Format(values.Length) + " rows rather than one.");
            }

            return values[0] as string ?? string.Empty;
        }
        finally
        {
            ExcelUtils.ReleaseComObject(checks);
        }
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
    /// Names a measurement that belongs to one pass of the iteration loop.
    /// </summary>
    /// <param name="phase">The name of the phase.</param>
    /// <param name="pass">The one based number of the pass.</param>
    /// <returns>The phase name qualified by the pass number.</returns>
    private static string Pass(string phase, int pass)
    {
        return phase + " (pass " + Format(pass) + ")";
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
}
