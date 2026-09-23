using System.Globalization;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pfm;

/// <summary>
/// The worksheets, tables and defined names of the analysis template, which is the contract between the template and
/// this tool.
/// </summary>
/// <remarks>
/// The template is authored by hand and is the only place the analysis is expressed; this tool writes two worksheets
/// into a copy of it and reads a handful of derived values back to check what it wrote.  Everything it needs to name is
/// named here, so what the tool knows about the template is one screen long.  See the AnalysisSpec worksheet of the
/// template, which is the design record the names below come from.
/// </remarks>
public static class AnalysisTemplate
{
    /// <summary>
    /// The worksheet the configuration of the sweep is written to.
    /// </summary>
    public const string RunConfigurationWorksheet = "RunConfiguration";

    /// <summary>
    /// The worksheet the results of the sweep are written to, and the table defined over them.  The table is the seam:
    /// every formula of the analysis reads it by name and none of them names the worksheet.
    /// </summary>
    public const string SimDataWorksheet = "SimData";

    /// <summary>
    /// The table defined over the results, which is what the analysis reads.
    /// </summary>
    public const string SimDataTable = "SimData";

    /// <summary>
    /// The worksheet holding the intermediate calculations and the chart plumbing of the analysis.
    /// </summary>
    public const string CalculationWorksheet = "SimAnalysisCalc";

    /// <summary>
    /// The worksheet holding the analysis itself: the tables a reader reads and the charts drawn over them.
    /// </summary>
    public const string AnalysisWorksheet = "SimAnalysis";

    /// <summary>
    /// The worksheet holding the design record of the analysis.  It is authored in the template and left out of the
    /// output, which carries no formulas for it to specify.
    /// </summary>
    public const string SpecificationWorksheet = "AnalysisSpec";

    /// <summary>
    /// The worksheets the output workbook is required to hold.
    /// </summary>
    /// <remarks>
    /// These are required rather than exhaustive: a presentation worksheet added to the template is flattened along
    /// with the rest and carried into the output, so growing the analysis by a sheet is a change to the template
    /// alone.  What is not optional is these four, because the transplant writes two of them and the checks read the
    /// other two.
    /// </remarks>
    public static readonly IReadOnlyList<string> RequiredWorksheets =
        [RunConfigurationWorksheet, SimDataWorksheet, CalculationWorksheet, AnalysisWorksheet];

    /// <summary>
    /// The worksheets that hold no formula and are therefore never flattened.
    /// </summary>
    /// <remarks>
    /// These two are the sweep as it was written, and the tables defined over them would not survive the paste that
    /// flattening is.  Every other worksheet of the template holds formulas and is flattened.
    /// </remarks>
    public static readonly IReadOnlyList<string> DataWorksheets = [RunConfigurationWorksheet, SimDataWorksheet];

    /// <summary>
    /// Indicates whether what a defined name refers to is anchored to a spill.
    /// </summary>
    /// <param name="refersTo">What the name refers to, as Excel states it.</param>
    /// <returns>True when it is.</returns>
    /// <remarks>
    /// Excel states a spill anchored name either way, ex. =SimAnalysisCalc!$B$2# through the object model and
    /// ANCHORARRAY in the file it saves, so both are recognized.  A name anchored to a spill resizes with the sweep,
    /// which is why the flatten drops it and why the checks do not measure the template against it.
    /// </remarks>
    public static bool IsSpillAnchored(string refersTo)
    {
        ArgumentNullException.ThrowIfNull(refersTo);

        return refersTo.EndsWith('#') || refersTo.Contains("ANCHORARRAY", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Orders the worksheets that hold formulas as they are flattened to values.
    /// </summary>
    /// <param name="worksheetNames">The worksheets of the workbook, in the order it holds them.</param>
    /// <returns>The worksheets to flatten, in the order they are flattened.</returns>
    /// <remarks>
    /// The calculation worksheet is flattened last, because it is the one every other sheet reads: the names the
    /// charts and tables of the analysis are built on are anchored to the spills it carries, and a name anchored to a
    /// spill refers to nothing once that spill is a block of values.  Flattened the other way round, every figure that
    /// reads one of those names is saved as #REF!.  The data worksheets are left alone, and so is the design record,
    /// which is removed rather than flattened.
    /// </remarks>
    public static IReadOnlyList<string> FlattenedWorksheetsOf(IReadOnlyList<string> worksheetNames)
    {
        ArgumentNullException.ThrowIfNull(worksheetNames);

        var flattened = worksheetNames
            .Where(name => !DataWorksheets.Contains(name, StringComparer.Ordinal)
                && !string.Equals(name, SpecificationWorksheet, StringComparison.Ordinal)
                && !string.Equals(name, CalculationWorksheet, StringComparison.Ordinal))
            .ToList();

        flattened.Add(CalculationWorksheet);

        return flattened;
    }

    /// <summary>
    /// The worksheet names a sweep's results may be found under in the workbook they are read from.  A sweep coalesced
    /// today writes one worksheet name whatever simulation produced it; the others are what earlier sweeps named theirs
    /// after the simulation that ran.
    /// </summary>
    public static readonly IReadOnlyList<string> SourceResultsWorksheets =
        [SimDataWorksheet, "MonteCarloData", "BackTestData"];

    /// <summary>
    /// The defined name holding the version of the analysis, which is authored in the template and stamped into the
    /// output.
    /// </summary>
    public const string TemplateVersionName = "AnalysisTemplateVersion";

    /// <summary>
    /// The defined name of the cell the provenance of the output is stamped into.
    /// </summary>
    public const string ProvenanceName = "AnalysisProvenance";

    /// <summary>
    /// The defined name of the derived count of simulations the analysis is reading.
    /// </summary>
    public const string SimulationCountName = "SimulationCount";

    /// <summary>
    /// The defined name of the derived count of year columns in one metric block.
    /// </summary>
    public const string ProjectionYearCountName = "ProjectionYearCount";

    /// <summary>
    /// The defined name of the derived first projection year, which the analysis parses off a metric header.
    /// </summary>
    public const string FirstProjectionYearName = "FirstProjectionYear";

    /// <summary>
    /// The defined name of the derived count of leading identity columns of the results.
    /// </summary>
    public const string IdentityColumnCountName = "IdentityColumnCount";

    /// <summary>
    /// The defined name of the count of simulations whose income fell below the floor, which is the one number of the
    /// analysis that is reconciled against the results it was computed from.
    /// </summary>
    public const string BelowFloorCountName = "SimulationsBelowFloor";

    /// <summary>
    /// What a defined name holds to mark a block of individual trials, which are the only cells of the analysis
    /// allowed to hold an error: a trial slot with no trial to plot returns #N/A so that the chart draws no line.
    /// </summary>
    /// <remarks>
    /// The blocks are recognized by their names rather than listed here, so a trial block added to the template needs
    /// no change to this tool.  ThrottleTraceSingleYear and ThrottleTraceWindow are what the template names today.
    /// </remarks>
    public const string TrialBlockNameMarker = "Trace";

    /// <summary>
    /// The suffix a defined name carries when it holds one metric key, or a column of them.
    /// </summary>
    public const string MetricKeyNameSuffix = "MetricKey";

    /// <summary>
    /// The suffix a defined name carries when it is a year header row.
    /// </summary>
    public const string YearRowNameSuffix = "Years";

    /// <summary>
    /// The defined name the template may use to state the metric carrying the shortfall of one simulated year against
    /// the income floor, which the guardrail block's count of failed simulations is reconciled against.
    /// </summary>
    /// <remarks>
    /// A template that defines this name states the metric itself; one that does not is read as carrying
    /// <see cref="DefaultShortfallMetricKey"/>, which is what the template names today.  The reconciliation is the one
    /// check that reads the results rather than what the analysis says about them, so it has to name a metric; where
    /// that name comes from is the template's business.
    /// </remarks>
    public const string ShortfallMetricKeyName = "ShortfallMetricKey";

    /// <summary>
    /// The metric carrying the shortfall of one simulated year against the income floor, when the template does not
    /// name one through <see cref="ShortfallMetricKeyName"/>.
    /// </summary>
    public const string DefaultShortfallMetricKey = "IncomeShortfall";

    /// <summary>
    /// The greatest number of leading identity columns the analysis has room for.
    /// </summary>
    /// <remarks>
    /// The per-simulation diagnostics on the calculation sheet stack the identity columns of a simulation beside the
    /// score being ranked, so each of those blocks is one column wider than the identity block of the sweep: a back
    /// test identifies a simulation by three columns where a Monte Carlo campaign uses two, and every one of those
    /// blocks is a column wider for it.  The narrowest allowance on that sheet is six columns, and a sweep that passes
    /// it spills into the block beside it, which Excel reports as #SPILL! on one cell and not as anything about the
    /// sweep.  Checking it here is what turns that into a sentence.
    ///
    /// This is the one capacity of the template that is stated here rather than measured from it: the gap between two
    /// spills on the calculation sheet is not something the workbook states anywhere a reader could find, so
    /// rearranging that sheet is the one template change that needs a change here as well.  The room for year columns
    /// is measured off the year header rows instead; see <see cref="AnalysisChecks"/>.
    /// </remarks>
    public const int MaxIdentityColumns = 5;
}

/// <summary>
/// Applies the analysis of a template workbook to one sweep, by writing the sweep's configuration and results into a
/// copy of the template and flattening what the template computes from them.
/// </summary>
/// <remarks>
/// The template is the base workbook and the data is written into it, never the other way around: worksheets carried
/// between workbooks bring their defined names with them as external references, which resolve silently against
/// whatever copy of the source is open and produce plausible numbers from the wrong sweep.  Nothing here writes a
/// formula, a chart or a defined name of the analysis, so the whole surface this tool can disturb is the two
/// worksheets it writes and the one cell it stamps.
/// </remarks>
public sealed class AnalysisTransplant
{
    /// <summary>
    /// The number of result rows copied per pair of calls.  A band is two calls whatever its size, so this trades the
    /// memory a band occupies against the number of calls a large sweep costs.
    /// </summary>
    private const int CopyBlockRows = 2000;

    private readonly ExcelSession _session;
    private readonly DiagnosticTimer _timer;
    private readonly TextWriter _output;

    /// <summary>
    /// Initializes a new instance of the <see cref="AnalysisTransplant"/> class.
    /// </summary>
    /// <param name="session">The session that owns the copy of the template being written.</param>
    /// <param name="timer">The timer that measures the phases of the operation.</param>
    /// <param name="output">The writer for normal output.</param>
    public AnalysisTransplant(ExcelSession session, DiagnosticTimer timer, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(session);
        ArgumentNullException.ThrowIfNull(timer);
        ArgumentNullException.ThrowIfNull(output);

        _session = session;
        _timer = timer;
        _output = output;
    }

    /// <summary>
    /// Transplants one sweep into the copy of the template, recalculates it, checks what it computed, stamps the
    /// provenance of the run and flattens every sheet to values.
    /// </summary>
    /// <param name="sourcePath">The path of the workbook holding the sweep's configuration and results.</param>
    /// <param name="runId">The identifier of the sweep, which the output is named and stamped with.</param>
    /// <returns>What was applied, for the caller to report.</returns>
    /// <exception cref="InvalidOperationException">
    /// The source workbook is not one sweep's data, the template is not one this tool can drive, or the analysis the
    /// template computed does not hold up to the checks of the AnalysisSpec worksheet.
    /// </exception>
    public AnalysisOutcome Apply(string sourcePath, string runId)
    {
        ArgumentNullException.ThrowIfNull(sourcePath);
        ArgumentNullException.ThrowIfNull(runId);

        string version = ReadTemplateVersion();

        TransplantedData data = Transplant(sourcePath);

        using (IDisposable scope = _timer.Measure("recalculate analysis"))
        {
            // Every formula, not only the ones Excel marked dirty: the spills that the analysis is built on resize
            // against the table that has just been redefined, and their dependents have to settle before anything is
            // read back.
            _session.CalculateFull();
        }

        AnalysisChecks checks;
        using (IDisposable scope = _timer.Measure("check analysis"))
        {
            checks = AnalysisChecks.Verify(_session.Workbook, data.Header, data.RowCount);
        }

        string provenance = Stamp(version, runId, sourcePath);

        using (IDisposable scope = _timer.Measure("flatten to values"))
        {
            Flatten();
        }

        Log.Logger.Information("PFM_APPLY_ANALYSIS_APPLIED: runId=" + runId + " templateVersion=" + version
            + " simulations=" + Format(data.RowCount) + " years=" + Format(checks.ProjectionYearCount)
            + " source=" + data.ResultsWorksheet);

        return new AnalysisOutcome(version, provenance, data.RowCount, data.ColumnCount, checks);
    }


    /// <summary>
    /// Reads the version of the analysis the template carries.
    /// </summary>
    /// <returns>The version, as the template states it.</returns>
    /// <remarks>
    /// This is the first thing read, because a workbook that does not carry it is not a template this tool can apply,
    /// and finding that out before anything has been written costs nothing.
    /// </remarks>
    private string ReadTemplateVersion()
    {
        object? value = ExcelUtils.EvaluateName(_session.Workbook, AnalysisTemplate.TemplateVersionName);
        string version = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;

        if (string.IsNullOrWhiteSpace(version))
        {
            throw new InvalidOperationException("The template does not state a version in "
                + AnalysisTemplate.TemplateVersionName + ".  A workbook without one is not an analysis template: the "
                + "version is what tells two applications of one sweep apart.");
        }

        return version;
    }

    /// <summary>
    /// Writes the sweep's configuration and results into the copy of the template.
    /// </summary>
    /// <param name="sourcePath">The path of the workbook holding them.</param>
    /// <returns>What was written.</returns>
    /// <remarks>
    /// The source workbook is opened beside the template in the same Excel instance, read, and closed again before
    /// anything is recalculated: a full calculation recalculates every workbook the instance holds open.
    /// </remarks>
    private TransplantedData Transplant(string sourcePath)
    {
        using IDisposable scope = _timer.Measure("transplant sweep");

        Excel.Workbook? source = null;
        try
        {
            source = _session.OpenCompanionWorkbook(sourcePath, readOnly: true);

            WriteConfiguration(source);

            return WriteResults(source);
        }
        finally
        {
            ExcelUtils.CloseWorkbook(source);
        }
    }

    /// <summary>
    /// Reads the sweep's configuration out of the source workbook and writes it into the template's own configuration
    /// worksheet.
    /// </summary>
    /// <param name="source">The workbook holding the sweep.</param>
    private void WriteConfiguration(Excel.Workbook source)
    {
        SweepConfiguration configuration;

        Excel.Worksheet sourceSheet = ExcelUtils.GetWorksheet(source, AnalysisTemplate.RunConfigurationWorksheet);
        try
        {
            configuration = RunConfigurationSheet.Read(sourceSheet);
        }
        finally
        {
            ExcelUtils.ReleaseComObject(sourceSheet);
        }

        Excel.Worksheet target = ExcelUtils.GetWorksheet(_session.Workbook,
            AnalysisTemplate.RunConfigurationWorksheet);
        try
        {
            RunConfigurationSheet.Update(target, configuration.Synopsis, configuration.Tables);
        }
        finally
        {
            ExcelUtils.ReleaseComObject(target);
        }

        _output.WriteLine("Wrote " + Format(configuration.Synopsis.Count) + " lines of run configuration and "
            + Format(configuration.Tables.Count) + " configuration tables to "
            + AnalysisTemplate.RunConfigurationWorksheet + ".");
    }

    /// <summary>
    /// Copies the sweep's results into the template's results worksheet and defines the table the analysis reads over
    /// them.
    /// </summary>
    /// <param name="source">The workbook holding the sweep.</param>
    /// <returns>What was written.</returns>
    /// <remarks>
    /// The table is resized to cover this sweep rather than replaced, because resizing is what leaves the analysis
    /// reading it: a table that is dissolved has every formula that read it structurally rewritten into the cell range
    /// it happened to occupy, so the next sweep would be read through the last one's geometry with nothing reported.
    /// The rows and columns the previous sweep left beyond this one's are cleared afterwards, so that nothing outside
    /// the table reads as data.
    /// </remarks>
    private TransplantedData WriteResults(Excel.Workbook source)
    {
        Excel.Worksheet sourceSheet = FindResultsWorksheet(source);
        Excel.Worksheet? target = null;
        Excel.ListObject? table = null;
        try
        {
            string sourceName = sourceSheet.Name;

            target = ExcelUtils.GetWorksheet(_session.Workbook, AnalysisTemplate.SimDataWorksheet);
            table = ExcelUtils.GetTable(target, AnalysisTemplate.SimDataTable);

            (int firstRow, int firstColumn, int lastRow, int lastColumn) = ExcelUtils.GetUsedExtent(sourceSheet);

            if (firstRow != 1 || firstColumn != 1)
            {
                throw new InvalidOperationException("The results on worksheet " + sourceName + " start at row "
                    + Format(firstRow) + " column " + Format(firstColumn) + " rather than at A1, so they do not line "
                    + "up with the heading row of the table the analysis reads.");
            }

            if (lastRow < 2)
            {
                throw new InvalidOperationException("The results on worksheet " + sourceName + " are a heading row "
                    + "with no simulations beneath it.");
            }

            // The table is resized before the values are written, so that every row of them is written into the table
            // rather than beside it, and the rows and columns of the last sweep that this one does not cover are
            // cleared rather than left to read as data.
            ExcelUtils.ResizeTable(table, 1, 1, lastRow, lastColumn);
            ExcelUtils.ClearBeyond(target, lastRow, lastColumn);
            ExcelUtils.CopyValues(sourceSheet, target, lastRow, lastColumn, CopyBlockRows);

            object?[,] headerRow = ExcelUtils.ReadGrid(target, 1, 1, 1, lastColumn);
            var header = new string[lastColumn];
            for (int column = 0; column < lastColumn; column++)
            {
                header[column] = Convert.ToString(headerRow[0, column], CultureInfo.InvariantCulture) ?? string.Empty;
            }

            _output.WriteLine("Copied " + Format(lastRow - 1) + " simulations of " + Format(lastColumn)
                + " columns from " + sourceName + " to " + AnalysisTemplate.SimDataWorksheet + ", and resized "
                + AnalysisTemplate.SimDataTable + " over them.");

            return new TransplantedData(sourceName, header, lastRow - 1, lastColumn);
        }
        finally
        {
            ExcelUtils.ReleaseComObject(table);
            ExcelUtils.ReleaseComObject(target);
            ExcelUtils.ReleaseComObject(sourceSheet);
        }
    }

    /// <summary>
    /// Finds the worksheet the sweep's results are on.
    /// </summary>
    /// <param name="source">The workbook holding the sweep.</param>
    /// <returns>The worksheet.</returns>
    private static Excel.Worksheet FindResultsWorksheet(Excel.Workbook source)
    {
        foreach (string name in AnalysisTemplate.SourceResultsWorksheets)
        {
            Excel.Worksheet? worksheet = ExcelUtils.FindWorksheet(source, name);
            if (worksheet != null)
            {
                return worksheet;
            }
        }

        throw new InvalidOperationException("The workbook holds no worksheet of results: it has none of "
            + string.Join(", ", AnalysisTemplate.SourceResultsWorksheets) + ".");
    }

    /// <summary>
    /// Stamps the provenance of this run into the cell the analysis keeps for it.
    /// </summary>
    /// <param name="version">The version of the analysis the template carries.</param>
    /// <param name="runId">The identifier of the sweep.</param>
    /// <param name="sourcePath">The path of the workbook the sweep was read from.</param>
    /// <returns>The line that was stamped.</returns>
    /// <remarks>
    /// The output carries no formulas and no design record, so this line is the whole of what it says about itself.
    /// Two outputs of one sweep under different versions of the analysis are otherwise indistinguishable.
    /// </remarks>
    private string Stamp(string version, string runId, string sourcePath)
    {
        string provenance = "Analysis template " + version + " | run " + runId + " | from "
            + Path.GetFileName(sourcePath) + " | applied "
            + DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

        ExcelUtils.SetNameValue(_session.Workbook, AnalysisTemplate.ProvenanceName, provenance);

        return provenance;
    }

    /// <summary>
    /// Replaces every formula of the output with the value it calculated, drops the names that only mean anything while
    /// those formulas are live, and removes the design record of the analysis.
    /// </summary>
    /// <remarks>
    /// What survives is what a reader of the output needs: the values, the number formats, the conditional formatting,
    /// the tables the results can still be filtered through, and the charts, which need no attention because every
    /// series reads a range that still holds the same values.  What does not is the spill anchored names, which without
    /// a live spill behind them evaluate to #REF!, and the version name, which refers to the sheet being removed.
    /// </remarks>
    private void Flatten()
    {
        IReadOnlyList<string> flattened =
            AnalysisTemplate.FlattenedWorksheetsOf(ExcelUtils.GetWorksheetNames(_session.Workbook));

        foreach (string worksheetName in flattened)
        {
            Excel.Worksheet worksheet = ExcelUtils.GetWorksheet(_session.Workbook, worksheetName);
            try
            {
                ExcelUtils.FlattenToValues(worksheet);
            }
            finally
            {
                ExcelUtils.ReleaseComObject(worksheet);
            }
        }

        int dropped = 0;
        foreach ((string name, string refersTo) in ExcelUtils.GetNameDefinitions(_session.Workbook))
        {
            bool spillAnchored = AnalysisTemplate.IsSpillAnchored(refersTo);
            bool onSpecification = refersTo.Contains(AnalysisTemplate.SpecificationWorksheet + "!",
                StringComparison.OrdinalIgnoreCase);

            if (name.StartsWith('_') || (!spillAnchored && !onSpecification))
            {
                continue;
            }

            ExcelUtils.DeleteName(_session.Workbook, name);
            dropped++;
        }

        ExcelUtils.DeleteWorksheet(_session.Workbook, AnalysisTemplate.SpecificationWorksheet);

        CheckOutputWorksheets();

        _output.WriteLine("Flattened " + string.Join(" and ", flattened)
            + " to values, dropped " + Format(dropped) + " defined names that a flattened workbook cannot hold, and "
            + "removed " + AnalysisTemplate.SpecificationWorksheet + ".");
    }

    /// <summary>
    /// Checks that what is left is the workbook the output is meant to be: the worksheets the analysis is made of, and
    /// no design record.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// The workbook is missing a worksheet the output is required to hold, or still holds the design record.
    /// </exception>
    /// <remarks>
    /// A worksheet the template carries beyond the required four is flattened along with the rest and carried into the
    /// output, so this states what has to be there rather than what may be: growing the analysis by a presentation
    /// sheet is a change to the template alone.  What is checked is that the flatten reached everything, because a
    /// sheet that kept its formulas would ship them in a workbook of values, against a design record that has been
    /// removed and names that have been dropped, which is to say as a block of #REF!.
    /// </remarks>
    private void CheckOutputWorksheets()
    {
        IReadOnlyList<string> found = ExcelUtils.GetWorksheetNames(_session.Workbook);

        var missing = AnalysisTemplate.RequiredWorksheets.Except(found, StringComparer.Ordinal).ToList();
        bool specificationRemains = found.Contains(AnalysisTemplate.SpecificationWorksheet, StringComparer.Ordinal);

        if (missing.Count == 0 && !specificationRemains)
        {
            return;
        }

        throw new InvalidOperationException("The analysis of this sweep holds "
            + (missing.Count > 0 ? "no worksheet " + string.Join(", ", missing) : string.Empty)
            + (missing.Count > 0 && specificationRemains ? " and " : string.Empty)
            + (specificationRemains ? "the design record " + AnalysisTemplate.SpecificationWorksheet + " still"
                : string.Empty)
            + ".  The output is required to hold " + string.Join(", ", AnalysisTemplate.RequiredWorksheets)
            + ", all values, and to carry no design record.");
    }

    /// <summary>
    /// Formats a count for a message.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <returns>The formatted value.</returns>
    private static string Format(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }
}

/// <summary>
/// What a transplant wrote into the copy of the template.
/// </summary>
internal sealed class TransplantedData
{
    /// <summary>
    /// Initializes a new instance of the <see cref="TransplantedData"/> class.
    /// </summary>
    /// <param name="resultsWorksheet">The name of the worksheet the results were read from.</param>
    /// <param name="header">The column headings of the results, in column order.</param>
    /// <param name="rowCount">The number of simulations written.</param>
    /// <param name="columnCount">The number of columns written.</param>
    internal TransplantedData(string resultsWorksheet, IReadOnlyList<string> header, int rowCount, int columnCount)
    {
        ResultsWorksheet = resultsWorksheet;
        Header = header;
        RowCount = rowCount;
        ColumnCount = columnCount;
    }

    /// <summary>
    /// Gets the name of the worksheet the results were read from.
    /// </summary>
    internal string ResultsWorksheet { get; }

    /// <summary>
    /// Gets the column headings of the results, in column order.
    /// </summary>
    internal IReadOnlyList<string> Header { get; }

    /// <summary>
    /// Gets the number of simulations written, not counting the heading row.
    /// </summary>
    internal int RowCount { get; }

    /// <summary>
    /// Gets the number of columns written.
    /// </summary>
    internal int ColumnCount { get; }
}

/// <summary>
/// What one application of the analysis produced, for the command that ran it to report.
/// </summary>
public sealed class AnalysisOutcome
{
    /// <summary>
    /// Initializes a new instance of the <see cref="AnalysisOutcome"/> class.
    /// </summary>
    /// <param name="templateVersion">The version of the analysis that was applied.</param>
    /// <param name="provenance">The line stamped into the output.</param>
    /// <param name="simulationCount">The number of simulations the analysis was applied to.</param>
    /// <param name="columnCount">The number of result columns.</param>
    /// <param name="checks">What the checks on the applied analysis found.</param>
    internal AnalysisOutcome(string templateVersion, string provenance, int simulationCount, int columnCount,
        AnalysisChecks checks)
    {
        TemplateVersion = templateVersion;
        Provenance = provenance;
        SimulationCount = simulationCount;
        ColumnCount = columnCount;
        Checks = checks;
    }

    /// <summary>
    /// Gets the version of the analysis that was applied.
    /// </summary>
    public string TemplateVersion { get; }

    /// <summary>
    /// Gets the line stamped into the output.
    /// </summary>
    public string Provenance { get; }

    /// <summary>
    /// Gets the number of simulations the analysis was applied to.
    /// </summary>
    public int SimulationCount { get; }

    /// <summary>
    /// Gets the number of result columns.
    /// </summary>
    public int ColumnCount { get; }

    /// <summary>
    /// Gets what the checks on the applied analysis found.
    /// </summary>
    public AnalysisChecks Checks { get; }
}
