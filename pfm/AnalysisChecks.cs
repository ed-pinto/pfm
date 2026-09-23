using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pfm;

/// <summary>
/// The checks an applied analysis is held to before it is flattened and saved, and what they found.
/// </summary>
/// <remarks>
/// The template computes the analysis; this is what establishes that it computed it over the sweep it was given rather
/// than producing something that merely reads plausibly.  The checks are the validation list of the AnalysisSpec
/// worksheet, and the one that matters most is the first: every metric the analysis reads by key has to be a block of
/// columns the results actually carry, so a harvest that renamed or dropped a metric is reported by name instead of
/// yielding a workbook of blanks.  Every check is run before any of them is reported, because a caller fixing a
/// template or a harvest wants the whole list rather than the first line of it.
/// </remarks>
public sealed class AnalysisChecks
{
    /// <summary>
    /// The greatest shortfall against the income floor, in dollars, that is still read as no shortfall.  It brackets
    /// the tolerance the analysis applies rather than restating it: the count the analysis reports has to lie between
    /// the count of simulations with any shortfall at all and the count of those whose shortfall is beyond argument.
    /// </summary>
    private const double ShortfallTolerance = 0.005;

    /// <summary>
    /// The number of result rows read per call when the shortfall metric is reconciled.
    /// </summary>
    private const int ReadBlockRows = 2000;

    /// <summary>
    /// Initializes a new instance of the <see cref="AnalysisChecks"/> class.
    /// </summary>
    /// <param name="projectionYearCount">The number of year columns in one metric block.</param>
    /// <param name="firstProjectionYear">The first projected year.</param>
    /// <param name="identityColumnCount">The number of leading identity columns of the results.</param>
    /// <param name="belowFloorCount">The number of simulations whose income fell below the floor.</param>
    /// <param name="blankTrialSlots">The number of cells holding the #N/A of a trial slot with no trial.</param>
    private AnalysisChecks(int projectionYearCount, int firstProjectionYear, int identityColumnCount,
        int belowFloorCount, int blankTrialSlots)
    {
        ProjectionYearCount = projectionYearCount;
        FirstProjectionYear = firstProjectionYear;
        IdentityColumnCount = identityColumnCount;
        BelowFloorCount = belowFloorCount;
        BlankTrialSlots = blankTrialSlots;
    }

    /// <summary>
    /// Gets the number of year columns in one metric block, as the analysis derived it.
    /// </summary>
    public int ProjectionYearCount { get; }

    /// <summary>
    /// Gets the first projected year, as the analysis parsed it off a metric heading.
    /// </summary>
    public int FirstProjectionYear { get; }

    /// <summary>
    /// Gets the number of leading identity columns of the results, as the analysis derived it.
    /// </summary>
    public int IdentityColumnCount { get; }

    /// <summary>
    /// Gets the number of simulations whose income fell below the floor, as the analysis counted them.
    /// </summary>
    public int BelowFloorCount { get; }

    /// <summary>
    /// Gets the number of cells holding the #N/A of a trial slot with no trial to plot.
    /// </summary>
    public int BlankTrialSlots { get; }

    /// <summary>
    /// Checks an applied analysis and reports what it found.
    /// </summary>
    /// <param name="workbook">The workbook holding the applied analysis, recalculated and not yet flattened.</param>
    /// <param name="header">The column headings of the results that were written.</param>
    /// <param name="rowCount">The number of simulations that were written.</param>
    /// <returns>What the checks found.</returns>
    /// <exception cref="InvalidOperationException">Any check failed.  The message states every failure.</exception>
    public static AnalysisChecks Verify(Excel.Workbook workbook, IReadOnlyList<string> header, int rowCount)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(header);

        var failures = new List<string>();

        MetricBlocks blocks = MetricBlocks.Read(header, ReadShortfallMetricKey(workbook), failures);

        int yearCount = (int)ExcelUtils.GetNameNumber(workbook, AnalysisTemplate.ProjectionYearCountName);
        int firstYear = (int)ExcelUtils.GetNameNumber(workbook, AnalysisTemplate.FirstProjectionYearName);
        int identityColumns = (int)ExcelUtils.GetNameNumber(workbook, AnalysisTemplate.IdentityColumnCountName);
        int simulations = (int)ExcelUtils.GetNameNumber(workbook, AnalysisTemplate.SimulationCountName);

        IReadOnlyList<string> worksheets =
            AnalysisTemplate.FlattenedWorksheetsOf(ExcelUtils.GetWorksheetNames(workbook));

        CheckDerivedCounts(blocks, rowCount, simulations, yearCount, firstYear, identityColumns,
            ReadYearColumnCapacity(workbook), failures);
        CheckMetricKeys(workbook, blocks, yearCount, firstYear, failures);
        CheckYearHeaderRows(workbook, yearCount, failures);
        CheckNames(workbook, failures);

        int blankTrialSlots = CheckCellContents(workbook, worksheets, failures);
        CheckChartClearance(workbook, worksheets, failures);

        int belowFloor = CheckShortfallReconciliation(workbook, blocks, rowCount, failures);

        if (failures.Count > 0)
        {
            string message = "The analysis the template computed did not hold up to "
                + (failures.Count == 1 ? "a check" : Format(failures.Count) + " checks") + ":"
                + Environment.NewLine + string.Join(Environment.NewLine, failures.Select(failure => "  " + failure));

            Log.Logger.Error("PFM_APPLY_ANALYSIS_CHECK_FAILED: " + string.Join(" | ", failures));

            throw new InvalidOperationException(message);
        }

        return new AnalysisChecks(yearCount, firstYear, identityColumns, belowFloor, blankTrialSlots);
    }

    /// <summary>
    /// Checks the counts the analysis derives against the results that were written.
    /// </summary>
    /// <param name="blocks">The metric blocks the headings describe.</param>
    /// <param name="rowCount">The number of simulations that were written.</param>
    /// <param name="simulations">The count of simulations the analysis derived.</param>
    /// <param name="yearCount">The count of year columns the analysis derived.</param>
    /// <param name="firstYear">The first projected year the analysis derived.</param>
    /// <param name="identityColumns">The count of identity columns the analysis derived.</param>
    /// <param name="yearColumnCapacity">
    /// The number of year columns the template has room for, or zero when it states none.
    /// </param>
    /// <param name="failures">The list every failure is added to.</param>
    private static void CheckDerivedCounts(MetricBlocks blocks, int rowCount, int simulations, int yearCount,
        int firstYear, int identityColumns, int yearColumnCapacity, List<string> failures)
    {
        if (simulations != rowCount)
        {
            failures.Add(AnalysisTemplate.SimulationCountName + " is " + Format(simulations) + " and "
                + Format(rowCount) + " simulations were written, so the table does not cover the rows.");
        }

        if (blocks.Years.Count > 0 && yearCount != blocks.Years.Count)
        {
            failures.Add(AnalysisTemplate.ProjectionYearCountName + " is " + Format(yearCount) + " and the results "
                + "carry " + Format(blocks.Years.Count) + " projected years.");
        }

        if (blocks.Years.Count > 0 && firstYear != blocks.Years[0])
        {
            failures.Add(AnalysisTemplate.FirstProjectionYearName + " is " + Format(firstYear) + " and the first "
                + "projected year of the results is " + Format(blocks.Years[0]) + ".");
        }

        if (yearColumnCapacity > 0 && yearCount > yearColumnCapacity)
        {
            failures.Add("The sweep projects " + Format(yearCount) + " years and the year header rows of the analysis "
                + "are " + Format(yearColumnCapacity) + " columns wide.  A longer horizon is a change to the template "
                + "rather than a dataset it can be applied to.");
        }

        if (blocks.IdentityColumnCount >= 0 && identityColumns != blocks.IdentityColumnCount)
        {
            failures.Add(AnalysisTemplate.IdentityColumnCountName + " is " + Format(identityColumns) + " and the "
                + "results carry " + Format(blocks.IdentityColumnCount) + " columns before the first metric block.");
        }

        if (identityColumns > AnalysisTemplate.MaxIdentityColumns)
        {
            failures.Add("The results identify a simulation by " + Format(identityColumns) + " columns and the "
                + "analysis has room for " + Format(AnalysisTemplate.MaxIdentityColumns) + ".  Each per-simulation "
                + "diagnostic stacks those columns beside the score it ranks, so a wider identity block reaches into "
                + "the block beside it; making room for it is a change to the template rather than a dataset it can be "
                + "applied to.");
        }
    }

    /// <summary>
    /// Reads how many year columns the template has room for, off the year header rows it is laid out on.
    /// </summary>
    /// <param name="workbook">The workbook holding the applied analysis.</param>
    /// <returns>The narrowest year header row, in columns, or zero when the template lays out none.</returns>
    /// <remarks>
    /// The capacity is measured from the template rather than stated here, so widening the blocks of the analysis is a
    /// change to the template alone.  Every year header row is a fixed block of cells that a spill of the projected
    /// years fills, so the narrowest of them is the horizon the analysis can show: a sweep projecting further spills
    /// past the end of that row and into whatever is beside it.  Rows anchored to a spill are left out of the
    /// measurement, because they resize with the sweep and so state its length rather than the template's room for it.
    /// </remarks>
    private static int ReadYearColumnCapacity(Excel.Workbook workbook)
    {
        int capacity = 0;

        foreach ((string name, string refersTo) in ExcelUtils.GetNameDefinitions(workbook))
        {
            if (name.StartsWith('_') || !IsYearRowName(name) || AnalysisTemplate.IsSpillAnchored(refersTo))
            {
                continue;
            }

            Excel.Range? range = null;
            Excel.Range? columns = null;
            try
            {
                range = ExcelUtils.GetNameRange(workbook, name);

                if (range.Count < 2)
                {
                    // A single cell of that name is an input rather than a category axis.
                    continue;
                }

                columns = range.Columns;

                if (capacity == 0 || columns.Count < capacity)
                {
                    capacity = columns.Count;
                }
            }
            finally
            {
                ExcelUtils.ReleaseComObject(columns);
                ExcelUtils.ReleaseComObject(range);
            }
        }

        return capacity;
    }

    /// <summary>
    /// Reads the metric key the shortfall reconciliation is run against.
    /// </summary>
    /// <param name="workbook">The workbook holding the applied analysis.</param>
    /// <returns>The key the template states, or the default when it states none.</returns>
    private static string ReadShortfallMetricKey(Excel.Workbook workbook)
    {
        foreach ((string name, _) in ExcelUtils.GetNameDefinitions(workbook))
        {
            if (!string.Equals(name, AnalysisTemplate.ShortfallMetricKeyName, StringComparison.Ordinal))
            {
                continue;
            }

            foreach (string key in ReadTexts(workbook, name))
            {
                if (key.Length > 0)
                {
                    return key;
                }
            }
        }

        return AnalysisTemplate.DefaultShortfallMetricKey;
    }

    /// <summary>
    /// Indicates whether a defined name holds one metric key, or a column of them.
    /// </summary>
    /// <param name="name">The defined name.</param>
    /// <returns>True when it does.</returns>
    private static bool IsMetricKeyName(string name)
    {
        return name.EndsWith(AnalysisTemplate.MetricKeyNameSuffix, StringComparison.Ordinal)
            || name.EndsWith(AnalysisTemplate.MetricKeyNameSuffix + "s", StringComparison.Ordinal);
    }

    /// <summary>
    /// Indicates whether a defined name is a year header row.
    /// </summary>
    /// <param name="name">The defined name.</param>
    /// <returns>True when it is.</returns>
    private static bool IsYearRowName(string name)
    {
        return name.EndsWith(AnalysisTemplate.YearRowNameSuffix, StringComparison.Ordinal);
    }

    /// <summary>
    /// Checks that every metric the analysis reads by key is a block of columns the results carry.
    /// </summary>
    /// <param name="workbook">The workbook holding the applied analysis.</param>
    /// <param name="blocks">The metric blocks the headings describe.</param>
    /// <param name="yearCount">The count of year columns the analysis derived.</param>
    /// <param name="firstYear">The first projected year the analysis derived.</param>
    /// <param name="failures">The list every failure is added to.</param>
    /// <remarks>
    /// The keys are found through the defined names that hold them rather than by reading a column of the analysis, so
    /// a block added to the template is checked without this tool being told about it.  This is the check that fails on
    /// a dataset the template cannot analyse at all, which is why it names the metric and the year it could not find.
    /// </remarks>
    private static void CheckMetricKeys(Excel.Workbook workbook, MetricBlocks blocks, int yearCount, int firstYear,
        List<string> failures)
    {
        var checkedKeys = new HashSet<string>(StringComparer.Ordinal);

        foreach ((string name, _) in ExcelUtils.GetNameDefinitions(workbook))
        {
            if (name.StartsWith('_') || !IsMetricKeyName(name))
            {
                continue;
            }

            foreach (string key in ReadTexts(workbook, name))
            {
                if (key.Length == 0 || !checkedKeys.Add(key))
                {
                    continue;
                }

                for (int year = firstYear; year < firstYear + yearCount; year++)
                {
                    string column = key + "_" + Format(year);
                    if (!blocks.Headings.Contains(column))
                    {
                        failures.Add("The metric key " + key + " of " + name + " has no column " + column
                            + " in the results.");
                        break;
                    }
                }
            }
        }

        if (checkedKeys.Count == 0)
        {
            failures.Add("The template states no metric keys, so nothing establishes which columns of the results the "
                + "analysis is reading.");
        }
    }

    /// <summary>
    /// Checks that every year header row of the analysis spilled across its block rather than holding one cell.
    /// </summary>
    /// <param name="workbook">The workbook holding the applied analysis.</param>
    /// <param name="yearCount">The count of year columns the analysis derived.</param>
    /// <param name="failures">The list every failure is added to.</param>
    /// <remarks>
    /// A year header row that was formatted as text before its formula was written holds that formula as a literal
    /// string, which spills nothing: every formula keyed off the row then fails while the chart above it still renders.
    /// The rows are found through the defined names that serve the category axes, and the single celled names among
    /// them, ex. the rolling window input, are not category axes and are left alone.
    /// </remarks>
    private static void CheckYearHeaderRows(Excel.Workbook workbook, int yearCount, List<string> failures)
    {
        foreach ((string name, _) in ExcelUtils.GetNameDefinitions(workbook))
        {
            if (name.StartsWith('_') || !IsYearRowName(name))
            {
                continue;
            }

            Excel.Range? range = null;
            try
            {
                range = ExcelUtils.GetNameRange(workbook, name);

                if (range.Count < 2)
                {
                    // A single cell of that name is an input rather than a category axis.
                    continue;
                }

                object?[,] values = ExcelUtils.ReadRange(range);
                int populated = 0;

                for (int row = 0; row < values.GetLength(0); row++)
                {
                    for (int column = 0; column < values.GetLength(1); column++)
                    {
                        if (values[row, column] is not null)
                        {
                            populated++;
                        }
                    }
                }

                if (populated != yearCount)
                {
                    failures.Add("The year header row " + name + " holds " + Format(populated) + " years rather than "
                        + Format(yearCount) + ".");
                }
            }
            finally
            {
                ExcelUtils.ReleaseComObject(range);
            }
        }
    }

    /// <summary>
    /// Checks that no defined name of the workbook refers to anything bracketed.
    /// </summary>
    /// <param name="workbook">The workbook holding the applied analysis.</param>
    /// <param name="failures">The list every failure is added to.</param>
    /// <remarks>
    /// A name carried in from another workbook refers to it as [Other.xlsx]SimAnalysisCalc!$AT$4#, resolves silently
    /// against whatever copy of that file is open, and produces plausible numbers from the wrong sweep with nothing
    /// reported.  It is the failure this whole command is arranged to make impossible, so it is also checked for.
    /// </remarks>
    private static void CheckNames(Excel.Workbook workbook, List<string> failures)
    {
        foreach ((string name, string refersTo) in ExcelUtils.GetNameDefinitions(workbook))
        {
            if (!name.StartsWith('_') && refersTo.Contains('[', StringComparison.Ordinal))
            {
                failures.Add("The defined name " + name + " refers to " + refersTo + ", which is bracketed: a name "
                    + "that reaches outside this workbook reads the wrong sweep without saying so.");
            }
        }
    }

    /// <summary>
    /// Checks the cells of the analysis worksheets for errors and for formulas stored as text, and counts the blank
    /// trial slots.
    /// </summary>
    /// <param name="workbook">The workbook holding the applied analysis.</param>
    /// <param name="worksheetNames">The worksheets that hold formulas.</param>
    /// <param name="failures">The list every failure is added to.</param>
    /// <returns>The number of cells holding the #N/A of a trial slot with no trial.</returns>
    /// <remarks>
    /// #N/A is allowed in the trial blocks and nowhere else: a slot with no trial to plot returns it deliberately so
    /// that the chart draws no line.  Every other error, and every cell whose value is a string beginning with an
    /// equals sign, is a formula that did not take.
    /// </remarks>
    private static int CheckCellContents(Excel.Workbook workbook, IReadOnlyList<string> worksheetNames,
        List<string> failures)
    {
        int blankTrialSlots = 0;
        Dictionary<string, HashSet<int>> trialRowsByWorksheet = ReadTrialRows(workbook);

        foreach (string worksheetName in worksheetNames)
        {
            Excel.Worksheet worksheet = ExcelUtils.GetWorksheet(workbook, worksheetName);
            try
            {
                HashSet<int> trialRows = trialRowsByWorksheet.TryGetValue(worksheetName, out HashSet<int>? rows)
                    ? rows
                    : [];

                CellBlock block = CellBlock.Read(worksheet);

                for (int row = 0; row < block.RowCount; row++)
                {
                    for (int column = 0; column < block.ColumnCount; column++)
                    {
                        object? value = block.Values[row, column];

                        if (ExcelUtils.IsError(value, out string? error))
                        {
                            if (value is ExcelUtils.NotAvailableError && trialRows.Contains(block.RowOf(row)))
                            {
                                blankTrialSlots++;
                                continue;
                            }

                            failures.Add(worksheetName + "!" + block.AddressOf(row, column) + " holds " + error + ".");
                        }
                        else if (value is string text && text.StartsWith('='))
                        {
                            failures.Add(worksheetName + "!" + block.AddressOf(row, column) + " holds the text "
                                + text + " rather than the formula it reads as.");
                        }
                    }
                }
            }
            finally
            {
                ExcelUtils.ReleaseComObject(worksheet);
            }
        }

        return blankTrialSlots;
    }

    /// <summary>
    /// Checks that no chart of the analysis covers a cell that holds something.
    /// </summary>
    /// <param name="workbook">The workbook holding the applied analysis.</param>
    /// <param name="worksheetNames">The worksheets that hold formulas, which are the ones charts are drawn on.</param>
    /// <param name="failures">The list every failure is added to.</param>
    /// <remarks>
    /// A chart that hides a block is a defect a reader of the output cannot see around and cannot fix, since the
    /// output holds no formulas to rebuild.  What is checked is the overlap of two boxes in both directions rather
    /// than the vertical clearance alone: a chart sits either in the blank band of rows beneath the block it charts or
    /// in blank columns beside it, ex. the sleeve snapshot pies, and a check that measured rows alone would report
    /// every chart of the second kind while passing a chart that covers a block to its left.
    /// </remarks>
    private static void CheckChartClearance(Excel.Workbook workbook, IReadOnlyList<string> worksheetNames,
        List<string> failures)
    {
        foreach (string worksheetName in worksheetNames)
        {
            Excel.Worksheet worksheet = ExcelUtils.GetWorksheet(workbook, worksheetName);
            try
            {
                CheckChartClearance(worksheet, worksheetName, failures);
            }
            finally
            {
                ExcelUtils.ReleaseComObject(worksheet);
            }
        }
    }

    /// <summary>
    /// Checks that no chart of one worksheet covers a cell that holds something.
    /// </summary>
    /// <param name="worksheet">The worksheet to check.</param>
    /// <param name="worksheetName">The name of the worksheet, for the message.</param>
    /// <param name="failures">The list every failure is added to.</param>
    private static void CheckChartClearance(Excel.Worksheet worksheet, string worksheetName, List<string> failures)
    {
        IReadOnlyList<(string Name, double Left, double Top, double Width, double Height)> charts =
            ExcelUtils.GetChartBoxes(worksheet);

        if (charts.Count == 0)
        {
            return;
        }

        CellBlock block = CellBlock.Read(worksheet);
        int lastRow = block.FirstRow + block.RowCount - 1;
        int lastColumn = block.FirstColumn + block.ColumnCount - 1;

        if (lastRow < 1 || lastColumn < 1)
        {
            return;
        }

        // Nothing below the lowest chart or right of the widest one can be covered by any of them, so that is as far
        // as the sheet is measured.  A calculation worksheet is as long as the sweep, and measuring it row by row to
        // the end would cost a call per simulation.
        double[] tops = ExcelUtils.GetRowTops(worksheet, lastRow, charts.Max(chart => chart.Top + chart.Height));
        double[] lefts = ExcelUtils.GetColumnLefts(worksheet, lastColumn,
            charts.Max(chart => chart.Left + chart.Width));

        lastRow = Math.Min(lastRow, tops.Length - 2);
        lastColumn = Math.Min(lastColumn, lefts.Length - 2);

        foreach ((string name, double left, double top, double width, double height) in charts)
        {
            for (int row = block.FirstRow; row <= lastRow; row++)
            {
                // A row whose band of points does not meet the chart's cannot hold a cell the chart covers, and the
                // rows are in order, so the first row past the bottom of the chart ends the search.
                if (tops[row + 1] <= top)
                {
                    continue;
                }

                if (tops[row] >= top + height)
                {
                    break;
                }

                int column = FirstCoveredColumn(block, lefts, row, left, left + width, lastColumn);

                if (column > 0)
                {
                    failures.Add("The chart " + name + " covers " + worksheetName + "!"
                        + CellBlock.Reference(row, column) + ", which holds data.  A chart belongs in the blank band "
                        + "of rows beneath the block it charts, or in blank columns beside it.");

                    break;
                }
            }
        }
    }

    /// <summary>
    /// Finds the first column of one row holding something that a chart's span across the sheet covers.
    /// </summary>
    /// <param name="block">The used cells of the worksheet.</param>
    /// <param name="lefts">The left of each column, in points, indexed by one based column.</param>
    /// <param name="row">The one based worksheet row to search.</param>
    /// <param name="chartLeft">The left of the chart, in points.</param>
    /// <param name="chartRight">The right of the chart, in points.</param>
    /// <param name="lastColumn">The last column of the block.</param>
    /// <returns>The one based column, or zero when the chart covers nothing on that row.</returns>
    private static int FirstCoveredColumn(CellBlock block, double[] lefts, int row, double chartLeft,
        double chartRight, int lastColumn)
    {
        for (int column = block.FirstColumn; column <= lastColumn; column++)
        {
            if (lefts[column + 1] <= chartLeft)
            {
                continue;
            }

            if (lefts[column] >= chartRight)
            {
                break;
            }

            if (!block.IsCellEmpty(row, column))
            {
                return column;
            }
        }

        return 0;
    }

    /// <summary>
    /// Reconciles the analysis's count of simulations that fell below the income floor against the shortfall metric of
    /// the results.
    /// </summary>
    /// <param name="workbook">The workbook holding the applied analysis.</param>
    /// <param name="blocks">The metric blocks the headings describe.</param>
    /// <param name="rowCount">The number of simulations that were written.</param>
    /// <param name="failures">The list every failure is added to.</param>
    /// <returns>The count the analysis reported.</returns>
    /// <remarks>
    /// This is the one check that reads the results themselves rather than what the analysis says about them, and so
    /// the one that establishes that the analysis is reading the columns it means to.  The count is bracketed rather
    /// than matched exactly, because the analysis applies a tolerance to a dollar amount and the tolerance is the
    /// analysis's to state: a count between the simulations with any shortfall at all and those whose shortfall is
    /// beyond rounding is a count computed over the right columns.
    /// </remarks>
    private static int CheckShortfallReconciliation(Excel.Workbook workbook, MetricBlocks blocks, int rowCount,
        List<string> failures)
    {
        int reported = (int)ExcelUtils.GetNameNumber(workbook, AnalysisTemplate.BelowFloorCountName);

        if (blocks.ShortfallFirstColumn < 1)
        {
            failures.Add("The results carry no " + blocks.ShortfallMetricKey + " columns, so the count of "
                + "simulations that failed to fund the floor cannot be reconciled.  The harvest is required to record "
                + "it: the funding ratio alone cannot tell a cut to the floor from a failure to reach it.");

            return reported;
        }

        int columns = blocks.ShortfallLastColumn - blocks.ShortfallFirstColumn + 1;
        int any = 0;
        int beyondTolerance = 0;

        Excel.Worksheet worksheet = ExcelUtils.GetWorksheet(workbook, AnalysisTemplate.SimDataWorksheet);
        try
        {
            for (int first = 0; first < rowCount; first += ReadBlockRows)
            {
                int rows = Math.Min(ReadBlockRows, rowCount - first);
                object?[,] block = ExcelUtils.ReadGrid(worksheet, first + 2, blocks.ShortfallFirstColumn, rows,
                    columns);

                for (int row = 0; row < rows; row++)
                {
                    double worst = 0;
                    for (int column = 0; column < columns; column++)
                    {
                        if (block[row, column] is double value && value > worst)
                        {
                            worst = value;
                        }
                    }

                    if (worst > 0)
                    {
                        any++;
                    }

                    if (worst > ShortfallTolerance)
                    {
                        beyondTolerance++;
                    }
                }
            }
        }
        finally
        {
            ExcelUtils.ReleaseComObject(worksheet);
        }

        if (reported < beyondTolerance || reported > any)
        {
            failures.Add(AnalysisTemplate.BelowFloorCountName + " is " + Format(reported) + " and the "
                + blocks.ShortfallMetricKey + " columns of the results report a shortfall in "
                + Format(beyondTolerance) + " to " + Format(any) + " simulations, so the two do not describe the same "
                + "sweep.");
        }

        return reported;
    }

    /// <summary>
    /// Reads the rows the blocks of individual trials occupy, which are the only rows allowed to hold #N/A.
    /// </summary>
    /// <param name="workbook">The workbook holding the applied analysis.</param>
    /// <returns>The one based rows of each worksheet that carries a trial block.</returns>
    /// <remarks>
    /// The blocks are found through the defined names that carry the trial block marker rather than listed by name,
    /// so a trial block added to the template is recognized without this tool being told about it.
    /// </remarks>
    private static Dictionary<string, HashSet<int>> ReadTrialRows(Excel.Workbook workbook)
    {
        var rows = new Dictionary<string, HashSet<int>>(StringComparer.Ordinal);

        foreach ((string name, _) in ExcelUtils.GetNameDefinitions(workbook))
        {
            if (name.StartsWith('_')
                || !name.Contains(AnalysisTemplate.TrialBlockNameMarker, StringComparison.Ordinal))
            {
                continue;
            }

            Excel.Range? range = null;
            Excel.Range? lines = null;
            Excel.Worksheet? worksheet = null;
            try
            {
                range = ExcelUtils.GetNameRange(workbook, name);
                lines = range.Rows;
                worksheet = (Excel.Worksheet)range.Worksheet;

                if (!rows.TryGetValue(worksheet.Name, out HashSet<int>? worksheetRows))
                {
                    worksheetRows = [];
                    rows.Add(worksheet.Name, worksheetRows);
                }

                for (int row = range.Row; row < range.Row + lines.Count; row++)
                {
                    worksheetRows.Add(row);
                }
            }
            finally
            {
                ExcelUtils.ReleaseComObject(worksheet);
                ExcelUtils.ReleaseComObject(lines);
                ExcelUtils.ReleaseComObject(range);
            }
        }

        return rows;
    }

    /// <summary>
    /// Reads the text of every cell a defined name refers to.
    /// </summary>
    /// <param name="workbook">The workbook that defines the name.</param>
    /// <param name="definedName">The name to read.</param>
    /// <returns>The text of each cell, in reading order.  An empty cell yields an empty string.</returns>
    private static List<string> ReadTexts(Excel.Workbook workbook, string definedName)
    {
        Excel.Range? range = null;
        try
        {
            range = ExcelUtils.GetNameRange(workbook, definedName);
            object?[,] values = ExcelUtils.ReadRange(range);

            var texts = new List<string>(values.Length);
            for (int row = 0; row < values.GetLength(0); row++)
            {
                for (int column = 0; column < values.GetLength(1); column++)
                {
                    texts.Add(Convert.ToString(values[row, column], CultureInfo.InvariantCulture) ?? string.Empty);
                }
            }

            return texts;
        }
        finally
        {
            ExcelUtils.ReleaseComObject(range);
        }
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
/// What the column headings of one sweep's results say about the metrics they carry.
/// </summary>
/// <remarks>
/// The headings are the contract between a sweep and the analysis: leading identity columns, then one block of year
/// columns per metric named MetricKey_YYYY.  Everything the checks need to know about a dataset is read off them once,
/// here, rather than by asking Excel about a column at a time.
/// </remarks>
internal sealed class MetricBlocks
{
    private MetricBlocks(HashSet<string> headings, IReadOnlyList<int> years, int identityColumnCount,
        string shortfallMetricKey, int shortfallFirstColumn, int shortfallLastColumn)
    {
        Headings = headings;
        Years = years;
        IdentityColumnCount = identityColumnCount;
        ShortfallMetricKey = shortfallMetricKey;
        ShortfallFirstColumn = shortfallFirstColumn;
        ShortfallLastColumn = shortfallLastColumn;
    }

    /// <summary>
    /// Gets every column heading of the results.
    /// </summary>
    internal HashSet<string> Headings { get; }

    /// <summary>
    /// Gets the projected years, ascending and without repetition, as every metric block of the results states them
    /// between them.
    /// </summary>
    /// <remarks>
    /// The years are read off the headings as a whole rather than off one metric named here, so which metric the
    /// template counts its horizon from is the template's business.  Every metric of a harvest carries the same block
    /// of years, so taking them together says what taking any one of them would.
    /// </remarks>
    internal IReadOnlyList<int> Years { get; }

    /// <summary>
    /// Gets the number of columns before the first metric block, or minus one when there is no metric block to
    /// establish it.
    /// </summary>
    internal int IdentityColumnCount { get; }

    /// <summary>
    /// Gets the metric key the shortfall reconciliation was located by.
    /// </summary>
    internal string ShortfallMetricKey { get; }

    /// <summary>
    /// Gets the one based first column of the shortfall metric's block, or zero when the results carry none.
    /// </summary>
    internal int ShortfallFirstColumn { get; }

    /// <summary>
    /// Gets the one based last column of the shortfall metric's block, or zero when the results carry none.
    /// </summary>
    internal int ShortfallLastColumn { get; }

    /// <summary>
    /// Reads what the column headings of a sweep's results say about the metrics they carry.
    /// </summary>
    /// <param name="header">The column headings, in column order.</param>
    /// <param name="shortfallMetricKey">The metric key the shortfall reconciliation is run against.</param>
    /// <param name="failures">The list a heading contract failure is added to.</param>
    /// <returns>The metric blocks.</returns>
    internal static MetricBlocks Read(IReadOnlyList<string> header, string shortfallMetricKey, List<string> failures)
    {
        var headings = new HashSet<string>(header, StringComparer.Ordinal);
        var years = new SortedSet<int>();
        int identityColumnCount = -1;
        int shortfallFirst = 0;
        int shortfallLast = 0;

        for (int column = 0; column < header.Count; column++)
        {
            string heading = header[column];
            int underscore = heading.LastIndexOf('_');

            if (underscore < 1 || !int.TryParse(heading[(underscore + 1)..], NumberStyles.None,
                    CultureInfo.InvariantCulture, out int year))
            {
                continue;
            }

            string key = heading[..underscore];

            if (identityColumnCount < 0)
            {
                identityColumnCount = column;
            }

            years.Add(year);

            if (string.Equals(key, shortfallMetricKey, StringComparison.Ordinal))
            {
                shortfallFirst = shortfallFirst == 0 ? column + 1 : shortfallFirst;
                shortfallLast = column + 1;
            }
        }

        if (years.Count == 0)
        {
            failures.Add("The results carry no column named MetricKey_YYYY, which is what the analysis reads the "
                + "projected years off.");
        }

        if (shortfallLast - shortfallFirst + 1 != years.Count && shortfallFirst > 0)
        {
            failures.Add("The " + shortfallMetricKey + " columns of the results are not one contiguous block of the "
                + "projected years.");
        }

        return new MetricBlocks(headings, [.. years], identityColumnCount, shortfallMetricKey, shortfallFirst,
            shortfallLast);
    }
}

/// <summary>
/// The used cells of one worksheet, read in a single call, with what it takes to say where a cell of the block sits on
/// the sheet.
/// </summary>
internal sealed class CellBlock
{
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    private CellBlock(object?[,] values, int firstRow, int firstColumn)
    {
        Values = values;
        FirstRow = firstRow;
        FirstColumn = firstColumn;
        RowCount = values.GetLength(0);
        ColumnCount = values.GetLength(1);
    }

    /// <summary>
    /// Gets the values, indexed from zero by row and then by column.
    /// </summary>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    internal object?[,] Values { get; }

    /// <summary>
    /// Gets the one based worksheet row the block starts at.
    /// </summary>
    internal int FirstRow { get; }

    /// <summary>
    /// Gets the one based worksheet column the block starts at.
    /// </summary>
    internal int FirstColumn { get; }

    /// <summary>
    /// Gets the number of rows the block covers.
    /// </summary>
    internal int RowCount { get; }

    /// <summary>
    /// Gets the number of columns the block covers.
    /// </summary>
    internal int ColumnCount { get; }

    /// <summary>
    /// Reads the used cells of a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to read.</param>
    /// <returns>The block.</returns>
    internal static CellBlock Read(Excel.Worksheet worksheet)
    {
        Excel.Range? used = null;
        try
        {
            used = worksheet.UsedRange;
            return new CellBlock(ExcelUtils.ReadRange(used), used.Row, used.Column);
        }
        finally
        {
            ExcelUtils.ReleaseComObject(used);
        }
    }

    /// <summary>
    /// Gets the one based worksheet row a row of the block sits on.
    /// </summary>
    /// <param name="row">The zero based row within the block.</param>
    /// <returns>The worksheet row.</returns>
    internal int RowOf(int row)
    {
        return FirstRow + row;
    }

    /// <summary>
    /// Indicates whether one cell of the worksheet holds nothing.
    /// </summary>
    /// <param name="row">The one based worksheet row.</param>
    /// <param name="column">The one based worksheet column.</param>
    /// <returns>True when the cell holds nothing, or lies outside the block.</returns>
    internal bool IsCellEmpty(int row, int column)
    {
        int rowIndex = row - FirstRow;
        int columnIndex = column - FirstColumn;

        if (rowIndex < 0 || rowIndex >= RowCount || columnIndex < 0 || columnIndex >= ColumnCount)
        {
            return true;
        }

        return Values[rowIndex, columnIndex] is null;
    }

    /// <summary>
    /// Describes where a cell of the block sits on the worksheet, in A1 notation.
    /// </summary>
    /// <param name="row">The zero based row within the block.</param>
    /// <param name="column">The zero based column within the block.</param>
    /// <returns>The address, ex. AP118.</returns>
    internal string AddressOf(int row, int column)
    {
        return Reference(FirstRow + row, FirstColumn + column);
    }

    /// <summary>
    /// Describes a cell of a worksheet in A1 notation.
    /// </summary>
    /// <param name="row">The one based worksheet row.</param>
    /// <param name="column">The one based worksheet column.</param>
    /// <returns>The address, ex. AP118.</returns>
    internal static string Reference(int row, int column)
    {
        return ColumnLetters(column) + row.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Converts a one based column number into the letters Excel names it with.
    /// </summary>
    /// <param name="column">The one based column number.</param>
    /// <returns>The letters, ex. AP for 42.</returns>
    private static string ColumnLetters(int column)
    {
        string letters = string.Empty;

        for (int remaining = column; remaining > 0; remaining = (remaining - 1) / 26)
        {
            letters = (char)('A' + (remaining - 1) % 26) + letters;
        }

        return letters;
    }
}
