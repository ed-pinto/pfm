using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pfm;

/// <summary>
/// COM automation helpers for Excel.  These helpers know how to reach an Excel instance and how to address the contents
/// of a workbook, but they carry no knowledge of the PortfolioManager model itself; that belongs in
/// <see cref="Operations"/>.
/// </summary>
public static class ExcelUtils
{
    /// <summary>
    /// Acquires the workbook at the specified path, attaching to it when it is already open in Excel and otherwise
    /// starting a hidden Excel instance and opening it.
    /// </summary>
    /// <param name="filePath">The path to the workbook.</param>
    /// <returns>
    /// A session that owns the Excel and workbook references.  Disposing the session restores the Excel application
    /// settings it changed and, when the session started Excel, closes the workbook and quits.
    /// </returns>
    public static ExcelSession OpenWorkbook(string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);

        string fullPath = Path.GetFullPath(filePath);

        Excel.Workbook? running = TryGetOpenWorkbook(fullPath);
        if (running != null)
        {
            Log.Logger.Information("PFM_EXCEL_ATTACHED: " + fullPath);
            return new ExcelSession(running.Application, running, startedExcel: false);
        }

        Log.Logger.Information("PFM_EXCEL_LAUNCHED: " + fullPath);

        var excel = new Excel.Application
        {
            Visible = false,
            DisplayAlerts = false,
            AskToUpdateLinks = false
        };

        Excel.Workbooks workbooks = excel.Workbooks;
        try
        {
            Excel.Workbook opened = workbooks.Open(Filename: fullPath, UpdateLinks: 0, ReadOnly: false);
            return new ExcelSession(excel, opened, startedExcel: true);
        }
        finally
        {
            ReleaseComObject(workbooks);
        }
    }

    /// <summary>
    /// Starts an Excel instance in a process of its own and opens the workbook at the specified path in it.
    /// </summary>
    /// <param name="filePath">The path to the workbook.</param>
    /// <returns>
    /// A session that owns the new instance.  Disposing the session closes the workbook and ends that process.
    /// </returns>
    /// <remarks>
    /// A sweep runs several instances of this tool at once, and two of them sharing one Excel process would share its
    /// calculation state: the workbook one job is driving would be recalculated by the other job's writes.  Isolation
    /// is therefore a correctness requirement rather than a preference, so this deliberately does not consult the
    /// running object table the way <see cref="OpenWorkbook"/> does.  Excel registers its class factory as single use,
    /// so each activation starts a fresh excel.exe; the resulting process is checked against the instances that were
    /// already running so that a violation of that assumption fails here rather than silently corrupting a sweep.
    /// </remarks>
    public static ExcelSession StartIsolatedWorkbook(string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);

        string fullPath = Path.GetFullPath(filePath);
        HashSet<int> existing = GetExcelProcessIds();

        var excel = new Excel.Application
        {
            Visible = false,
            DisplayAlerts = false,
            AskToUpdateLinks = false
        };

        int processId = GetProcessId(excel);

        if (processId == 0 || existing.Contains(processId))
        {
            // Quitting is the only safe way to leave an instance we cannot claim; another job may be driving it.
            TryQuit(excel);
            ReleaseComObject(excel);

            throw new InvalidOperationException("Excel did not start in a process of its own; it "
                + (processId == 0 ? "did not report a process id" : "reused process " + Format(processId))
                + ".  A simulation cannot share an Excel process with another job.");
        }

        Log.Logger.Information("PFM_EXCEL_ISOLATED: process=" + Format(processId) + " " + fullPath);

        Excel.Workbooks workbooks = excel.Workbooks;
        try
        {
            Excel.Workbook opened = workbooks.Open(Filename: fullPath, UpdateLinks: 0, ReadOnly: false);
            return new ExcelSession(excel, opened, startedExcel: true, processId: processId);
        }
        finally
        {
            ReleaseComObject(workbooks);
        }
    }

    /// <summary>
    /// Creates an empty workbook in an Excel instance of its own.
    /// </summary>
    /// <returns>
    /// A session that owns the new instance and workbook.  The workbook has never been saved, so a caller that wants
    /// to keep it must call <see cref="ExcelSession.SaveAs"/> before disposing the session, which otherwise closes it
    /// without saving.
    /// </returns>
    /// <remarks>
    /// The instance is started rather than attached to for the same reason a simulation starts its own: a workbook
    /// added to whatever Excel the user happens to have open would appear in front of them, and the settings this
    /// session suspends would be theirs to have disturbed.  Asking for the one worksheet template rather than deleting
    /// the surplus sheets afterwards is what keeps the workbook's shape independent of the machine's Excel settings.
    /// </remarks>
    public static ExcelSession CreateWorkbook()
    {
        HashSet<int> existing = GetExcelProcessIds();

        var excel = new Excel.Application
        {
            Visible = false,
            DisplayAlerts = false,
            AskToUpdateLinks = false
        };

        int processId = GetProcessId(excel);

        if (processId != 0 && existing.Contains(processId))
        {
            TryQuit(excel);
            ReleaseComObject(excel);

            throw new InvalidOperationException("Excel did not start in a process of its own; it reused process "
                + Format(processId) + ".  A new workbook cannot be built in an Excel process that is already in use.");
        }

        Log.Logger.Information("PFM_EXCEL_CREATED: process=" + Format(processId));

        Excel.Workbooks workbooks = excel.Workbooks;
        try
        {
            Excel.Workbook created = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            return new ExcelSession(excel, created, startedExcel: true, processId: processId);
        }
        finally
        {
            ReleaseComObject(workbooks);
        }
    }

    /// <summary>
    /// Gets a worksheet of the workbook by its position.
    /// </summary>
    /// <param name="workbook">The workbook to read.</param>
    /// <param name="index">The one based position of the worksheet.</param>
    /// <returns>The worksheet.</returns>
    /// <remarks>
    /// A workbook Excel has just created names its worksheet in the language of the installation, ex. Feuil1 rather
    /// than Sheet1, so the sheet that is about to be renamed has to be reached by position rather than by name.
    /// </remarks>
    public static Excel.Worksheet GetWorksheet(Excel.Workbook workbook, int index)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.Sheets worksheets = workbook.Worksheets;
        try
        {
            return (Excel.Worksheet)worksheets[index];
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("The workbook has no worksheet at position " + Format(index) + ".",
                ex);
        }
        finally
        {
            ReleaseComObject(worksheets);
        }
    }

    /// <summary>
    /// Adds a worksheet to the end of the workbook.
    /// </summary>
    /// <param name="workbook">The workbook to add to.</param>
    /// <param name="worksheetName">The name to give the worksheet.</param>
    /// <returns>The worksheet that was added.</returns>
    public static Excel.Worksheet AddWorksheet(Excel.Workbook workbook, string worksheetName)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(worksheetName);

        Excel.Sheets worksheets = workbook.Worksheets;
        Excel.Worksheet? last = null;
        try
        {
            last = (Excel.Worksheet)worksheets[worksheets.Count];
            var added = (Excel.Worksheet)worksheets.Add(After: last);
            added.Name = worksheetName;
            return added;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Worksheet " + worksheetName + " could not be added to the workbook.",
                ex);
        }
        finally
        {
            ReleaseComObject(last);
            ReleaseComObject(worksheets);
        }
    }

    /// <summary>
    /// Writes a rectangular block of values into a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to write into.</param>
    /// <param name="firstRow">The one based row the block starts at.</param>
    /// <param name="firstColumn">The one based column the block starts at.</param>
    /// <param name="values">The values to write, indexed by row and then by column.  Null leaves a cell empty.</param>
    /// <remarks>
    /// The block is assigned to the range in one cross process call, which is what makes importing thousands of rows
    /// cost seconds rather than minutes; writing a cell at a time would be one call per cell.
    /// </remarks>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    public static void WriteGrid(Excel.Worksheet worksheet, int firstRow, int firstColumn, object?[,] values)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        ArgumentNullException.ThrowIfNull(values);

        int rowCount = values.GetLength(0);
        int columnCount = values.GetLength(1);

        if (rowCount == 0 || columnCount == 0)
        {
            return;
        }

        Excel.Range? first = null;
        Excel.Range? last = null;
        Excel.Range? block = null;
        try
        {
            first = (Excel.Range)worksheet.Cells[firstRow, firstColumn];
            last = (Excel.Range)worksheet.Cells[firstRow + rowCount - 1, firstColumn + columnCount - 1];
            block = worksheet.Range[first, last];
            block.Value2 = values;
        }
        finally
        {
            ReleaseComObject(block);
            ReleaseComObject(last);
            ReleaseComObject(first);
        }
    }

    /// <summary>
    /// Formats a block of a worksheet as text, so that Excel stores what is written into it verbatim.
    /// </summary>
    /// <param name="worksheet">The worksheet whose cells are being formatted.</param>
    /// <param name="firstRow">The one based row the block starts at.</param>
    /// <param name="firstColumn">The one based column the block starts at.</param>
    /// <param name="rowCount">The number of rows the block covers.</param>
    /// <param name="columnCount">The number of columns the block covers.</param>
    /// <remarks>
    /// A line of prose that happens to begin with an equals sign, or that reads like a date, would otherwise be taken
    /// as a formula or converted as it was written.  The synopsis is a report to be read back exactly as the workbook
    /// wrote it, so the cells it lands in say so before anything is written to them.  It is the block rather than the
    /// whole column, because a sheet that carries the synopsis also carries data below it, and a number written into
    /// a text formatted cell is stored as text.
    /// </remarks>
    public static void FormatRangeAsText(Excel.Worksheet worksheet, int firstRow, int firstColumn, int rowCount,
        int columnCount)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        if (rowCount <= 0 || columnCount <= 0)
        {
            return;
        }

        Excel.Range? first = null;
        Excel.Range? last = null;
        Excel.Range? block = null;
        try
        {
            first = (Excel.Range)worksheet.Cells[firstRow, firstColumn];
            last = (Excel.Range)worksheet.Cells[firstRow + rowCount - 1, firstColumn + columnCount - 1];
            block = worksheet.Range[first, last];
            block.NumberFormat = "@";
        }
        finally
        {
            ReleaseComObject(block);
            ReleaseComObject(last);
            ReleaseComObject(first);
        }
    }

    /// <summary>
    /// Emboldens one cell of a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet whose cell is being formatted.</param>
    /// <param name="row">The one based row of the cell.</param>
    /// <param name="column">The one based column of the cell.</param>
    /// <remarks>
    /// This exists for the labels that name the configuration tables stacked down a sheet.  A table styles its own
    /// header, and the label above it is the only part of the block that would otherwise read as loose text.
    /// </remarks>
    public static void SetCellBold(Excel.Worksheet worksheet, int row, int column)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        Excel.Range? cell = null;
        Excel.Font? font = null;
        try
        {
            cell = (Excel.Range)worksheet.Cells[row, column];
            font = cell.Font;
            font.Bold = true;
        }
        finally
        {
            ReleaseComObject(font);
            ReleaseComObject(cell);
        }
    }

    /// <summary>
    /// Defines a table (ListObject) over a block of a worksheet whose first row is the column headings.
    /// </summary>
    /// <param name="worksheet">The worksheet the block sits on.</param>
    /// <param name="tableName">The name to give the table.</param>
    /// <param name="firstRow">The one based row of the heading row.</param>
    /// <param name="firstColumn">The one based column the block starts at.</param>
    /// <param name="rowCount">The number of rows the block covers, the heading row included.</param>
    /// <param name="columnCount">The number of columns the block covers.</param>
    /// <remarks>
    /// Excel holds no table of a heading row alone, so a caller with nothing to show must still leave the blank row
    /// beneath the headings to the table.
    /// </remarks>
    public static void AddTable(Excel.Worksheet worksheet, string tableName, int firstRow, int firstColumn,
        int rowCount, int columnCount)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        ArgumentNullException.ThrowIfNull(tableName);
        ArgumentOutOfRangeException.ThrowIfLessThan(rowCount, 2);
        ArgumentOutOfRangeException.ThrowIfLessThan(columnCount, 1);

        Excel.Range? first = null;
        Excel.Range? last = null;
        Excel.Range? block = null;
        Excel.ListObjects? tables = null;
        Excel.ListObject? added = null;
        try
        {
            first = (Excel.Range)worksheet.Cells[firstRow, firstColumn];
            last = (Excel.Range)worksheet.Cells[firstRow + rowCount - 1, firstColumn + columnCount - 1];
            block = worksheet.Range[first, last];
            tables = worksheet.ListObjects;

            added = tables.Add(Excel.XlListObjectSourceType.xlSrcRange, block, Type.Missing,
                Excel.XlYesNoGuess.xlYes);
            added.Name = tableName;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Table " + tableName + " could not be defined on worksheet "
                + worksheet.Name + ".", ex);
        }
        finally
        {
            ReleaseComObject(added);
            ReleaseComObject(tables);
            ReleaseComObject(block);
            ReleaseComObject(last);
            ReleaseComObject(first);
        }
    }

    /// <summary>
    /// Gets the value a defined name resolves to, whatever kind of name it is.
    /// </summary>
    /// <param name="workbook">The workbook that defines the name.</param>
    /// <param name="definedName">The defined name to read.</param>
    /// <returns>The value the name resolves to.</returns>
    /// <remarks>
    /// The workbook defines names three ways: as a reference to a cell, ex. TaxProvisionTolerance, as a numeric
    /// constant, ex. SemestersPerYear, and as a text constant, ex. Hist.  Evaluating the name rather than parsing what
    /// it refers to reads all three the same way, and reads them the way the workbook's own formulas do.
    /// </remarks>
    public static object? EvaluateName(Excel.Workbook workbook, string definedName)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.Application excel = workbook.Application;
        try
        {
            object? value = excel.Evaluate(definedName);

            // A name that refers to a cell evaluates to the range, not to what the cell holds, so the reading kinds of
            // name are only alike once the range has been unwrapped.
            if (value is Excel.Range range)
            {
                try
                {
                    if (range.Count != 1)
                    {
                        throw new InvalidOperationException("Defined name " + definedName + " refers to "
                            + Format(range.Count) + " cells rather than one, so it has no single value.");
                    }

                    value = range.Value2;
                }
                finally
                {
                    ReleaseComObject(range);
                }
            }

            // Evaluate reports a name it cannot resolve as an error value rather than by throwing, and an error value
            // marshals back as the integer error code.
            if (value is null or int)
            {
                throw new InvalidOperationException("Defined name " + definedName + " did not evaluate to a value.");
            }

            return value;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Defined name " + definedName + " could not be evaluated.", ex);
        }
        finally
        {
            ReleaseComObject(excel);
        }
    }

    /// <summary>
    /// Reads the number a defined name resolves to.
    /// </summary>
    /// <param name="workbook">The workbook that defines the name.</param>
    /// <param name="definedName">The defined name to read.</param>
    /// <returns>The number the name resolves to.</returns>
    public static double GetNameNumber(Excel.Workbook workbook, string definedName)
    {
        object? value = EvaluateName(workbook, definedName);

        return value as double?
            ?? throw new InvalidOperationException("Defined name " + definedName + " resolved to a value that is not "
                + "a number.");
    }

    /// <summary>
    /// Writes a value into the single cell a defined name refers to.
    /// </summary>
    /// <param name="workbook">The workbook that defines the name.</param>
    /// <param name="definedName">The defined name to write through.</param>
    /// <param name="value">The value to write.</param>
    /// <remarks>
    /// Writing through the name rather than the address leaves the location of a parameter in the workbook's hands, so
    /// a cell that moves does not break this tool.  Ex. SimulationMode is 10_Parameters!B6 today.
    /// </remarks>
    public static void SetNameValue(Excel.Workbook workbook, string definedName, object value)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.Names? names = null;
        Excel.Name? name = null;
        Excel.Range? range = null;
        try
        {
            names = workbook.Names;

            try
            {
                name = names.Item(definedName);
                range = name.RefersToRange;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("Defined name " + definedName
                    + " was not found in the workbook, or does not refer to a range.", ex);
            }

            if (range.Count != 1)
            {
                throw new InvalidOperationException("Defined name " + definedName + " refers to "
                    + Format(range.Count) + " cells rather than one, so it cannot be assigned a value.");
            }

            range.Value2 = value;
        }
        finally
        {
            ReleaseComObject(range);
            ReleaseComObject(name);
            ReleaseComObject(names);
        }
    }

    /// <summary>
    /// Maps the header text of each of a table's columns to its zero based position in the table.
    /// </summary>
    /// <param name="table">The table to describe.</param>
    /// <returns>The column positions, keyed by header text.</returns>
    /// <remarks>
    /// The positions index the grid that <see cref="ReadTable"/> returns.  Resolving them once and reading the whole
    /// grid afterwards is what keeps a harvest to a single cross process call, however many columns it wants.
    /// </remarks>
    public static Dictionary<string, int> GetColumnIndexes(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        var indexes = new Dictionary<string, int>(StringComparer.Ordinal);

        Excel.ListColumns columns = table.ListColumns;
        try
        {
            foreach (Excel.ListColumn column in columns)
            {
                try
                {
                    indexes[column.Name] = column.Index - 1;
                }
                finally
                {
                    ReleaseComObject(column);
                }
            }
        }
        finally
        {
            ReleaseComObject(columns);
        }

        return indexes;
    }

    /// <summary>
    /// Reads the header text of a table's columns, in column order.
    /// </summary>
    /// <param name="table">The table to describe.</param>
    /// <returns>The column headings, in the order they appear in the table.</returns>
    /// <remarks>
    /// <see cref="GetColumnIndexes"/> answers where a named column is; this answers what the columns are, which is
    /// what copying a whole table out of the workbook needs.  The order is the table's own, so the headings line up
    /// with the grid that <see cref="ReadTableRows"/> returns.
    /// </remarks>
    public static string[] GetColumnNames(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        Excel.ListColumns columns = table.ListColumns;
        try
        {
            var names = new string[columns.Count];

            foreach (Excel.ListColumn column in columns)
            {
                try
                {
                    names[column.Index - 1] = column.Name;
                }
                finally
                {
                    ReleaseComObject(column);
                }
            }

            return names;
        }
        finally
        {
            ReleaseComObject(columns);
        }
    }

    /// <summary>
    /// Counts the columns of a table.
    /// </summary>
    /// <param name="table">The table to measure.</param>
    /// <returns>The number of columns, including any the caller does not read.</returns>
    /// <remarks>
    /// This asks Excel for the count rather than enumerating the columns, which is what makes it usable on a table as
    /// wide as MCSeeds: describing that one through <see cref="GetColumnIndexes"/> would be a thousand cross process
    /// calls to learn one number.
    /// </remarks>
    public static int GetColumnCount(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        Excel.ListColumns? columns = null;
        try
        {
            columns = table.ListColumns;
            return columns.Count;
        }
        finally
        {
            ReleaseComObject(columns);
        }
    }

    /// <summary>
    /// Gets the one based position of the named column within a table.
    /// </summary>
    /// <param name="table">The table that contains the column.</param>
    /// <param name="columnName">The header text of the column.</param>
    /// <returns>The position, which is the one an INDEX over the table's data body addresses the column by.</returns>
    public static int GetColumnPosition(Excel.ListObject table, string columnName)
    {
        ArgumentNullException.ThrowIfNull(table);

        Excel.ListColumns? columns = null;
        Excel.ListColumn? column = null;
        try
        {
            columns = table.ListColumns;

            try
            {
                column = columns[columnName];
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("Column " + columnName + " was not found in table " + table.Name
                    + ".", ex);
            }

            return column.Index;
        }
        finally
        {
            ReleaseComObject(column);
            ReleaseComObject(columns);
        }
    }

    /// <summary>
    /// Reads every data row of every column of a table in a single call.
    /// </summary>
    /// <param name="table">The table to read.</param>
    /// <returns>The values as a zero based grid indexed by row and then by column.  Empty cells are null.</returns>
    /// <remarks>
    /// A simulation harvests ten scattered columns of one table per pass.  Reading the whole data body costs one cross
    /// process call rather than ten, and the columns that are not harvested cost nothing beyond the marshalling.
    /// </remarks>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    public static object?[,] ReadTable(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        return ReadDataBody(table)
            ?? throw new InvalidOperationException("Table " + table.Name + " has no data rows.");
    }

    /// <summary>
    /// Reads every data row of every column of a table that may hold none.
    /// </summary>
    /// <param name="table">The table to read.</param>
    /// <returns>
    /// The values as a zero based grid indexed by row and then by column, with no rows when the table is empty.
    /// Empty cells are null.
    /// </returns>
    /// <remarks>
    /// The counterpart of <see cref="ReadTable"/>, for the tables a configuration is copied from: a plan that directs
    /// no draws is a plan, not a broken workbook, so an empty table is read as the empty table it is rather than
    /// reported.
    /// </remarks>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    public static object?[,] ReadTableRows(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        return ReadDataBody(table) ?? new object?[0, GetColumnCount(table)];
    }

    /// <summary>
    /// Reads the data body of a table in a single call.
    /// </summary>
    /// <param name="table">The table to read.</param>
    /// <returns>The values as a zero based grid, or null when the table has no data rows.</returns>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    private static object?[,]? ReadDataBody(Excel.ListObject table)
    {
        Excel.Range? range = table.DataBodyRange;

        if (range is null)
        {
            return null;
        }

        try
        {
            object? value = range.Value2;

            if (value is not object[,] grid)
            {
                // A single cell table yields the scalar value rather than a two dimensional array.
                return new object?[1, 1] { { value } };
            }

            int firstRow = grid.GetLowerBound(0);
            int firstColumn = grid.GetLowerBound(1);
            int rowCount = grid.GetLength(0);
            int columnCount = grid.GetLength(1);

            var values = new object?[rowCount, columnCount];
            for (int row = 0; row < rowCount; row++)
            {
                for (int column = 0; column < columnCount; column++)
                {
                    values[row, column] = grid[firstRow + row, firstColumn + column];
                }
            }

            return values;
        }
        finally
        {
            ReleaseComObject(range);
        }
    }

    /// <summary>
    /// Gets the named worksheet from the workbook.
    /// </summary>
    /// <param name="workbook">The workbook to search.</param>
    /// <param name="worksheetName">The name of the worksheet.</param>
    /// <returns>The worksheet.</returns>
    public static Excel.Worksheet GetWorksheet(Excel.Workbook workbook, string worksheetName)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.Sheets worksheets = workbook.Worksheets;
        try
        {
            return (Excel.Worksheet)worksheets[worksheetName];
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Worksheet " + worksheetName + " was not found in the workbook.", ex);
        }
        finally
        {
            ReleaseComObject(worksheets);
        }
    }

    /// <summary>
    /// Gets the named table (ListObject) from the named worksheet.
    /// </summary>
    /// <param name="workbook">The workbook to search.</param>
    /// <param name="worksheetName">The name of the worksheet that hosts the table.</param>
    /// <param name="tableName">The name of the table.</param>
    /// <returns>The table.</returns>
    public static Excel.ListObject GetListObject(Excel.Workbook workbook, string worksheetName, string tableName)
    {
        Excel.Worksheet worksheet = GetWorksheet(workbook, worksheetName);
        Excel.ListObjects? tables = null;
        try
        {
            tables = worksheet.ListObjects;
            return tables[tableName];
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Table " + tableName + " was not found on worksheet " + worksheetName
                + ".", ex);
        }
        finally
        {
            ReleaseComObject(tables);
            ReleaseComObject(worksheet);
        }
    }

    /// <summary>
    /// Gets the data body range of the named column of a table.  The header and any totals row are excluded.
    /// </summary>
    /// <param name="table">The table that contains the column.</param>
    /// <param name="columnName">The header text of the column.</param>
    /// <returns>The range covering the column's data rows.</returns>
    public static Excel.Range GetColumnRange(Excel.ListObject table, string columnName)
    {
        ArgumentNullException.ThrowIfNull(table);

        Excel.ListColumns? columns = null;
        Excel.ListColumn? column = null;
        try
        {
            columns = table.ListColumns;

            try
            {
                column = columns[columnName];
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("Column " + columnName + " was not found in table " + table.Name
                    + ".", ex);
            }

            return column.DataBodyRange
                ?? throw new InvalidOperationException("Column " + columnName + " in table " + table.Name
                    + " has no data rows.");
        }
        finally
        {
            ReleaseComObject(column);
            ReleaseComObject(columns);
        }
    }

    /// <summary>
    /// Reads every data row of the named column in a single call.
    /// </summary>
    /// <param name="table">The table that contains the column.</param>
    /// <param name="columnName">The header text of the column.</param>
    /// <returns>The column's values in row order.  Empty cells are null.</returns>
    /// <remarks>
    /// Reading a whole column at once rather than cell by cell keeps the number of cross process COM calls proportional
    /// to the number of columns instead of the number of rows.
    /// </remarks>
    public static object?[] ReadColumn(Excel.ListObject table, string columnName)
    {
        Excel.Range range = GetColumnRange(table, columnName);
        try
        {
            object? value = range.Value2;

            if (value is object[,] grid)
            {
                int firstRow = grid.GetLowerBound(0);
                int firstColumn = grid.GetLowerBound(1);
                int rowCount = grid.GetLength(0);

                var values = new object?[rowCount];
                for (int i = 0; i < rowCount; i++)
                {
                    values[i] = grid[firstRow + i, firstColumn];
                }

                return values;
            }

            // A single cell range yields the scalar value rather than a two dimensional array.
            return [value];
        }
        finally
        {
            ReleaseComObject(range);
        }
    }

    /// <summary>
    /// Writes a contiguous block of values into the named column in a single call.
    /// </summary>
    /// <param name="table">The table that contains the column.</param>
    /// <param name="columnName">The header text of the column.</param>
    /// <param name="rowOffset">The zero based data row at which the block starts.</param>
    /// <param name="values">The values to write.  One value per row.</param>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range assignment as a rectangular variant array; a jagged array "
            + "is not accepted.")]
    public static void WriteColumnBlock(Excel.ListObject table, string columnName, int rowOffset, object?[] values)
    {
        ArgumentNullException.ThrowIfNull(values);
        ArgumentOutOfRangeException.ThrowIfNegative(rowOffset);

        if (values.Length == 0)
        {
            return;
        }

        Excel.Range column = GetColumnRange(table, columnName);
        Excel.Range? shifted = null;
        Excel.Range? block = null;
        try
        {
            shifted = column.Offset[rowOffset, 0];
            block = shifted.Resize[values.Length, 1];

            if (values.Length == 1)
            {
                block.Value2 = values[0];
                return;
            }

            var grid = new object?[values.Length, 1];
            for (int i = 0; i < values.Length; i++)
            {
                grid[i, 0] = values[i];
            }

            block.Value2 = grid;
        }
        finally
        {
            ReleaseComObject(block);
            ReleaseComObject(shifted);
            ReleaseComObject(column);
        }
    }

    /// <summary>
    /// Reads the text of a defined name that holds a string constant rather than a cell reference.  The workbook uses
    /// these as enumerations, ex. ProvisionModelled.
    /// </summary>
    /// <param name="workbook">The workbook that defines the name.</param>
    /// <param name="definedName">The defined name to read.</param>
    /// <returns>The text the name refers to.</returns>
    public static string GetNameConstantText(Excel.Workbook workbook, string definedName)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.Names? names = null;
        Excel.Name? name = null;
        try
        {
            names = workbook.Names;

            try
            {
                name = names.Item(definedName);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("Defined name " + definedName + " was not found in the workbook.",
                    ex);
            }

            // A text constant name refers to a quoted literal, ex. ="ProvisionModelled".
            string refersTo = name.RefersTo as string ?? string.Empty;
            string literal = refersTo.StartsWith('=') ? refersTo[1..] : refersTo;

            if (literal.Length < 2 || literal[0] != '"' || literal[^1] != '"')
            {
                throw new InvalidOperationException("Defined name " + definedName + " refers to " + refersTo
                    + ", which is not a text constant.");
            }

            return literal[1..^1].Replace("\"\"", "\"", StringComparison.Ordinal);
        }
        finally
        {
            ReleaseComObject(name);
            ReleaseComObject(names);
        }
    }

    /// <summary>
    /// Gets the named worksheet from the workbook, or null when the workbook holds no worksheet of that name.
    /// </summary>
    /// <param name="workbook">The workbook to search.</param>
    /// <param name="worksheetName">The name of the worksheet.</param>
    /// <returns>The worksheet, or null when it is not there.</returns>
    /// <remarks>
    /// The counterpart of <see cref="GetWorksheet(Excel.Workbook, string)"/>, for a caller that is asking whether a
    /// sheet is present rather than requiring it, ex. one looking for the worksheet a sweep named its results after.
    /// </remarks>
    public static Excel.Worksheet? FindWorksheet(Excel.Workbook workbook, string worksheetName)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.Sheets worksheets = workbook.Worksheets;
        try
        {
            return (Excel.Worksheet)worksheets[worksheetName];
        }
        catch (COMException)
        {
            return null;
        }
        finally
        {
            ReleaseComObject(worksheets);
        }
    }

    /// <summary>
    /// Resizes a table to cover a block of its worksheet, leaving the table itself and every reference to it alone.
    /// </summary>
    /// <param name="table">The table to resize.</param>
    /// <param name="firstRow">The one based row of the heading row, which a resize cannot move.</param>
    /// <param name="firstColumn">The one based column the block starts at.</param>
    /// <param name="rowCount">The number of rows the block covers, the heading row included.</param>
    /// <param name="columnCount">The number of columns the block covers.</param>
    /// <remarks>
    /// This is how a table comes to cover a different sweep's rows, and the only way that leaves the analysis reading
    /// it.  A table that is dissolved takes its references with it: Excel rewrites every formula that read it
    /// structurally into the cell range the table happened to occupy, ex. ROWS(SimData) becomes
    /// ROWS(SimData!$A$2:$B$6), so the next sweep would be read through the last one's geometry with nothing reported.
    /// Deleting the table instead leaves those formulas as #REF!, and defining a second table of the same name is not
    /// possible at all: table names are workbook scoped, so Excel names it SimData1 and the references go on reading
    /// the original.  Resizing is therefore not one way of several; it is the only one.
    /// </remarks>
    public static void ResizeTable(Excel.ListObject table, int firstRow, int firstColumn, int rowCount,
        int columnCount)
    {
        ArgumentNullException.ThrowIfNull(table);
        ArgumentOutOfRangeException.ThrowIfLessThan(rowCount, 2);
        ArgumentOutOfRangeException.ThrowIfLessThan(columnCount, 1);

        Excel.Worksheet? worksheet = null;
        Excel.Range? first = null;
        Excel.Range? last = null;
        Excel.Range? block = null;
        try
        {
            worksheet = (Excel.Worksheet)table.Parent;
            first = (Excel.Range)worksheet.Cells[firstRow, firstColumn];
            last = (Excel.Range)worksheet.Cells[firstRow + rowCount - 1, firstColumn + columnCount - 1];
            block = worksheet.Range[first, last];

            table.Resize(block);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Table " + table.Name + " could not be resized to cover "
                + Format(rowCount) + " rows and " + Format(columnCount) + " columns from row " + Format(firstRow)
                + ".", ex);
        }
        finally
        {
            ReleaseComObject(block);
            ReleaseComObject(last);
            ReleaseComObject(first);
            ReleaseComObject(worksheet);
        }
    }

    /// <summary>
    /// Gets the block of the worksheet a table covers, its heading row included.
    /// </summary>
    /// <param name="table">The table to measure.</param>
    /// <returns>The corners of the block, as rows and columns.</returns>
    public static (int FirstRow, int FirstColumn, int RowCount, int ColumnCount) GetTableExtent(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        Excel.Range? range = null;
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        try
        {
            range = table.Range;
            rows = range.Rows;
            columns = range.Columns;

            return (range.Row, range.Column, rows.Count, columns.Count);
        }
        finally
        {
            ReleaseComObject(columns);
            ReleaseComObject(rows);
            ReleaseComObject(range);
        }
    }

    /// <summary>
    /// Gets the named table of a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet that hosts the table.</param>
    /// <param name="tableName">The name of the table.</param>
    /// <returns>The table.</returns>
    public static Excel.ListObject GetTable(Excel.Worksheet worksheet, string tableName)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        Excel.ListObjects? tables = null;
        try
        {
            tables = worksheet.ListObjects;
            return tables[tableName];
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Table " + tableName + " was not found on worksheet "
                + worksheet.Name + ".", ex);
        }
        finally
        {
            ReleaseComObject(tables);
        }
    }

    /// <summary>
    /// Inserts blank worksheet rows, moving everything below them down.
    /// </summary>
    /// <param name="worksheet">The worksheet to insert into.</param>
    /// <param name="firstRow">The one based row the blank rows are inserted at.</param>
    /// <param name="rowCount">The number of rows to insert.</param>
    /// <remarks>
    /// Whole rows rather than cells, because that is what moves a table below the insertion without splitting it, and
    /// what makes a table grow when the insertion lands inside its body: Excel maintains the ranges of the tables it
    /// moves, so the references to them survive.  Adding rows through a table itself is refused on a sheet that stacks
    /// tables, since that would shift part of a row and not the rest of it.
    /// </remarks>
    public static void InsertRows(Excel.Worksheet worksheet, int firstRow, int rowCount)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        if (rowCount <= 0)
        {
            return;
        }

        Excel.Range? rows = null;
        try
        {
            rows = (Excel.Range)worksheet.Rows[Format(firstRow) + ":" + Format(firstRow + rowCount - 1)];
            rows.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
        }
        finally
        {
            ReleaseComObject(rows);
        }
    }

    /// <summary>
    /// Deletes worksheet rows, moving everything below them up.
    /// </summary>
    /// <param name="worksheet">The worksheet to delete from.</param>
    /// <param name="firstRow">The one based first row to delete.</param>
    /// <param name="rowCount">The number of rows to delete.</param>
    public static void DeleteRows(Excel.Worksheet worksheet, int firstRow, int rowCount)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        if (rowCount <= 0)
        {
            return;
        }

        Excel.Range? rows = null;
        try
        {
            rows = (Excel.Range)worksheet.Rows[Format(firstRow) + ":" + Format(firstRow + rowCount - 1)];
            rows.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        }
        finally
        {
            ReleaseComObject(rows);
        }
    }

    /// <summary>
    /// Clears the contents of a band of worksheet rows, leaving their formatting alone.
    /// </summary>
    /// <param name="worksheet">The worksheet to clear.</param>
    /// <param name="firstRow">The one based first row to clear.</param>
    /// <param name="rowCount">The number of rows to clear.</param>
    public static void ClearRowContents(Excel.Worksheet worksheet, int firstRow, int rowCount)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        if (rowCount <= 0)
        {
            return;
        }

        Excel.Range? rows = null;
        try
        {
            rows = (Excel.Range)worksheet.Rows[Format(firstRow) + ":" + Format(firstRow + rowCount - 1)];
            rows.ClearContents();
        }
        finally
        {
            ReleaseComObject(rows);
        }
    }

    /// <summary>
    /// Clears whatever a worksheet holds outside a block anchored at its first cell.
    /// </summary>
    /// <param name="worksheet">The worksheet to clear.</param>
    /// <param name="lastRow">The last row of the block to keep.</param>
    /// <param name="lastColumn">The last column of the block to keep.</param>
    /// <remarks>
    /// A sheet written twice holds whatever the wider or longer of the two writes left behind.  The rows below and the
    /// columns beyond the block just written are therefore cleared rather than assumed empty: a stale row beneath a
    /// table reads as data to anything that measures the sheet rather than the table.
    /// </remarks>
    public static void ClearBeyond(Excel.Worksheet worksheet, int lastRow, int lastColumn)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        Excel.Range? used = null;
        try
        {
            used = worksheet.UsedRange;

            int usedLastRow;
            int usedLastColumn;
            Excel.Range? rows = null;
            Excel.Range? columns = null;
            try
            {
                rows = used.Rows;
                columns = used.Columns;
                usedLastRow = used.Row + rows.Count - 1;
                usedLastColumn = used.Column + columns.Count - 1;
            }
            finally
            {
                ReleaseComObject(columns);
                ReleaseComObject(rows);
            }

            if (usedLastRow > lastRow)
            {
                ClearRange(worksheet, lastRow + 1, 1, usedLastRow, usedLastColumn);
            }

            if (usedLastColumn > lastColumn)
            {
                ClearRange(worksheet, 1, lastColumn + 1, Math.Min(usedLastRow, lastRow), usedLastColumn);
            }
        }
        finally
        {
            ReleaseComObject(used);
        }
    }

    /// <summary>
    /// Clears the contents of a block of a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to clear.</param>
    /// <param name="firstRow">The one based first row of the block.</param>
    /// <param name="firstColumn">The one based first column of the block.</param>
    /// <param name="lastRow">The one based last row of the block.</param>
    /// <param name="lastColumn">The one based last column of the block.</param>
    private static void ClearRange(Excel.Worksheet worksheet, int firstRow, int firstColumn, int lastRow,
        int lastColumn)
    {
        if (lastRow < firstRow || lastColumn < firstColumn)
        {
            return;
        }

        Excel.Range? first = null;
        Excel.Range? last = null;
        Excel.Range? block = null;
        try
        {
            first = (Excel.Range)worksheet.Cells[firstRow, firstColumn];
            last = (Excel.Range)worksheet.Cells[lastRow, lastColumn];
            block = worksheet.Range[first, last];
            block.ClearContents();
        }
        finally
        {
            ReleaseComObject(block);
            ReleaseComObject(last);
            ReleaseComObject(first);
        }
    }

    /// <summary>
    /// Gets the block of cells a worksheet uses.
    /// </summary>
    /// <param name="worksheet">The worksheet to measure.</param>
    /// <returns>The corners of the block, as rows and columns.</returns>
    public static (int FirstRow, int FirstColumn, int LastRow, int LastColumn) GetUsedExtent(
        Excel.Worksheet worksheet)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        Excel.Range? used = null;
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        try
        {
            used = worksheet.UsedRange;
            rows = used.Rows;
            columns = used.Columns;

            return (used.Row, used.Column, used.Row + rows.Count - 1, used.Column + columns.Count - 1);
        }
        finally
        {
            ReleaseComObject(columns);
            ReleaseComObject(rows);
            ReleaseComObject(used);
        }
    }

    /// <summary>
    /// Copies a block of cells from one worksheet onto the same block of another as values alone, a band of rows at a
    /// time.
    /// </summary>
    /// <param name="source">The worksheet to copy from.</param>
    /// <param name="destination">The worksheet to copy onto.</param>
    /// <param name="lastRow">The last row of the block, which starts at the first cell of the worksheet.</param>
    /// <param name="lastColumn">The last column of the block.</param>
    /// <param name="blockRows">The number of rows read and written per pair of calls.</param>
    /// <remarks>
    /// Read and written as arrays rather than copied and pasted, because a paste carries structure as well as values: a
    /// paste over the block a table covers removes that table, and every formula that read it structurally is rewritten
    /// to #REF!.  Assigning values into the cells of a table is ordinary editing, with one exception that matters here:
    /// a single assignment covering a table's whole block, heading row included, replaces the table the same way a paste
    /// does.  The heading row is therefore written on its own, and the rows beneath it in bands, which also keeps what
    /// one call marshals bounded however large the sweep is.
    /// </remarks>
    public static void CopyValues(Excel.Worksheet source, Excel.Worksheet destination, int lastRow, int lastColumn,
        int blockRows)
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(destination);
        ArgumentOutOfRangeException.ThrowIfLessThan(blockRows, 1);

        WriteGrid(destination, 1, 1, ReadGrid(source, 1, 1, 1, lastColumn));

        for (int firstRow = 2; firstRow <= lastRow; firstRow += blockRows)
        {
            int rows = Math.Min(blockRows, lastRow - firstRow + 1);

            WriteGrid(destination, firstRow, 1, ReadGrid(source, firstRow, 1, rows, lastColumn));
        }
    }

    /// <summary>
    /// Replaces every formula on a worksheet with the value it last calculated.
    /// </summary>
    /// <param name="worksheet">The worksheet to flatten.</param>
    /// <remarks>
    /// The paste is onto the range it was copied from, which is how a sheet is flattened without leaving the Excel
    /// process.  Values alone, so the number formats, the conditional formatting and the charts are untouched, and an
    /// error a cell deliberately holds, ex. the #N/A of a blank trial slot, is carried across as that error rather than
    /// as a blank.  Assigning the values of the range to itself would not do: an error reads back as the number Excel
    /// codes it with, and a cell that held #N/A would come to hold -2146826246.
    /// <para>
    /// A caller must have recalculated first, because what is kept is what the cells currently show, and must flatten
    /// the sheets that read a spill before the sheet that spills: a name anchored to a spill refers to nothing once the
    /// spill is a block of values, and every formula reading that name shows #REF! from then on.  A sheet that any table
    /// is defined on is not flattened at all - it holds no formula to flatten, and the paste would remove the table.
    /// </para>
    /// </remarks>
    public static void FlattenToValues(Excel.Worksheet worksheet)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        Excel.Range? used = null;
        Excel.Application? application = null;
        try
        {
            used = worksheet.UsedRange;

            used.Copy();
            used.PasteSpecial(Excel.XlPasteType.xlPasteValues);

            application = worksheet.Application;
            application.CutCopyMode = 0;
        }
        finally
        {
            ReleaseComObject(application);
            ReleaseComObject(used);
        }
    }

    /// <summary>
    /// Reads a block of a worksheet in a single call.
    /// </summary>
    /// <param name="worksheet">The worksheet to read from.</param>
    /// <param name="firstRow">The one based row the block starts at.</param>
    /// <param name="firstColumn">The one based column the block starts at.</param>
    /// <param name="rowCount">The number of rows the block covers.</param>
    /// <param name="columnCount">The number of columns the block covers.</param>
    /// <returns>The values as a zero based grid indexed by row and then by column.  Empty cells are null.</returns>
    /// <remarks>
    /// The counterpart of <see cref="WriteGrid"/>, and one cross process call for the same reason.  A cell holding an
    /// error is read as the integer code of that error, which is what <see cref="IsError"/> recognizes.
    /// </remarks>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    public static object?[,] ReadGrid(Excel.Worksheet worksheet, int firstRow, int firstColumn, int rowCount,
        int columnCount)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        if (rowCount <= 0 || columnCount <= 0)
        {
            return new object?[0, 0];
        }

        Excel.Range? first = null;
        Excel.Range? last = null;
        Excel.Range? block = null;
        try
        {
            first = (Excel.Range)worksheet.Cells[firstRow, firstColumn];
            last = (Excel.Range)worksheet.Cells[firstRow + rowCount - 1, firstColumn + columnCount - 1];
            block = worksheet.Range[first, last];

            return ReadRange(block);
        }
        finally
        {
            ReleaseComObject(block);
            ReleaseComObject(last);
            ReleaseComObject(first);
        }
    }

    /// <summary>
    /// Reads every cell of a range in a single call.
    /// </summary>
    /// <param name="range">The range to read.</param>
    /// <returns>The values as a zero based grid indexed by row and then by column.  Empty cells are null.</returns>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    public static object?[,] ReadRange(Excel.Range range)
    {
        ArgumentNullException.ThrowIfNull(range);

        object? value = range.Value2;

        if (value is not object[,] grid)
        {
            // A single cell range yields the scalar value rather than a two dimensional array.
            return new object?[1, 1] { { value } };
        }

        int firstRow = grid.GetLowerBound(0);
        int firstColumn = grid.GetLowerBound(1);
        int rowCount = grid.GetLength(0);
        int columnCount = grid.GetLength(1);

        var values = new object?[rowCount, columnCount];
        for (int row = 0; row < rowCount; row++)
        {
            for (int column = 0; column < columnCount; column++)
            {
                values[row, column] = grid[firstRow + row, firstColumn + column];
            }
        }

        return values;
    }

    /// <summary>
    /// The error codes Excel returns for the cell errors, keyed by the code and valued by what the cell displays.
    /// </summary>
    /// <remarks>
    /// A cell holding an error marshals as the integer code of that error rather than as text, so a caller reading a
    /// block of cells recognizes an error by the value it reads rather than by asking Excel a second time.
    /// </remarks>
    private static readonly Dictionary<int, string> CellErrors = new()
    {
        // The code of an error is 2000 + its position in Excel's own list, marshalled as -2146826288 + that offset,
        // ex. #DIV/0! is 2007 and arrives as -2146826281.  The list skips 2044, so the errors after #GETTING_DATA are
        // not where counting from the one before them would put them.
        [-2146826288] = "#NULL!",
        [-2146826281] = "#DIV/0!",
        [-2146826273] = "#VALUE!",
        [-2146826265] = "#REF!",
        [-2146826259] = "#NAME?",
        [-2146826252] = "#NUM!",
        [-2146826246] = "#N/A",
        [-2146826245] = "#GETTING_DATA",
        [-2146826243] = "#SPILL!",
        [-2146826242] = "#CONNECT!",
        [-2146826241] = "#BLOCKED!",
        [-2146826240] = "#UNKNOWN!",
        [-2146826239] = "#FIELD!",
        [-2146826238] = "#CALC!"
    };

    /// <summary>
    /// The code of the error a cell shows as #N/A, which the analysis uses deliberately to keep a line off a chart.
    /// </summary>
    public const int NotAvailableError = -2146826246;

    /// <summary>
    /// Indicates whether a value read from a cell is an error, and says which error it is.
    /// </summary>
    /// <param name="value">The value read from the cell.</param>
    /// <param name="error">The error the cell shows, ex. #N/A, or null when the value is not an error.</param>
    /// <returns>True when the value is a cell error.</returns>
    public static bool IsError(object? value, [NotNullWhen(true)] out string? error)
    {
        if (value is int code && CellErrors.TryGetValue(code, out string? text))
        {
            error = text;
            return true;
        }

        error = null;
        return false;
    }

    /// <summary>
    /// Gets the range a defined name refers to.
    /// </summary>
    /// <param name="workbook">The workbook that defines the name.</param>
    /// <param name="definedName">The defined name to resolve.</param>
    /// <returns>The range the name refers to.</returns>
    public static Excel.Range GetNameRange(Excel.Workbook workbook, string definedName)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.Names? names = null;
        Excel.Name? name = null;
        try
        {
            names = workbook.Names;

            try
            {
                name = names.Item(definedName);
                return name.RefersToRange;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("Defined name " + definedName
                    + " was not found in the workbook, or does not refer to a range.", ex);
            }
        }
        finally
        {
            ReleaseComObject(name);
            ReleaseComObject(names);
        }
    }

    /// <summary>
    /// Lists every defined name of the workbook with the formula it refers to.
    /// </summary>
    /// <param name="workbook">The workbook to describe.</param>
    /// <returns>The names, each with what it refers to.</returns>
    /// <remarks>
    /// The list includes the hidden names Excel maintains for the functions a workbook uses, ex. _xlfn.XLOOKUP, because
    /// a caller checking what the names of a workbook refer to is asking about all of them.
    /// </remarks>
    public static IReadOnlyList<(string Name, string RefersTo)> GetNameDefinitions(Excel.Workbook workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        var definitions = new List<(string, string)>();

        Excel.Names? names = null;
        try
        {
            names = workbook.Names;

            for (int index = 1; index <= names.Count; index++)
            {
                Excel.Name? name = null;
                try
                {
                    name = names.Item(index);
                    definitions.Add((name.Name, name.RefersTo as string ?? string.Empty));
                }
                finally
                {
                    ReleaseComObject(name);
                }
            }
        }
        finally
        {
            ReleaseComObject(names);
        }

        return definitions;
    }

    /// <summary>
    /// Deletes a defined name from the workbook.
    /// </summary>
    /// <param name="workbook">The workbook that defines the name.</param>
    /// <param name="definedName">The name to delete.</param>
    public static void DeleteName(Excel.Workbook workbook, string definedName)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.Names? names = null;
        Excel.Name? name = null;
        try
        {
            names = workbook.Names;

            try
            {
                name = names.Item(definedName);
                name.Delete();
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("Defined name " + definedName + " could not be deleted.", ex);
            }
        }
        finally
        {
            ReleaseComObject(name);
            ReleaseComObject(names);
        }
    }

    /// <summary>
    /// Lists the worksheets of the workbook, in the order it holds them.
    /// </summary>
    /// <param name="workbook">The workbook to describe.</param>
    /// <returns>The worksheet names.</returns>
    public static IReadOnlyList<string> GetWorksheetNames(Excel.Workbook workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        var names = new List<string>();

        Excel.Sheets worksheets = workbook.Worksheets;
        try
        {
            for (int index = 1; index <= worksheets.Count; index++)
            {
                Excel.Worksheet? worksheet = null;
                try
                {
                    worksheet = (Excel.Worksheet)worksheets[index];
                    names.Add(worksheet.Name);
                }
                finally
                {
                    ReleaseComObject(worksheet);
                }
            }
        }
        finally
        {
            ReleaseComObject(worksheets);
        }

        return names;
    }

    /// <summary>
    /// Deletes a worksheet from the workbook.
    /// </summary>
    /// <param name="workbook">The workbook that holds the worksheet.</param>
    /// <param name="worksheetName">The name of the worksheet to delete.</param>
    /// <remarks>
    /// Excel asks for confirmation before deleting a sheet that holds anything, so the caller's session must have
    /// suppressed alerts, which <see cref="ExcelSession.SuspendCalculation"/> does.
    /// </remarks>
    public static void DeleteWorksheet(Excel.Workbook workbook, string worksheetName)
    {
        Excel.Worksheet worksheet = GetWorksheet(workbook, worksheetName);
        try
        {
            worksheet.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Worksheet " + worksheetName + " could not be deleted.", ex);
        }
        finally
        {
            ReleaseComObject(worksheet);
        }
    }

    /// <summary>
    /// Reads the distance from the top of a worksheet to the top of each of its first rows, in points.
    /// </summary>
    /// <param name="worksheet">The worksheet to measure.</param>
    /// <param name="rowCount">The number of rows to measure.</param>
    /// <returns>
    /// The tops, indexed by one based row, so that index zero is unused and index rowCount + 1 holds the top of the
    /// row after the last one measured.
    /// </returns>
    /// <remarks>
    /// This is what a check on the placement of a chart is expressed in: a chart is positioned in points and the data
    /// it must not cover is positioned in rows, so one of the two has to be converted into the other's terms.  The
    /// heights are measured rather than assumed, because a row whose height was changed by hand would otherwise move
    /// every row beneath it out from under the check.
    /// </remarks>
    public static double[] GetRowTops(Excel.Worksheet worksheet, int rowCount)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        ArgumentOutOfRangeException.ThrowIfLessThan(rowCount, 1);

        var tops = new double[rowCount + 2];

        for (int row = 1; row <= rowCount + 1; row++)
        {
            Excel.Range? cell = null;
            try
            {
                cell = (Excel.Range)worksheet.Cells[row, 1];
                tops[row] = (double)cell.Top;
            }
            finally
            {
                ReleaseComObject(cell);
            }
        }

        return tops;
    }

    /// <summary>
    /// Describes every chart on a worksheet by the name it carries and the box it occupies, in points.
    /// </summary>
    /// <param name="worksheet">The worksheet whose charts are wanted.</param>
    /// <returns>The charts, in the order the worksheet holds them.</returns>
    public static IReadOnlyList<(string Name, double Top, double Height)> GetChartBoxes(Excel.Worksheet worksheet)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        var boxes = new List<(string, double, double)>();

        Excel.ChartObjects charts = (Excel.ChartObjects)worksheet.ChartObjects();
        try
        {
            for (int index = 1; index <= charts.Count; index++)
            {
                Excel.ChartObject? chart = null;
                try
                {
                    chart = (Excel.ChartObject)charts.Item(index);
                    boxes.Add((chart.Name, chart.Top, chart.Height));
                }
                finally
                {
                    ReleaseComObject(chart);
                }
            }
        }
        finally
        {
            ReleaseComObject(charts);
        }

        return boxes;
    }

    /// <summary>
    /// Closes a workbook without saving it, and releases the reference to it.
    /// </summary>
    /// <param name="workbook">The workbook to close.  Null is ignored.</param>
    /// <remarks>
    /// This is for a workbook a session opened beside its own, ex. the one a transplant reads.  Nothing is saved
    /// because nothing was written: a workbook opened to be read from is closed the way it was found.
    /// </remarks>
    public static void CloseWorkbook(Excel.Workbook? workbook)
    {
        if (workbook is null)
        {
            return;
        }

        try
        {
            workbook.Close(SaveChanges: false);
        }
        catch (COMException ex)
        {
            Log.Logger.Warning("PFM_EXCEL_COMPANION_CLOSE_FAILED: " + ex.Message);
        }
        finally
        {
            ReleaseComObject(workbook);
        }
    }

    /// <summary>
    /// Releases one reference to a runtime callable wrapper.  Every helper that acquires an intermediate COM object
    /// releases it here so that Excel can shut down promptly.
    /// </summary>
    /// <param name="comObject">The COM object to release.  Null is ignored.</param>
    /// <remarks>
    /// This deliberately decrements the reference count by one rather than calling
    /// <see cref="Marshal.FinalReleaseComObject(object)"/>.  Excel hands back the same wrapper for objects with a stable
    /// identity, such as the Application, so releasing a wrapper outright could invalidate a reference that is still in
    /// use elsewhere in the session.
    /// </remarks>
    public static void ReleaseComObject(object? comObject)
    {
        if (comObject is null)
        {
            return;
        }

        try
        {
            if (Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        catch (ArgumentException)
        {
            // The object was not a runtime callable wrapper after all.  There is nothing to release.
        }
    }

    /// <summary>
    /// Searches the running object table for the workbook at the specified path.
    /// </summary>
    /// <param name="fullPath">The fully qualified path of the workbook.</param>
    /// <returns>The workbook if it is already open in some Excel instance; otherwise null.</returns>
    private static Excel.Workbook? TryGetOpenWorkbook(string fullPath)
    {
        string fileName = Path.GetFileName(fullPath);
        Excel.Workbook? match = null;

        foreach (object comObject in EnumerateRunningObjects())
        {
            if (match is null)
            {
                try
                {
                    if (comObject is Excel.Workbook workbook)
                    {
                        match = IsSameWorkbook(workbook, fullPath, fileName) ? workbook : null;
                    }
                    else if (comObject is Excel.Application excel)
                    {
                        match = FindWorkbook(excel, fullPath, fileName);
                    }
                }
                catch (COMException)
                {
                    // The registration is stale, or the object is not an Excel object we can interrogate.
                }
            }

            // Keep the wrapper the match came from; the session needs it to reach the owning Excel instance.
            if (match is null)
            {
                ReleaseComObject(comObject);
            }
        }

        return match;
    }

    /// <summary>
    /// Searches an Excel instance for the workbook at the specified path.
    /// </summary>
    /// <param name="excel">The Excel instance to search.</param>
    /// <param name="fullPath">The fully qualified path of the workbook.</param>
    /// <param name="fileName">The file name portion of that path.</param>
    /// <returns>The matching workbook, or null when the instance does not have it open.</returns>
    private static Excel.Workbook? FindWorkbook(Excel.Application excel, string fullPath, string fileName)
    {
        Excel.Workbooks workbooks = excel.Workbooks;
        try
        {
            foreach (Excel.Workbook workbook in workbooks)
            {
                if (IsSameWorkbook(workbook, fullPath, fileName))
                {
                    return workbook;
                }

                ReleaseComObject(workbook);
            }

            return null;
        }
        finally
        {
            ReleaseComObject(workbooks);
        }
    }

    /// <summary>
    /// Determines whether an open workbook is the one at the specified path.
    /// </summary>
    /// <param name="workbook">The candidate workbook.</param>
    /// <param name="fullPath">The fully qualified path of the workbook being sought.</param>
    /// <param name="fileName">The file name portion of that path.</param>
    /// <returns>True when the candidate is the workbook being sought.</returns>
    /// <remarks>
    /// A workbook that Excel opened through a cloud service reports a URL as its FullName rather than the synced local
    /// path, so a URL falls back to matching on the file name alone.
    /// </remarks>
    private static bool IsSameWorkbook(Excel.Workbook workbook, string fullPath, string fileName)
    {
        try
        {
            string name = workbook.FullName;

            if (string.Equals(name, fullPath, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            bool isUrl = name.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

            return isUrl && string.Equals(workbook.Name, fileName, StringComparison.OrdinalIgnoreCase);
        }
        catch (COMException)
        {
            return false;
        }
    }

    /// <summary>
    /// Enumerates the objects currently registered in the running object table.
    /// </summary>
    /// <returns>The registered objects.  The caller owns every wrapper it does not keep.</returns>
    private static List<object> EnumerateRunningObjects()
    {
        var objects = new List<object>();

        IBindCtx? bindContext = null;
        IRunningObjectTable? runningObjectTable = null;
        IEnumMoniker? monikers = null;
        try
        {
            if (CreateBindCtx(0, out bindContext) != 0 || bindContext is null)
            {
                return objects;
            }

            bindContext.GetRunningObjectTable(out runningObjectTable);
            if (runningObjectTable is null)
            {
                return objects;
            }

            runningObjectTable.EnumRunning(out monikers);
            if (monikers is null)
            {
                return objects;
            }

            var buffer = new IMoniker[1];
            while (monikers.Next(1, buffer, IntPtr.Zero) == 0)
            {
                IMoniker moniker = buffer[0];
                try
                {
                    if (runningObjectTable.GetObject(moniker, out object comObject) == 0 && comObject != null)
                    {
                        objects.Add(comObject);
                    }
                }
                catch (COMException)
                {
                    // The registration is stale.  Skip it.
                }
                finally
                {
                    ReleaseComObject(moniker);
                }
            }
        }
        catch (COMException ex)
        {
            Log.Logger.Warning("PFM_EXCEL_ROT_ENUMERATION_FAILED: " + ex.Message);
        }
        finally
        {
            ReleaseComObject(monikers);
            ReleaseComObject(runningObjectTable);
            ReleaseComObject(bindContext);
        }

        return objects;
    }

    /// <summary>
    /// Gets the process ids of the Excel instances that are currently running.
    /// </summary>
    /// <returns>The process ids.</returns>
    private static HashSet<int> GetExcelProcessIds()
    {
        var processIds = new HashSet<int>();

        foreach (Process process in Process.GetProcessesByName(ExcelProcessName))
        {
            try
            {
                processIds.Add(process.Id);
            }
            finally
            {
                process.Dispose();
            }
        }

        return processIds;
    }

    /// <summary>
    /// Resolves the id of the process hosting an Excel instance.
    /// </summary>
    /// <param name="excel">The Excel instance.</param>
    /// <returns>The process id, or zero when it could not be resolved.</returns>
    /// <remarks>
    /// Excel exposes no process id of its own, so it is taken from the window the instance reports as its main window.
    /// That window belongs to the instance's process by construction.
    /// </remarks>
    internal static int GetProcessId(Excel.Application excel)
    {
        ArgumentNullException.ThrowIfNull(excel);

        try
        {
            IntPtr handle = new(excel.Hwnd);
            if (handle == IntPtr.Zero)
            {
                return 0;
            }

            _ = GetWindowThreadProcessId(handle, out int processId);
            return processId;
        }
        catch (COMException)
        {
            return 0;
        }
    }

    /// <summary>
    /// Quits an Excel instance, logging rather than throwing when it refuses.
    /// </summary>
    /// <param name="excel">The Excel instance to quit.</param>
    internal static void TryQuit(Excel.Application excel)
    {
        ArgumentNullException.ThrowIfNull(excel);

        try
        {
            excel.Quit();
        }
        catch (COMException ex)
        {
            Log.Logger.Warning("PFM_EXCEL_SHUTDOWN_FAILED: " + ex.Message);
        }
    }

    /// <summary>
    /// Ends an Excel process that did not exit when it was asked to quit.
    /// </summary>
    /// <param name="processId">The id of the process the session started.</param>
    /// <remarks>
    /// Excel outlives a quit whenever a reference to one of its objects survives, and a sweep that leaks one instance
    /// per job would exhaust the machine long before it finished.  Only a process this tool started and identified is
    /// ever ended this way, and only after it has been asked to quit and given time to do so.
    /// </remarks>
    internal static void EndProcess(int processId)
    {
        Process process;
        try
        {
            process = Process.GetProcessById(processId);
        }
        catch (ArgumentException)
        {
            // The process already exited, which is the outcome this method exists to reach.
            return;
        }

        try
        {
            if (process.HasExited)
            {
                return;
            }

            Log.Logger.Warning("PFM_EXCEL_PROCESS_KILLED: process=" + Format(processId)
                + " did not exit after being asked to quit.");
            process.Kill();
            process.WaitForExit(ProcessExitTimeoutMilliseconds);
        }
        catch (InvalidOperationException)
        {
            // The process exited between the checks above and the kill.
        }
        catch (System.ComponentModel.Win32Exception ex)
        {
            Log.Logger.Warning("PFM_EXCEL_PROCESS_KILL_FAILED: " + ex.Message);
        }
        finally
        {
            process.Dispose();
        }
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

    /// <summary>
    /// The name of the Excel process, as the process list reports it.
    /// </summary>
    private const string ExcelProcessName = "EXCEL";

    /// <summary>
    /// How long to wait for an Excel process to disappear after it has been ended.
    /// </summary>
    private const int ProcessExitTimeoutMilliseconds = 10_000;

    [DllImport("ole32.dll")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern int CreateBindCtx(int reserved, out IBindCtx bindContext);

    [DllImport("user32.dll")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern int GetWindowThreadProcessId(IntPtr windowHandle, out int processId);
}

/// <summary>
/// Owns an Excel application and workbook reference for the duration of an operation.  The session records the
/// application settings it changes so that they can be restored, and it only closes the workbook and quits Excel when it
/// was the one that started them.
/// </summary>
public sealed class ExcelSession : IDisposable
{
    private readonly Excel.Application _excel;
    private readonly Excel.Workbook _workbook;
    private readonly bool _startedExcel;
    private readonly int _processId;
    private readonly Excel.XlCalculation _originalCalculation;
    private readonly bool _originalScreenUpdating;
    private readonly bool _originalEnableEvents;
    private readonly bool _originalDisplayAlerts;

    private bool _disposed;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelSession"/> class.
    /// </summary>
    /// <param name="excel">The Excel application that hosts the workbook.</param>
    /// <param name="workbook">The workbook the session operates on.</param>
    /// <param name="startedExcel">True when this session started Excel and opened the workbook itself.</param>
    /// <param name="processId">
    /// The id of the Excel process, when the session started it and knows which process that is; otherwise zero.  A
    /// session that knows its process ends it if quitting leaves it running.
    /// </param>
    internal ExcelSession(Excel.Application excel, Excel.Workbook workbook, bool startedExcel, int processId = 0)
    {
        _excel = excel;
        _workbook = workbook;
        _startedExcel = startedExcel;
        _processId = processId;

        _originalCalculation = excel.Calculation;
        _originalScreenUpdating = excel.ScreenUpdating;
        _originalEnableEvents = excel.EnableEvents;
        _originalDisplayAlerts = excel.DisplayAlerts;

        OleMessageFilter.Register();
    }

    /// <summary>
    /// Gets the workbook the session operates on.
    /// </summary>
    public Excel.Workbook Workbook => _workbook;

    /// <summary>
    /// Indicates whether the workbook was already open in Excel when the session attached to it.
    /// </summary>
    public bool WasAlreadyOpen => !_startedExcel;

    /// <summary>
    /// Gets the id of the Excel process this session started, or zero when the session did not start Excel or could
    /// not identify its process.
    /// </summary>
    public int ProcessId => _processId;

    /// <summary>
    /// Switches Excel to manual calculation and suppresses the screen updates, events and alerts that would otherwise
    /// fire while an operation writes to the workbook.  The previous settings are restored when the session is disposed.
    /// </summary>
    public void SuspendCalculation()
    {
        _excel.Calculation = Excel.XlCalculation.xlCalculationManual;
        _excel.ScreenUpdating = false;
        _excel.EnableEvents = false;
        _excel.DisplayAlerts = false;
    }

    /// <summary>
    /// Opens a second workbook in the Excel instance this session owns.
    /// </summary>
    /// <param name="filePath">The path to the workbook.</param>
    /// <param name="readOnly">True to open it read only, which is what a caller that only reads it asks for.</param>
    /// <returns>The opened workbook, which the caller closes with <see cref="ExcelUtils.CloseWorkbook"/>.</returns>
    /// <remarks>
    /// Two workbooks that are to be copied between have to be open in one instance, because a copy between instances
    /// goes through the system clipboard rather than through Excel.  The companion is the caller's to close, and it is
    /// closed before the session's own workbook is recalculated: a full calculation recalculates every workbook the
    /// instance holds open.
    /// </remarks>
    public Excel.Workbook OpenCompanionWorkbook(string filePath, bool readOnly)
    {
        ArgumentNullException.ThrowIfNull(filePath);

        string fullPath = Path.GetFullPath(filePath);

        Excel.Workbooks workbooks = _excel.Workbooks;
        try
        {
            Excel.Workbook opened = workbooks.Open(Filename: fullPath, UpdateLinks: 0, ReadOnly: readOnly);
            Log.Logger.Information("PFM_EXCEL_COMPANION_OPENED: " + fullPath);
            return opened;
        }
        finally
        {
            ExcelUtils.ReleaseComObject(workbooks);
        }
    }

    /// <summary>
    /// Recalculates the formulas that have been marked dirty since the last calculation.
    /// </summary>
    public void Calculate()
    {
        _excel.Calculate();
    }

    /// <summary>
    /// Recalculates every formula in every open workbook, whether or not it is marked dirty.
    /// </summary>
    public void CalculateFull()
    {
        _excel.CalculateFull();
    }

    /// <summary>
    /// Saves the workbook in place, first returning the calculation mode to what it was when the session started.
    /// </summary>
    /// <remarks>
    /// Excel persists the application calculation mode into the workbook, so saving while calculation is suspended
    /// would leave the workbook in manual calculation the next time it is opened.  Restoring the mode first also means
    /// the values being saved have been calculated the way the workbook expects.
    /// </remarks>
    public void Save()
    {
        Restore(() => _excel.Calculation = _originalCalculation);

        _workbook.Save();
        Log.Logger.Information("PFM_EXCEL_SAVED: " + _workbook.FullName + " (already open in Excel: "
            + WasAlreadyOpen + ")");
    }

    /// <summary>
    /// Saves the workbook to a path it has not been saved to before, first returning the calculation mode to what it
    /// was when the session started.
    /// </summary>
    /// <param name="path">The path to save the workbook to.</param>
    /// <remarks>
    /// The format is stated rather than inferred, because a workbook Excel created has no format of its own to keep
    /// and would otherwise be saved in whatever the installation defaults to.  Excel persists the application
    /// calculation mode into the workbook, so the mode is restored first for the same reason <see cref="Save"/>
    /// restores it: a workbook saved while calculation is suspended opens in manual calculation.
    /// </remarks>
    public void SaveAs(string path)
    {
        ArgumentNullException.ThrowIfNull(path);

        Restore(() => _excel.Calculation = _originalCalculation);

        _workbook.SaveAs(Filename: Path.GetFullPath(path),
            FileFormat: Excel.XlFileFormat.xlOpenXMLWorkbook);

        Log.Logger.Information("PFM_EXCEL_SAVED_AS: " + _workbook.FullName);
    }

    /// <summary>
    /// Restores the Excel application settings the session changed and, when the session started Excel, closes the
    /// workbook and quits.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        RestoreApplicationSettings();
        OleMessageFilter.Revoke();

        if (_startedExcel)
        {
            CloseAndQuit();
        }

        ExcelUtils.ReleaseComObject(_workbook);
        ExcelUtils.ReleaseComObject(_excel);

        if (_startedExcel)
        {
            // The Excel process only exits once the wrappers this session handed out have been collected.
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (_processId != 0)
            {
                ExcelUtils.EndProcess(_processId);
            }
        }
    }

    /// <summary>
    /// Restores each application setting the session changed.  A failure to restore one setting must not prevent the
    /// rest from being restored, so each is attempted independently.
    /// </summary>
    private void RestoreApplicationSettings()
    {
        Restore(() => _excel.Calculation = _originalCalculation);
        Restore(() => _excel.ScreenUpdating = _originalScreenUpdating);
        Restore(() => _excel.EnableEvents = _originalEnableEvents);
        Restore(() => _excel.DisplayAlerts = _originalDisplayAlerts);
    }

    /// <summary>
    /// Applies a restore action, logging rather than throwing when Excel refuses it.
    /// </summary>
    /// <param name="restore">The restore action to apply.</param>
    private static void Restore(Action restore)
    {
        try
        {
            restore();
        }
        catch (COMException ex)
        {
            Log.Logger.Warning("PFM_EXCEL_RESTORE_FAILED: " + ex.Message);
        }
    }

    /// <summary>
    /// Closes the workbook this session opened and quits Excel when no other workbook remains open in the instance.
    /// </summary>
    /// <remarks>
    /// Any save the operation wanted has already happened through <see cref="Save"/>, so the workbook is closed without
    /// saving.  The remaining workbook count guards against quitting an instance that turned out to be shared.
    /// </remarks>
    private void CloseAndQuit()
    {
        try
        {
            _workbook.Close(SaveChanges: false);

            Excel.Workbooks workbooks = _excel.Workbooks;
            try
            {
                if (workbooks.Count == 0)
                {
                    _excel.Quit();
                }
            }
            finally
            {
                ExcelUtils.ReleaseComObject(workbooks);
            }
        }
        catch (COMException ex)
        {
            Log.Logger.Warning("PFM_EXCEL_SHUTDOWN_FAILED: " + ex.Message);
        }
    }
}

/// <summary>
/// An OLE message filter for the calling single threaded apartment.  Excel rejects incoming automation calls while it is
/// busy, for instance while a cell is in edit mode; without a filter those rejections surface as RPC_E_CALL_REJECTED.
/// With one registered, the rejected call is retried until Excel accepts it or the retry window elapses.
/// </summary>
internal sealed class OleMessageFilter : IOleMessageFilter
{
    private const int ServerCallIsHandled = 0;
    private const int ServerCallRetryLater = 2;
    private const int PendingMessageWaitDefaultProcess = 2;
    private const int CancelCall = -1;
    private const int RetryDelayMilliseconds = 250;
    private const int RetryWindowMilliseconds = 120_000;

    private static OleMessageFilter? _registered;

    /// <summary>
    /// Registers the filter for the calling apartment.  Registration is only supported on a single threaded apartment;
    /// elsewhere it is logged and skipped, leaving rejected calls to surface as exceptions.
    /// </summary>
    public static void Register()
    {
        var filter = new OleMessageFilter();
        int hr = CoRegisterMessageFilter(filter, out _);

        if (hr != 0)
        {
            Log.Logger.Warning("PFM_EXCEL_MESSAGE_FILTER_NOT_REGISTERED: 0x" + Hresult(hr));
            return;
        }

        _registered = filter;
    }

    /// <summary>
    /// Removes the filter previously registered for the calling apartment.
    /// </summary>
    public static void Revoke()
    {
        if (_registered is null)
        {
            return;
        }

        int hr = CoRegisterMessageFilter(null, out _);
        if (hr != 0)
        {
            Log.Logger.Warning("PFM_EXCEL_MESSAGE_FILTER_NOT_REVOKED: 0x" + Hresult(hr));
        }

        _registered = null;
    }

    /// <summary>
    /// Handles an incoming call.  This process makes no outgoing calls that Excel calls back into, so every incoming
    /// call is accepted.
    /// </summary>
    /// <param name="callType">The type of the incoming call.</param>
    /// <param name="taskCaller">A handle to the calling task.</param>
    /// <param name="tickCount">The elapsed tick count.</param>
    /// <param name="interfaceInfo">Information about the interface being called.</param>
    /// <returns>SERVERCALL_ISHANDLED.</returns>
    public int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo)
    {
        return ServerCallIsHandled;
    }

    /// <summary>
    /// Decides what to do when Excel rejects a call.
    /// </summary>
    /// <param name="taskCallee">A handle to the task that rejected the call.</param>
    /// <param name="tickCount">The number of milliseconds elapsed since the call was made.</param>
    /// <param name="rejectType">Whether the call was rejected outright or merely deferred.</param>
    /// <returns>The number of milliseconds to wait before retrying, or -1 to cancel the call.</returns>
    public int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType)
    {
        if (rejectType == ServerCallRetryLater && tickCount < RetryWindowMilliseconds)
        {
            return RetryDelayMilliseconds;
        }

        return CancelCall;
    }

    /// <summary>
    /// Decides what to do while a call this process made is still outstanding.
    /// </summary>
    /// <param name="taskCallee">A handle to the task being called.</param>
    /// <param name="tickCount">The number of milliseconds elapsed since the call was made.</param>
    /// <param name="pendingType">The type of the pending call.</param>
    /// <returns>PENDINGMSG_WAITDEFPROCESS.</returns>
    public int MessagePending(IntPtr taskCallee, int tickCount, int pendingType)
    {
        return PendingMessageWaitDefaultProcess;
    }

    /// <summary>
    /// Formats an HRESULT for a log message.
    /// </summary>
    /// <param name="hr">The HRESULT to format.</param>
    /// <returns>The HRESULT as eight hexadecimal digits.</returns>
    private static string Hresult(int hr)
    {
        return hr.ToString("X8", CultureInfo.InvariantCulture);
    }

    [DllImport("ole32.dll")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);
}

/// <summary>
/// The OLE message filter interface, IID_IMessageFilter.  It is declared here because it is not exposed by the .NET
/// interop assemblies.
/// </summary>
[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    /// <summary>
    /// Handles an incoming call.
    /// </summary>
    /// <param name="callType">The type of the incoming call.</param>
    /// <param name="taskCaller">A handle to the calling task.</param>
    /// <param name="tickCount">The elapsed tick count.</param>
    /// <param name="interfaceInfo">Information about the interface being called.</param>
    /// <returns>A SERVERCALL value.</returns>
    [PreserveSig]
    int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo);

    /// <summary>
    /// Decides what to do when a call is rejected.
    /// </summary>
    /// <param name="taskCallee">A handle to the task that rejected the call.</param>
    /// <param name="tickCount">The number of milliseconds elapsed since the call was made.</param>
    /// <param name="rejectType">A SERVERCALL value describing the rejection.</param>
    /// <returns>The retry delay in milliseconds, or -1 to cancel the call.</returns>
    [PreserveSig]
    int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType);

    /// <summary>
    /// Decides what to do while an outgoing call is outstanding.
    /// </summary>
    /// <param name="taskCallee">A handle to the task being called.</param>
    /// <param name="tickCount">The number of milliseconds elapsed since the call was made.</param>
    /// <param name="pendingType">The type of the pending call.</param>
    /// <returns>A PENDINGMSG value.</returns>
    [PreserveSig]
    int MessagePending(IntPtr taskCallee, int tickCount, int pendingType);
}
