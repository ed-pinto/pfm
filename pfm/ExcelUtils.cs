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
    /// Formats a column of a worksheet as text, so that Excel stores what is written into it verbatim.
    /// </summary>
    /// <param name="worksheet">The worksheet whose column is being formatted.</param>
    /// <param name="column">The one based column to format.</param>
    /// <remarks>
    /// A line of prose that happens to begin with an equals sign, or that reads like a date, would otherwise be taken
    /// as a formula or converted as it was written.  The synopsis is a report to be read back exactly as the workbook
    /// wrote it, so the column it lands in says so before anything is written to it.
    /// </remarks>
    public static void FormatColumnAsText(Excel.Worksheet worksheet, int column)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        Excel.Range? columns = null;
        Excel.Range? target = null;
        try
        {
            columns = worksheet.Columns;
            target = (Excel.Range)columns[column];
            target.NumberFormat = "@";
        }
        finally
        {
            ReleaseComObject(target);
            ReleaseComObject(columns);
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

        Excel.Range range = table.DataBodyRange
            ?? throw new InvalidOperationException("Table " + table.Name + " has no data rows.");
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
