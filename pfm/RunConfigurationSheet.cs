using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pfm;

/// <summary>
/// The configuration of one sweep as a worksheet holds it: the synopsis as prose at the top, then each structured
/// table of the plan, labelled and defined as an Excel table, stacked beneath it with a blank row between.
/// </summary>
/// <remarks>
/// A sweep records its configuration once and it is read back twice: by a person, and by the analysis, which resolves
/// the era schedule through structured references to these tables.  The writer, the updater and the reader of the layout
/// all live here, because the three have to move together: coalesce writes this sheet onto a blank one, apply-analysis
/// reads it back out of that workbook and updates the same sheet in a copy of the analysis template.
/// </remarks>
public static class RunConfigurationSheet
{
    /// <summary>
    /// The number of rows between the synopsis and the heading row of the first table: one blank row, then the label
    /// naming the table, then the heading row itself.
    /// </summary>
    private const int RowsBeforeFirstTableHeader = 3;

    /// <summary>
    /// Writes a sweep's configuration onto a blank worksheet, defining a table over each block.
    /// </summary>
    /// <param name="worksheet">The blank worksheet to write.</param>
    /// <param name="synopsis">The lines of the synopsis, in order.</param>
    /// <param name="tables">The structured tables, in the order they are stacked down the sheet.</param>
    public static void Write(Excel.Worksheet worksheet, IReadOnlyList<string> synopsis,
        IReadOnlyList<ConfigurationTable> tables)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        ArgumentNullException.ThrowIfNull(synopsis);
        ArgumentNullException.ThrowIfNull(tables);

        WriteSynopsis(worksheet, synopsis);

        int labelRow = synopsis.Count == 0 ? 1 : synopsis.Count + 2;

        foreach (ConfigurationTable table in tables)
        {
            int headerRow = labelRow + 1;
            int bodyRows = WriteTable(worksheet, table, labelRow);

            ExcelUtils.AddTable(worksheet, table.Name, headerRow, 1, bodyRows + 1, table.Header.Count);

            labelRow = headerRow + bodyRows + 2;
        }
    }

    /// <summary>
    /// Updates a worksheet that already holds a configuration so that it holds this sweep's instead, leaving the tables
    /// defined over it in place.
    /// </summary>
    /// <param name="worksheet">The worksheet to update.</param>
    /// <param name="synopsis">The lines of the synopsis, in order.</param>
    /// <param name="tables">The structured tables, in the order they are stacked down the sheet.</param>
    /// <exception cref="InvalidOperationException">
    /// The worksheet carries no table of its own for one of this sweep's configuration tables, so there is nothing to
    /// resize; or a table of this sweep no longer carries a column the table on the worksheet does, so the analysis
    /// reading that column by name would be reading something else.
    /// </exception>
    /// <remarks>
    /// The tables are resized rather than replaced, and the blocks are moved by inserting and deleting whole worksheet
    /// rows, because both leave a table's identity and every structured reference to it intact.  Replacing a table would
    /// not: a table that is dissolved has each formula that read it rewritten into the cell range it happened to occupy,
    /// and one that is deleted leaves them as #REF!.  Everything else about the sheet, ex. the width of its columns,
    /// belongs to the workbook this is being written into rather than to the sweep.
    /// </remarks>
    public static void Update(Excel.Worksheet worksheet, IReadOnlyList<string> synopsis,
        IReadOnlyList<ConfigurationTable> tables)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        ArgumentNullException.ThrowIfNull(synopsis);
        ArgumentNullException.ThrowIfNull(tables);

        MoveFirstTable(worksheet, synopsis.Count);

        ExcelUtils.ClearRowContents(worksheet, 1, synopsis.Count + 1);
        WriteSynopsis(worksheet, synopsis);

        int lastRow = synopsis.Count;
        int lastColumn = 1;

        foreach (ConfigurationTable table in tables)
        {
            Excel.ListObject listObject = GetTableBeingWritten(worksheet, table.Name);
            try
            {
                (int headerRow, int firstColumn, int rowCount, _) = ExcelUtils.GetTableExtent(listObject);

                CheckColumns(listObject, table);

                int bodyRows = Math.Max(table.Rows.Count, 1);
                int currentBodyRows = rowCount - 1;

                // Inside the body, so that Excel takes the inserted rows into the table and moves the tables below it
                // down.  A table always has a body row, so there is always a row to insert against.
                ExcelUtils.InsertRows(worksheet, headerRow + 1, bodyRows - currentBodyRows);
                ExcelUtils.DeleteRows(worksheet, headerRow + 1, currentBodyRows - bodyRows);

                ExcelUtils.ResizeTable(listObject, headerRow, firstColumn, bodyRows + 1, table.Header.Count);

                WriteTable(worksheet, table, headerRow - 1);

                lastRow = headerRow + bodyRows;
                lastColumn = Math.Max(lastColumn, table.Header.Count);
            }
            finally
            {
                ExcelUtils.ReleaseComObject(listObject);
            }
        }

        ExcelUtils.ClearBeyond(worksheet, lastRow, lastColumn);
    }

    /// <summary>
    /// Reads a sweep's configuration back off a worksheet that holds it.
    /// </summary>
    /// <param name="worksheet">The worksheet to read.</param>
    /// <returns>The synopsis and the structured tables, in <see cref="RunConfiguration.TableNames"/> order.</returns>
    /// <exception cref="InvalidOperationException">
    /// The worksheet does not hold every table the configuration is made of.  A workbook that records a sweep without
    /// them states its plan in prose alone, which a reader can read and the analysis cannot: the era schedule it
    /// resolves the income target and floor from is one of these tables.
    /// </exception>
    /// <remarks>
    /// The tables are read by name and the synopsis is measured from where the first of them sits, so a plan whose
    /// tables have grown or shrunk since is read as it stands rather than as it was when the layout was designed.
    /// </remarks>
    public static SweepConfiguration Read(Excel.Worksheet worksheet)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        var tables = new List<ConfigurationTable>(RunConfiguration.TableNames.Count);
        int firstHeaderRow = int.MaxValue;

        foreach (string tableName in RunConfiguration.TableNames)
        {
            Excel.ListObject listObject = GetRecordedTable(worksheet, tableName);
            try
            {
                (int headerRow, _, _, _) = ExcelUtils.GetTableExtent(listObject);
                firstHeaderRow = Math.Min(firstHeaderRow, headerRow);

                tables.Add(ReadTable(listObject, tableName));
            }
            finally
            {
                ExcelUtils.ReleaseComObject(listObject);
            }
        }

        int synopsisLines = firstHeaderRow - RowsBeforeFirstTableHeader;

        return new SweepConfiguration(ReadSynopsis(worksheet, synopsisLines), tables);
    }

    /// <summary>
    /// Moves the stack of tables up or down the worksheet so that a synopsis of the given length fits above it.
    /// </summary>
    /// <param name="worksheet">The worksheet being updated.</param>
    /// <param name="synopsisLines">The number of lines the new synopsis occupies.</param>
    /// <remarks>
    /// The rows are inserted at the top of the sheet, above everything, so that every table moves as a whole and none
    /// of them is stretched by the move.
    /// </remarks>
    private static void MoveFirstTable(Excel.Worksheet worksheet, int synopsisLines)
    {
        Excel.ListObject listObject = GetTableBeingWritten(worksheet, RunConfiguration.TableNames[0]);
        try
        {
            (int headerRow, _, _, _) = ExcelUtils.GetTableExtent(listObject);
            int wantedHeaderRow = synopsisLines + RowsBeforeFirstTableHeader;

            ExcelUtils.InsertRows(worksheet, 1, wantedHeaderRow - headerRow);
            ExcelUtils.DeleteRows(worksheet, 1, headerRow - wantedHeaderRow);
        }
        finally
        {
            ExcelUtils.ReleaseComObject(listObject);
        }
    }

    /// <summary>
    /// Writes the synopsis into the top of the worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet being written.</param>
    /// <param name="lines">The lines of the synopsis.</param>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    private static void WriteSynopsis(Excel.Worksheet worksheet, IReadOnlyList<string> lines)
    {
        if (lines.Count == 0)
        {
            return;
        }

        // The synopsis is a report rather than data: its cells are formatted as text before anything is written, so a
        // line reading like a date or beginning with an equals sign is stored as the workbook wrote it.  Only those
        // cells, because the tables below hold numbers, and a number written into a text formatted cell is text.
        ExcelUtils.FormatRangeAsText(worksheet, 1, 1, lines.Count, 1);

        var block = new object?[lines.Count, 1];
        for (int row = 0; row < lines.Count; row++)
        {
            block[row, 0] = lines[row];
        }

        ExcelUtils.WriteGrid(worksheet, 1, 1, block);
    }

    /// <summary>
    /// Writes one configuration table's label, headings and rows into the worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet being written.</param>
    /// <param name="table">The table to write.</param>
    /// <param name="labelRow">The one based row the label naming the table goes on.</param>
    /// <returns>The number of body rows written, which is one for a table the plan left empty.</returns>
    /// <remarks>
    /// A table the plan left empty is written as its headings and one blank row, which is the smallest table Excel
    /// holds, and which is what distinguishes a plan that directs no draws from a sweep whose configuration was never
    /// recorded.
    /// </remarks>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    private static int WriteTable(Excel.Worksheet worksheet, ConfigurationTable table, int labelRow)
    {
        var label = new object?[1, 1] { { table.Name } };
        ExcelUtils.WriteGrid(worksheet, labelRow, 1, label);
        ExcelUtils.SetCellBold(worksheet, labelRow, 1);

        int headerRow = labelRow + 1;

        var headings = new object?[1, table.Header.Count];
        for (int column = 0; column < table.Header.Count; column++)
        {
            headings[0, column] = table.Header[column];
        }

        ExcelUtils.WriteGrid(worksheet, headerRow, 1, headings);

        if (table.Rows.Count == 0)
        {
            return 1;
        }

        var values = new object?[table.Rows.Count, table.Header.Count];
        for (int row = 0; row < table.Rows.Count; row++)
        {
            object?[] source = table.Rows[row];
            for (int column = 0; column < table.Header.Count && column < source.Length; column++)
            {
                values[row, column] = source[column];
            }
        }

        ExcelUtils.WriteGrid(worksheet, headerRow + 1, 1, values);

        return table.Rows.Count;
    }

    /// <summary>
    /// Checks that a sweep's table still carries every column the table on the worksheet does, in the same order.
    /// </summary>
    /// <param name="listObject">The table on the worksheet being updated.</param>
    /// <param name="table">The sweep's table, which is about to be written over it.</param>
    /// <remarks>
    /// The analysis reads these columns by name.  Writing a heading row that has dropped or reordered one of them would
    /// leave Excel to rename the column, and every formula that referred to it would follow the rename onto whatever is
    /// now in that position.  Columns added at the end of a plan's table are another matter, and are accepted: the names
    /// the analysis reads are unchanged by them.
    /// </remarks>
    private static void CheckColumns(Excel.ListObject listObject, ConfigurationTable table)
    {
        string[] current = ExcelUtils.GetColumnNames(listObject);

        for (int column = 0; column < current.Length; column++)
        {
            string wanted = current[column];
            string found = column < table.Header.Count ? table.Header[column] : string.Empty;

            if (!string.Equals(wanted, found, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Column " + (column + 1).ToString(CultureInfo.InvariantCulture)
                    + " of the " + table.Name + " table of this sweep is "
                    + (found.Length == 0 ? "missing" : found) + " where the workbook being written expects " + wanted
                    + ".  The analysis reads these columns by name, so a plan whose table has changed shape is a change "
                    + "to the analysis template rather than a sweep it can be applied to.");
            }
        }
    }

    /// <summary>
    /// Reads the synopsis off the top of the worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet being read.</param>
    /// <param name="lineCount">The number of lines the synopsis occupies.</param>
    /// <returns>The lines, in order.</returns>
    private static string[] ReadSynopsis(Excel.Worksheet worksheet, int lineCount)
    {
        if (lineCount <= 0)
        {
            return [];
        }

        object?[,] block = ExcelUtils.ReadGrid(worksheet, 1, 1, lineCount, 1);
        var lines = new string[lineCount];

        for (int row = 0; row < lineCount; row++)
        {
            lines[row] = Convert.ToString(block[row, 0], CultureInfo.InvariantCulture) ?? string.Empty;
        }

        return lines;
    }

    /// <summary>
    /// Reads one configuration table off the worksheet.
    /// </summary>
    /// <param name="listObject">The table to read.</param>
    /// <param name="tableName">The name the table is recorded under.</param>
    /// <returns>The table's headings and rows.</returns>
    /// <remarks>
    /// A row of nothing but empty cells is the blank row that stands in for a table the plan left empty, which is read
    /// back as the empty table it represents rather than as a row of nulls.
    /// </remarks>
    private static ConfigurationTable ReadTable(Excel.ListObject listObject, string tableName)
    {
        string[] header = ExcelUtils.GetColumnNames(listObject);
        object?[,] grid = ExcelUtils.ReadTableRows(listObject);

        int rowCount = grid.GetLength(0);
        int columnCount = grid.GetLength(1);
        var rows = new List<object?[]>(rowCount);

        for (int row = 0; row < rowCount; row++)
        {
            var values = new object?[columnCount];
            bool empty = true;

            for (int column = 0; column < columnCount; column++)
            {
                values[column] = grid[row, column];
                empty = empty && values[column] is null;
            }

            if (!empty)
            {
                rows.Add(values);
            }
        }

        return new ConfigurationTable(tableName, header, rows);
    }

    /// <summary>
    /// Gets a configuration table of a worksheet a sweep was recorded onto.
    /// </summary>
    /// <param name="worksheet">The worksheet that hosts the table.</param>
    /// <param name="tableName">The name of the table.</param>
    /// <returns>The table.</returns>
    private static Excel.ListObject GetRecordedTable(Excel.Worksheet worksheet, string tableName)
    {
        return GetRequiredTable(worksheet, tableName,
            "A workbook that records a sweep carries one table per configuration table of the plan, and the analysis "
            + "resolves the income target and floor through them.");
    }

    /// <summary>
    /// Gets a configuration table of a worksheet this sweep is being written onto.
    /// </summary>
    /// <param name="worksheet">The worksheet that hosts the table.</param>
    /// <param name="tableName">The name of the table.</param>
    /// <returns>The table.</returns>
    /// <remarks>
    /// This is the other side of the same layout and it fails for a different reason, so it says so: the worksheet here
    /// is the analysis template's, which is authored by hand and can therefore be a table behind the plan the sweeps
    /// now record.
    /// </remarks>
    private static Excel.ListObject GetTableBeingWritten(Excel.Worksheet worksheet, string tableName)
    {
        return GetRequiredTable(worksheet, tableName,
            "The tables of this sheet are resized rather than replaced, so the workbook being written carries one of "
            + "its own for every configuration table the sweep records and there is nothing to resize without it.  A "
            + "table the plan has gained since is added to the analysis template before a sweep carrying it can be "
            + "applied.");
    }

    /// <summary>
    /// Gets a configuration table of a worksheet by name.
    /// </summary>
    /// <param name="worksheet">The worksheet that hosts the table.</param>
    /// <param name="tableName">The name of the table.</param>
    /// <param name="explanation">What a missing table means on this worksheet, and what to do about it.</param>
    /// <returns>The table.</returns>
    private static Excel.ListObject GetRequiredTable(Excel.Worksheet worksheet, string tableName, string explanation)
    {
        try
        {
            return ExcelUtils.GetTable(worksheet, tableName);
        }
        catch (InvalidOperationException ex)
        {
            throw new InvalidOperationException("The configuration table " + tableName + " was not found on worksheet "
                + worksheet.Name + ".  " + explanation, ex);
        }
    }
}

/// <summary>
/// The configuration of one sweep as it was read back off a worksheet: the synopsis and the structured tables.
/// </summary>
public sealed class SweepConfiguration
{
    /// <summary>
    /// Initializes a new instance of the <see cref="SweepConfiguration"/> class.
    /// </summary>
    /// <param name="synopsis">The lines of the synopsis, in order.</param>
    /// <param name="tables">The structured tables, in <see cref="RunConfiguration.TableNames"/> order.</param>
    public SweepConfiguration(IReadOnlyList<string> synopsis, IReadOnlyList<ConfigurationTable> tables)
    {
        Synopsis = synopsis;
        Tables = tables;
    }

    /// <summary>
    /// Gets the lines of the synopsis, in order.
    /// </summary>
    public IReadOnlyList<string> Synopsis { get; }

    /// <summary>
    /// Gets the structured tables, in <see cref="RunConfiguration.TableNames"/> order.
    /// </summary>
    public IReadOnlyList<ConfigurationTable> Tables { get; }
}
