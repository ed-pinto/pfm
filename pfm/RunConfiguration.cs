using System.Globalization;
using System.Text;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pfm;

/// <summary>
/// The synopsis of the workbook a simulation exercised, recorded beside that simulation's results.
/// </summary>
/// <remarks>
/// A sweep drives a throwaway copy of the workbook and deletes it when the job ends, so the results file would
/// otherwise be all that survives of a run whose numbers only mean anything read against the plan that produced them.
/// The synopsis is maintained on 00_Overview as one line of plain text per configured parameter, which is what lets a
/// run record its configuration by copying it rather than by describing the model a second time in this tool, where it
/// would have to be kept in step with the workbook.
/// <para>
/// Every operation that drives a copy of the workbook records it the same way, so a back test and a Monte Carlo run
/// leave the same file, named the same, in their run directories.
/// </para>
/// </remarks>
public static class RunConfiguration
{
    /// <summary>
    /// The name of the file the synopsis is written to.
    /// </summary>
    public const string FileName = "RunConfiguration.txt";

    private const string SynopsisWorksheet = "00_Overview";
    private const string SynopsisTable = "Synopsis";
    private const string SynopsisColumn = "Line";

    /// <summary>
    /// Reads the synopsis of a workbook, one entry per row of the table in table order.
    /// </summary>
    /// <param name="workbook">The workbook whose configuration is wanted.</param>
    /// <returns>The lines of the synopsis.</returns>
    /// <remarks>
    /// Every row of the table is a formula, so what this returns is whatever the workbook last calculated.  A caller
    /// that has suspended calculation must therefore recalculate before reading, or it records the values the workbook
    /// was saved with rather than the ones the run will be driven from.
    /// </remarks>
    public static string[] Read(Excel.Workbook workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        Excel.ListObject synopsis = ExcelUtils.GetListObject(workbook, SynopsisWorksheet, SynopsisTable);
        try
        {
            object?[] values = ExcelUtils.ReadColumn(synopsis, SynopsisColumn);
            return Array.ConvertAll(values, FormatLine);
        }
        finally
        {
            ExcelUtils.ReleaseComObject(synopsis);
        }
    }

    /// <summary>
    /// Reads the synopsis of a workbook and writes it into a directory as <see cref="FileName"/>.
    /// </summary>
    /// <param name="workbook">The workbook whose configuration is being recorded.</param>
    /// <param name="directory">The directory to write the file into, ex. a job's run directory.</param>
    /// <returns>The path of the file that was written.</returns>
    public static string Write(Excel.Workbook workbook, string directory)
    {
        ArgumentNullException.ThrowIfNull(directory);

        string[] lines = Read(workbook);
        string path = Path.Combine(directory, FileName);

        // A byte order mark, unlike the results file, which has none: the synopsis carries whatever punctuation the
        // workbook wrote into it, en dashes included, and the mark is what makes a plain text viewer read those bytes
        // as UTF-8 rather than as the machine's ANSI code page.  WriteAllLines ends each line the way the platform
        // does, which on Windows is the CRLF such a viewer expects.
        File.WriteAllLines(path, lines, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

        Log.Logger.Information("PFM_RUN_CONFIGURATION: path=" + path + " lines=" + Format(lines.Length));

        return path;
    }

    /// <summary>
    /// Formats one cell of the synopsis as a line of the file.
    /// </summary>
    /// <param name="value">The value the cell holds.  Null when the cell is empty.</param>
    /// <returns>The line.</returns>
    /// <remarks>
    /// A synopsis cell is text, because every row of the table is a formula that concatenates one.  A cell that holds
    /// something else, ex. one whose formula returned an error, is written as Excel returned it rather than failing a
    /// run over one line of a report.
    /// </remarks>
    private static string FormatLine(object? value)
    {
        return value switch
        {
            null => string.Empty,
            string text => text,
            bool flag => flag ? "TRUE" : "FALSE",
            double number => number.ToString("R", CultureInfo.InvariantCulture),
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
        };
    }

    /// <summary>
    /// Formats a count for a log message.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <returns>The formatted value.</returns>
    private static string Format(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }
}
