using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;

namespace Pfm;

/// <summary>
/// Reads back the results file one simulation job wrote, a block of rows at a time.
/// </summary>
/// <remarks>
/// The counterpart of <see cref="SimulationResultWriter"/>.  The writer records numbers and flags in a form that reads
/// back as the same values, so this converts each field to the type it was written from rather than handing Excel
/// text: a coalesced workbook is there to be charted and filtered, which text in a numeric column would prevent.
/// <para>
/// A sweep can be long enough that its rows are worth reading in blocks rather than all at once, so the file is
/// streamed and the caller decides how much of it to hold at a time.
/// </para>
/// </remarks>
public sealed class SimulationResultReader : IDisposable
{
    private readonly StreamReader _reader;
    private readonly string _path;

    private int _line;
    private bool _disposed;

    /// <summary>
    /// Initializes a new instance of the <see cref="SimulationResultReader"/> class and reads the header row.
    /// </summary>
    /// <param name="path">The path of the results file to read.</param>
    /// <exception cref="InvalidOperationException">The file holds no header row.</exception>
    public SimulationResultReader(string path)
    {
        ArgumentNullException.ThrowIfNull(path);

        _path = path;

        // Detecting the byte order mark rather than assuming one: the writer emits none, and a file that has acquired
        // one, ex. from having been opened and saved by something else, would otherwise start with it inside the first
        // column heading.
        _reader = new StreamReader(path, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            detectEncodingFromByteOrderMarks: true);

        string[]? header = ReadRow();

        Header = header
            ?? throw new InvalidOperationException("The results file " + path + " is empty.  A finished job writes a "
                + "header row before its first simulation.");
    }

    /// <summary>
    /// Gets the column headings of the file, in column order.
    /// </summary>
    public IReadOnlyList<string> Header { get; }

    /// <summary>
    /// Reads the next block of data rows.
    /// </summary>
    /// <param name="maxRows">The greatest number of rows to read.</param>
    /// <returns>
    /// The values as a zero based grid indexed by row and then by column, or null once the file is exhausted.  The
    /// last block of a file may hold fewer than <paramref name="maxRows"/> rows.
    /// </returns>
    /// <exception cref="InvalidOperationException">A row does not have as many fields as the header.</exception>
    [SuppressMessage("Performance", "CA1814:Prefer jagged arrays over multidimensional",
        Justification = "Excel marshals a multi cell range as a rectangular variant array.")]
    public object?[,]? ReadBlock(int maxRows)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentOutOfRangeException.ThrowIfLessThan(maxRows, 1);

        var rows = new List<string[]>(maxRows);

        while (rows.Count < maxRows && ReadRow() is string[] row)
        {
            if (row.Length != Header.Count)
            {
                throw new InvalidOperationException("Line " + Format(_line) + " of " + _path + " has "
                    + Format(row.Length) + " fields rather than the " + Format(Header.Count)
                    + " its header declares.  The file is truncated or damaged.");
            }

            rows.Add(row);
        }

        if (rows.Count == 0)
        {
            return null;
        }

        var values = new object?[rows.Count, Header.Count];
        for (int row = 0; row < rows.Count; row++)
        {
            for (int column = 0; column < Header.Count; column++)
            {
                values[row, column] = Convert(rows[row][column]);
            }
        }

        return values;
    }

    /// <summary>
    /// Closes the file.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _reader.Dispose();
    }

    /// <summary>
    /// Reads the next non empty line as fields.
    /// </summary>
    /// <returns>The fields, or null at the end of the file.</returns>
    /// <remarks>
    /// Blank lines are skipped rather than read as a row of one empty field, which is what the trailing newline of the
    /// last row would otherwise produce on a reader less forgiving than this one.
    /// </remarks>
    private string[]? ReadRow()
    {
        while (_reader.ReadLine() is string text)
        {
            _line++;

            if (text.Length > 0)
            {
                return SplitFields(text);
            }
        }

        return null;
    }

    /// <summary>
    /// Splits a line into its fields.
    /// </summary>
    /// <param name="line">The line to split.</param>
    /// <returns>The fields.</returns>
    /// <remarks>
    /// A sweep writes plain numbers and flags, none of which need quoting, so this exists to be right about a file
    /// that has been through something that does quote, ex. a spreadsheet that re-saved it.  A field spanning a line
    /// break is not supported, because nothing that writes these files produces one.
    /// </remarks>
    private static string[] SplitFields(string line)
    {
        var fields = new List<string>();
        var field = new StringBuilder();
        bool quoted = false;

        for (int position = 0; position < line.Length; position++)
        {
            char character = line[position];

            if (quoted)
            {
                if (character != '"')
                {
                    field.Append(character);
                }
                else if (position + 1 < line.Length && line[position + 1] == '"')
                {
                    // A doubled quote inside a quoted field is one quote.
                    field.Append('"');
                    position++;
                }
                else
                {
                    quoted = false;
                }
            }
            else if (character == '"' && field.Length == 0)
            {
                quoted = true;
            }
            else if (character == ',')
            {
                fields.Add(field.ToString());
                field.Clear();
            }
            else
            {
                field.Append(character);
            }
        }

        fields.Add(field.ToString());
        return [.. fields];
    }

    /// <summary>
    /// Converts a field to the value it was written from.
    /// </summary>
    /// <param name="field">The field to convert.</param>
    /// <returns>A number, a flag, null for an empty field, or the text itself when it is none of those.</returns>
    /// <remarks>
    /// The order matters: TRUE and FALSE are tested before numbers because Excel would otherwise be handed the text of
    /// a flag, and a field that is neither is passed through as text rather than failing the import, which is the same
    /// stance the synopsis takes towards a line it did not expect.
    /// </remarks>
    private static object? Convert(string field)
    {
        if (field.Length == 0)
        {
            return null;
        }

        if (string.Equals(field, "TRUE", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (string.Equals(field, "FALSE", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        // Excel holds no infinity or NaN, so a field that parses as one is left as the text it already is rather than
        // becoming a value the workbook cannot represent.
        if (double.TryParse(field, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
            && double.IsFinite(number))
        {
            return number;
        }

        return field;
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
