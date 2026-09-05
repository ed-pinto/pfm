using System.Diagnostics;
using System.Globalization;
using Serilog;

namespace Pfm;

/// <summary>
/// Reports how long the phases of an operation take.  An operation wraps each phase it wants measured in a scope taken
/// from <see cref="Measure"/>; when timing was not requested every scope is inert, so an ordinary run pays nothing for
/// the instrumentation beyond the scope itself.
/// </summary>
/// <remarks>
/// The measurements exist to locate the cost of a slow run, which for this workbook is dominated by the file I/O that
/// Excel performs when it opens and saves it.  Each report goes to both the log and the console so that a run can be
/// diagnosed from either.  The log entries are written at the verbose level, which keeps the measurements out of the
/// log of an ordinary run unless the configured minimum level admits them.
/// </remarks>
public sealed class DiagnosticTimer
{
    /// <summary>
    /// The scope handed out when timing is off.  It is stateless, so one instance serves every caller.
    /// </summary>
    private static readonly IDisposable InertScope = new NullScope();

    private readonly bool _enabled;
    private readonly TextWriter _output;
    private readonly Stopwatch _total;

    /// <summary>
    /// Initializes a new instance of the <see cref="DiagnosticTimer"/> class and starts measuring the total elapsed
    /// time of the operation.
    /// </summary>
    /// <param name="enabled">True when the caller asked for diagnostic timing.</param>
    /// <param name="output">The writer that receives the measurements.</param>
    public DiagnosticTimer(bool enabled, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(output);

        _enabled = enabled;
        _output = output;
        _total = Stopwatch.StartNew();
    }

    /// <summary>
    /// Indicates whether measurements are being collected and reported.
    /// </summary>
    public bool Enabled => _enabled;

    /// <summary>
    /// Begins measuring a phase of the operation.  Disposing the returned scope reports how long the phase took.
    /// </summary>
    /// <param name="phase">The name of the phase, ex. open workbook.</param>
    /// <returns>The scope covering the phase.</returns>
    public IDisposable Measure(string phase)
    {
        return _enabled ? new Scope(this, phase) : InertScope;
    }

    /// <summary>
    /// Reports a measurement the caller took itself, ex. one that spans work this timer did not scope.
    /// </summary>
    /// <param name="phase">The name of the phase.</param>
    /// <param name="elapsed">The time the phase took.</param>
    public void Report(string phase, TimeSpan elapsed)
    {
        if (!_enabled)
        {
            return;
        }

        string message = "PFM_TIMING: " + phase + " " + Format(elapsed);
        Log.Logger.Verbose(message);
        _output.WriteLine(message);
    }

    /// <summary>
    /// Reports the time elapsed since this timer was created.  An operation calls this once, as it finishes.
    /// </summary>
    public void ReportTotal()
    {
        Report("total", _total.Elapsed);
    }

    /// <summary>
    /// Formats an elapsed time for a log or console message.
    /// </summary>
    /// <param name="elapsed">The elapsed time to format.</param>
    /// <returns>The elapsed time in seconds, to the millisecond.</returns>
    private static string Format(TimeSpan elapsed)
    {
        return elapsed.TotalSeconds.ToString("F3", CultureInfo.InvariantCulture) + " s";
    }

    /// <summary>
    /// Measures one phase of an operation.  Disposal reports the elapsed time to the timer that issued the scope.
    /// </summary>
    private sealed class Scope : IDisposable
    {
        private readonly DiagnosticTimer _timer;
        private readonly string _phase;
        private readonly Stopwatch _stopwatch;

        private bool _disposed;

        /// <summary>
        /// Initializes a new instance of the <see cref="Scope"/> class and starts measuring.
        /// </summary>
        /// <param name="timer">The timer that receives the measurement.</param>
        /// <param name="phase">The name of the phase being measured.</param>
        internal Scope(DiagnosticTimer timer, string phase)
        {
            _timer = timer;
            _phase = phase;
            _stopwatch = Stopwatch.StartNew();
        }

        /// <summary>
        /// Stops measuring and reports the elapsed time.
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _stopwatch.Stop();
            _timer.Report(_phase, _stopwatch.Elapsed);
        }
    }

    /// <summary>
    /// The scope used when timing is off.  It measures nothing and reports nothing.
    /// </summary>
    private sealed class NullScope : IDisposable
    {
        /// <summary>
        /// Does nothing.
        /// </summary>
        public void Dispose()
        {
        }
    }
}
