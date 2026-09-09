using System.Globalization;
using Serilog;

namespace Pfm;

/// <summary>
/// The slice of a sweep that one job runs.
/// </summary>
/// <remarks>
/// A sweep is split across jobs by handing each a contiguous run of the simulations, so a job never has to know what
/// any other job is doing and the slices reassemble by concatenation.  The slices differ in length by at most one when
/// the sweep does not divide evenly.
/// </remarks>
public readonly struct JobPartition : IEquatable<JobPartition>
{
    /// <summary>
    /// Initializes a new instance of the <see cref="JobPartition"/> struct.
    /// </summary>
    /// <param name="startOffset">The zero based index of the first simulation of the slice.</param>
    /// <param name="count">The number of simulations in the slice.</param>
    private JobPartition(int startOffset, int count)
    {
        StartOffset = startOffset;
        Count = count;
    }

    /// <summary>
    /// Gets the zero based index, within the whole sweep, of the first simulation this job runs.
    /// </summary>
    public int StartOffset { get; }

    /// <summary>
    /// Gets the number of simulations this job runs.  Zero when there are fewer simulations than jobs.
    /// </summary>
    public int Count { get; }

    /// <summary>
    /// Divides a sweep among the jobs running it and yields the slice one of them owns.
    /// </summary>
    /// <param name="total">The number of simulations in the whole sweep.</param>
    /// <param name="jobCount">The number of jobs the sweep is split across.</param>
    /// <param name="jobIndex">The zero based index of the job whose slice is wanted.</param>
    /// <returns>The slice that job owns.</returns>
    public static JobPartition Create(int total, int jobCount, int jobIndex)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(total);
        ArgumentOutOfRangeException.ThrowIfLessThan(jobCount, 1);
        ArgumentOutOfRangeException.ThrowIfNegative(jobIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(jobIndex, jobCount);

        // The first remainder jobs each take one extra simulation, so the slices stay contiguous and cover the sweep.
        int share = total / jobCount;
        int remainder = total % jobCount;

        int count = share + (jobIndex < remainder ? 1 : 0);
        int startOffset = (jobIndex * share) + Math.Min(jobIndex, remainder);

        return new JobPartition(startOffset, count);
    }

    /// <inheritdoc/>
    public bool Equals(JobPartition other)
    {
        return StartOffset == other.StartOffset && Count == other.Count;
    }

    /// <inheritdoc/>
    public override bool Equals(object? obj)
    {
        return obj is JobPartition other && Equals(other);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return HashCode.Combine(StartOffset, Count);
    }

    /// <summary>
    /// Determines whether two partitions are the same slice.
    /// </summary>
    /// <param name="left">The first partition.</param>
    /// <param name="right">The second partition.</param>
    /// <returns>True when they are the same slice.</returns>
    public static bool operator ==(JobPartition left, JobPartition right)
    {
        return left.Equals(right);
    }

    /// <summary>
    /// Determines whether two partitions are different slices.
    /// </summary>
    /// <param name="left">The first partition.</param>
    /// <param name="right">The second partition.</param>
    /// <returns>True when they are different slices.</returns>
    public static bool operator !=(JobPartition left, JobPartition right)
    {
        return !left.Equals(right);
    }
}

/// <summary>
/// The working directory of one simulation job: its own copy of the workbook and the file its results are written to.
/// </summary>
/// <remarks>
/// A sweep never drives the workbook the user pointed at.  Every job copies it and drives the copy, for two reasons:
/// the sweep writes throwaway parameter values that must not be saved back, and the jobs of a parallel sweep must not
/// contend for one file.
/// <para>
/// The copy is working state rather than a result, so disposing the run deletes it and leaves the results file behind
/// in the directory.  A workbook this size costs tens of megabytes a job, which a sweep run often enough would
/// otherwise accumulate silently.
/// </para>
/// </remarks>
public sealed class SimulationRun : IDisposable
{
    private const string DirectoryPrefix = "pfm.";
    private const string TimestampFormat = "yyyyMMddHHmm";

    /// <summary>
    /// The number of times the copy is deleted before the attempt is given up on.  Excel releases the file as its
    /// process ends, which can lag the close by a moment.
    /// </summary>
    private const int DeleteAttempts = 3;

    /// <summary>
    /// How long to wait between attempts to delete the copy.
    /// </summary>
    private const int DeleteRetryMilliseconds = 250;

    private bool _disposed;

    /// <summary>
    /// Initializes a new instance of the <see cref="SimulationRun"/> class.
    /// </summary>
    /// <param name="directory">The directory the job works in.</param>
    /// <param name="workbookPath">The path of the job's copy of the workbook.</param>
    /// <param name="resultsPath">The path of the file the job's results are written to.</param>
    private SimulationRun(string directory, string workbookPath, string resultsPath)
    {
        Directory = directory;
        WorkbookPath = workbookPath;
        ResultsPath = resultsPath;
    }

    /// <summary>
    /// Gets the directory the job works in.
    /// </summary>
    public string Directory { get; }

    /// <summary>
    /// Gets the path of the job's own copy of the workbook.
    /// </summary>
    public string WorkbookPath { get; }

    /// <summary>
    /// Gets the path of the file the job's results are written to.
    /// </summary>
    public string ResultsPath { get; }

    /// <summary>
    /// Creates the job's working directory and copies the workbook into it.
    /// </summary>
    /// <param name="sourceFilePath">The path of the workbook to simulate.</param>
    /// <param name="resultsName">What the results file is called, ex. backtest.</param>
    /// <param name="jobCount">The number of jobs the sweep is split across.</param>
    /// <param name="jobIndex">The zero based index of this job.</param>
    /// <returns>The prepared run.</returns>
    /// <remarks>
    /// The directory is created under the current directory rather than beside the source workbook, so that a sweep
    /// leaves nothing behind where the real workbook lives.  The job index and count are part of both names, so the
    /// jobs of one sweep neither collide nor need to be told apart afterwards by their contents.
    /// </remarks>
    public static SimulationRun Create(string sourceFilePath, string resultsName, int jobCount, int jobIndex)
    {
        ArgumentNullException.ThrowIfNull(sourceFilePath);
        ArgumentNullException.ThrowIfNull(resultsName);

        string sourceFullPath = Path.GetFullPath(sourceFilePath);
        string stamp = DateTime.Now.ToString(TimestampFormat, CultureInfo.InvariantCulture);
        string job = Format(jobIndex) + "_of_" + Format(jobCount);

        string directory = Path.Combine(System.IO.Directory.GetCurrentDirectory(),
            DirectoryPrefix + stamp + "." + job);

        // The jobs of one sweep start within the same minute often enough that the timestamp alone would collide, so
        // the job suffix is what makes the name unique.  Copying without overwrite then refuses to reuse a directory
        // that already holds a workbook, rather than quietly running a second sweep on top of a finished one.
        System.IO.Directory.CreateDirectory(directory);

        string workbookPath = Path.Combine(directory, Path.GetFileName(sourceFullPath));
        File.Copy(sourceFullPath, workbookPath, overwrite: false);

        string resultsPath = Path.Combine(directory, resultsName + "." + stamp + "." + job + ".csv");

        Log.Logger.Information("PFM_SIMULATION_RUN: directory=" + directory + " workbook=" + workbookPath
            + " results=" + resultsPath);

        return new SimulationRun(directory, workbookPath, resultsPath);
    }

    /// <summary>
    /// Deletes the job's copy of the workbook.
    /// </summary>
    /// <remarks>
    /// The session that drove the copy must be disposed first, because Excel holds the file open until it is.  The
    /// directory itself is left in place: the results file the job wrote lives in it and is what the sweep is for.
    /// </remarks>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        for (int attempt = 1; attempt <= DeleteAttempts; attempt++)
        {
            try
            {
                File.Delete(WorkbookPath);
                Log.Logger.Information("PFM_SIMULATION_RUN_CLEANED: workbook=" + WorkbookPath);
                return;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // A copy that cannot be deleted is left behind rather than failing a job whose results are already
                // written.  The last attempt is the one that reports it, so a copy Excel simply had not released yet
                // is not reported at all.
                if (attempt == DeleteAttempts)
                {
                    Log.Logger.Warning("PFM_SIMULATION_RUN_CLEANUP_FAILED: workbook=" + WorkbookPath + " error="
                        + ex.Message);
                    return;
                }

                Thread.Sleep(DeleteRetryMilliseconds);
            }
        }
    }

    /// <summary>
    /// Formats a count for a file or directory name.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <returns>The formatted value.</returns>
    private static string Format(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }
}
