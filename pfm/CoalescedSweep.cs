using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Serilog;

namespace Pfm;

/// <summary>
/// One job's contribution to a sweep, as it was left on disk: the directory the job worked in and the results file it
/// wrote there.
/// </summary>
public sealed class SweepJob
{
    /// <summary>
    /// Initializes a new instance of the <see cref="SweepJob"/> class.
    /// </summary>
    /// <param name="directory">The path of the job's run directory.</param>
    /// <param name="jobIndex">The zero based index of the job within the sweep.</param>
    /// <param name="jobCount">The number of jobs the sweep was split across.</param>
    /// <param name="resultsPath">The path of the results file the job wrote.</param>
    internal SweepJob(string directory, int jobIndex, int jobCount, string resultsPath)
    {
        Directory = directory;
        JobIndex = jobIndex;
        JobCount = jobCount;
        ResultsPath = resultsPath;
    }

    /// <summary>
    /// Gets the path of the job's run directory.
    /// </summary>
    public string Directory { get; }

    /// <summary>
    /// Gets the name of the job's run directory, ex. pfm.202609091624.0_of_8.
    /// </summary>
    public string Name => Path.GetFileName(Directory);

    /// <summary>
    /// Gets the zero based index of the job within the sweep.
    /// </summary>
    public int JobIndex { get; }

    /// <summary>
    /// Gets the number of jobs the sweep was split across.
    /// </summary>
    public int JobCount { get; }

    /// <summary>
    /// Gets the path of the results file the job wrote.
    /// </summary>
    public string ResultsPath { get; }
}

/// <summary>
/// A complete sweep as it was left on disk: every job's run directory, the configuration they were all driven from,
/// and the identifier the coalesced workbook is named for.
/// </summary>
/// <remarks>
/// A sweep is only meaningful whole.  The jobs cover contiguous slices of one range of simulations, so a missing job is
/// a gap in the middle of the results rather than a shorter run, and there is nothing in a results file that would say
/// so afterwards.  Discovery therefore insists on the full set rather than coalescing whatever it finds.
/// </remarks>
public sealed class CoalescedSweep
{
    /// <summary>
    /// The results file each job writes, matched by the name of the simulation that produced it.  A sweep names its
    /// files after the command that ran, which is lower case; the canonical spelling is what the coalesced workbook
    /// names its sheet after.
    /// </summary>
    private static readonly Dictionary<string, string> TestTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["backtest"] = "BackTest",
        ["montecarlo"] = "MonteCarlo"
    };

    /// <summary>
    /// The run directories of a sweep, ex. pfm.202609091624.3_of_8.  The jobs of one sweep stamp their directories as
    /// they start, so two of them that straddle a minute boundary carry different stamps; it is the job count that
    /// identifies the sweep, and the stamp is only part of what makes the names unique.
    /// </summary>
    private static readonly Regex DirectoryPattern = new(@"^pfm\..*\.(?<index>\d+)_of_(?<count>\d+)$",
        RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture);

    /// <summary>
    /// The results file in a run directory, ex. backtest.202609091624.3_of_8.csv.
    /// </summary>
    private static readonly Regex ResultsPattern = new(@"^(?<type>[^.]+)\..*\.(?<index>\d+)_of_(?<count>\d+)\.csv$",
        RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture);

    /// <summary>
    /// Initializes a new instance of the <see cref="CoalescedSweep"/> class.
    /// </summary>
    /// <param name="jobs">The jobs of the sweep, in job index order.</param>
    /// <param name="testType">The canonical name of the simulation the sweep ran.</param>
    /// <param name="runConfiguration">The lines of the configuration the sweep was driven from.</param>
    /// <param name="runId">The identifier the coalesced workbook is named for.</param>
    private CoalescedSweep(IReadOnlyList<SweepJob> jobs, string testType, string[] runConfiguration, string runId)
    {
        Jobs = jobs;
        TestType = testType;
        RunConfiguration = runConfiguration;
        RunId = runId;
    }

    /// <summary>
    /// Gets the jobs of the sweep, in job index order.  Concatenating their results in this order reproduces what a
    /// single job would have written.
    /// </summary>
    public IReadOnlyList<SweepJob> Jobs { get; }

    /// <summary>
    /// Gets the canonical name of the simulation the sweep ran, ex. BackTest.
    /// </summary>
    public string TestType { get; }

    /// <summary>
    /// Gets the lines of the configuration the sweep was driven from, as the first job recorded it.
    /// </summary>
    public IReadOnlyList<string> RunConfiguration { get; }

    /// <summary>
    /// Gets the eight character identifier of this sweep, which the coalesced workbook is named for.
    /// </summary>
    public string RunId { get; }

    /// <summary>
    /// Gets the name of the worksheet the results are imported into, ex. BackTestData.
    /// </summary>
    public string DataWorksheetName => TestType + "Data";

    /// <summary>
    /// Finds the run directories of a sweep under a directory and checks that they are a complete set.
    /// </summary>
    /// <param name="inputPath">The directory holding the run directories.</param>
    /// <returns>The discovered sweep.</returns>
    /// <exception cref="InvalidOperationException">
    /// The directories are not one complete sweep, ex. a job is missing, more than one sweep is present, or a run
    /// directory does not hold the files a finished job leaves.
    /// </exception>
    public static CoalescedSweep Discover(string inputPath)
    {
        ArgumentNullException.ThrowIfNull(inputPath);

        string fullPath = Path.GetFullPath(inputPath);
        List<SweepJob> jobs = FindJobs(fullPath);

        string[] runConfiguration = ReadRunConfiguration(jobs[0]);
        string runId = ComputeRunId(jobs, runConfiguration);
        string testType = TestTypes[TestTypeOf(jobs[0])];

        Log.Logger.Information("PFM_COALESCE_SWEEP: input=" + fullPath + " testType=" + testType + " jobs="
            + Format(jobs.Count) + " runId=" + runId);

        return new CoalescedSweep(jobs, testType, runConfiguration, runId);
    }

    /// <summary>
    /// Finds the run directories under a directory and checks that they are one complete sweep.
    /// </summary>
    /// <param name="fullPath">The fully qualified path of the directory holding the run directories.</param>
    /// <returns>The jobs of the sweep, in job index order.</returns>
    private static List<SweepJob> FindJobs(string fullPath)
    {
        var found = new List<(int Index, int Count, string Directory)>();

        foreach (string directory in System.IO.Directory.EnumerateDirectories(fullPath))
        {
            Match match = DirectoryPattern.Match(Path.GetFileName(directory));
            if (!match.Success)
            {
                continue;
            }

            // The pattern only matches digits, so the parse can only fail on a number too large to be one, which is
            // not a run directory this tool wrote either.
            if (int.TryParse(match.Groups["index"].Value, NumberStyles.None, CultureInfo.InvariantCulture,
                    out int index)
                && int.TryParse(match.Groups["count"].Value, NumberStyles.None, CultureInfo.InvariantCulture,
                    out int count))
            {
                found.Add((index, count, directory));
            }
        }

        if (found.Count == 0)
        {
            throw new InvalidOperationException("No run directories were found in " + fullPath
                + ".  A sweep leaves one directory per job, named pfm.<stamp>.<index>_of_<count>.");
        }

        int[] counts = found.Select(job => job.Count).Distinct().Order().ToArray();
        if (counts.Length > 1)
        {
            throw new InvalidOperationException("The run directories in " + fullPath + " are from more than one sweep: "
                + "they report job counts of " + string.Join(", ", counts.Select(Format))
                + ".  Coalesce the sweeps separately.");
        }

        int jobCount = counts[0];
        CheckComplete(found, jobCount, fullPath);

        return found
            .OrderBy(job => job.Index)
            .Select(job => new SweepJob(job.Directory, job.Index, jobCount, FindResults(job.Directory, job.Index,
                jobCount)))
            .ToList();
    }

    /// <summary>
    /// Checks that the discovered directories hold each job of the sweep exactly once.
    /// </summary>
    /// <param name="found">The directories that were discovered.</param>
    /// <param name="jobCount">The number of jobs the sweep was split across.</param>
    /// <param name="fullPath">The directory the run directories were found in, for the message.</param>
    private static void CheckComplete(List<(int Index, int Count, string Directory)> found, int jobCount,
        string fullPath)
    {
        var missing = new List<int>();
        var duplicated = new List<string>();

        for (int index = 0; index < jobCount; index++)
        {
            int matches = found.Count(job => job.Index == index);

            if (matches == 0)
            {
                missing.Add(index);
            }
            else if (matches > 1)
            {
                duplicated.Add(Format(index) + "_of_" + Format(jobCount));
            }
        }

        // An index at or beyond the job count cannot belong to a sweep of this size, so it is reported rather than
        // ignored: it means the directory holds the remains of something other than the sweep being coalesced.
        string[] foreign = found
            .Where(job => job.Index >= jobCount)
            .Select(job => Path.GetFileName(job.Directory))
            .Order(StringComparer.Ordinal)
            .ToArray();

        if (missing.Count > 0)
        {
            throw new InvalidOperationException("The sweep in " + fullPath + " is incomplete: it is split across "
                + Format(jobCount) + " jobs, and the run "
                + (missing.Count == 1 ? "directory for job " : "directories for jobs ")
                + string.Join(", ", missing.Select(index => Format(index) + "_of_" + Format(jobCount)))
                + " " + (missing.Count == 1 ? "is" : "are") + " not there.");
        }

        if (duplicated.Count > 0)
        {
            throw new InvalidOperationException("The directory " + fullPath + " holds more than one run directory for "
                + (duplicated.Count == 1 ? "job " : "jobs ") + string.Join(", ", duplicated)
                + ".  Coalesce the sweeps separately.");
        }

        if (foreign.Length > 0)
        {
            throw new InvalidOperationException("The directory " + fullPath + " holds run "
                + (foreign.Length == 1 ? "directory " : "directories ") + string.Join(", ", foreign)
                + ", whose job "
                + (foreign.Length == 1 ? "index is" : "indexes are")
                + " beyond the " + Format(jobCount) + " jobs of this sweep.");
        }
    }

    /// <summary>
    /// Finds the results file a job wrote in its run directory.
    /// </summary>
    /// <param name="directory">The job's run directory.</param>
    /// <param name="jobIndex">The zero based index of the job.</param>
    /// <param name="jobCount">The number of jobs the sweep was split across.</param>
    /// <returns>The path of the results file.</returns>
    private static string FindResults(string directory, int jobIndex, int jobCount)
    {
        string job = Format(jobIndex) + "_of_" + Format(jobCount);

        string[] candidates = System.IO.Directory.GetFiles(directory, "*.csv")
            .Where(path => IsResultsOf(path, job))
            .Order(StringComparer.Ordinal)
            .ToArray();

        if (candidates.Length == 0)
        {
            throw new InvalidOperationException("The run directory " + directory + " holds no results file.  A "
                + "finished job leaves one named <test type>.<stamp>." + job + ".csv, ex. backtest.  A job that is "
                + "still running has not written one yet.");
        }

        if (candidates.Length > 1)
        {
            throw new InvalidOperationException("The run directory " + directory + " holds more than one results "
                + "file: " + string.Join(", ", candidates.Select(Path.GetFileName)) + ".");
        }

        return candidates[0];
    }

    /// <summary>
    /// Determines whether a file is the results file of a given job of a recognized simulation.
    /// </summary>
    /// <param name="path">The path of the file.</param>
    /// <param name="job">The job suffix the file must carry, ex. 3_of_8.</param>
    /// <returns>True when the file is that job's results file.</returns>
    private static bool IsResultsOf(string path, string job)
    {
        Match match = ResultsPattern.Match(Path.GetFileName(path));

        return match.Success
            && TestTypes.ContainsKey(match.Groups["type"].Value)
            && string.Equals(match.Groups["index"].Value + "_of_" + match.Groups["count"].Value, job,
                StringComparison.Ordinal);
    }

    /// <summary>
    /// Reads the name of the simulation that produced a job's results file.
    /// </summary>
    /// <param name="job">The job whose results file is being identified.</param>
    /// <returns>The name as it appears in the file name, ex. backtest.</returns>
    private static string TestTypeOf(SweepJob job)
    {
        // The file was matched against the same pattern when it was found, so it matches here too.
        return ResultsPattern.Match(Path.GetFileName(job.ResultsPath)).Groups["type"].Value;
    }

    /// <summary>
    /// Reads the configuration the sweep was driven from, as the first job recorded it.
    /// </summary>
    /// <param name="job">The first job of the sweep.</param>
    /// <returns>The lines of the file.</returns>
    /// <remarks>
    /// Every job of a sweep drives its own copy of one workbook, so every job records the same configuration and one
    /// of them is enough.  The first is used because it is the one a sweep always has.
    /// </remarks>
    private static string[] ReadRunConfiguration(SweepJob job)
    {
        string path = Path.Combine(job.Directory, Pfm.RunConfiguration.FileName);

        if (!File.Exists(path))
        {
            throw new InvalidOperationException("The run directory " + job.Directory + " holds no "
                + Pfm.RunConfiguration.FileName + ".  A job writes it before its first simulation, so a directory "
                + "without one is not a sweep this tool produced.");
        }

        return File.ReadAllLines(path);
    }

    /// <summary>
    /// Computes the identifier of a sweep.
    /// </summary>
    /// <param name="jobs">The jobs of the sweep.</param>
    /// <param name="runConfiguration">The lines of the configuration the sweep was driven from.</param>
    /// <returns>Eight hexadecimal characters.</returns>
    /// <remarks>
    /// The configuration is what the identifier is mostly for: it is the plan the sweep exercised, and it is what a
    /// reader comparing two coalesced workbooks is comparing.  The run directory names are folded in as well, because
    /// the configuration alone would give the same identifier to two sweeps of one unchanged plan, and the second
    /// would then be written over the first.  The names carry each job's timestamp, so they distinguish those sweeps
    /// while still being the same for every coalesce of one sweep: coalescing the same directories twice produces the
    /// same identifier, and therefore the same workbook name, rather than accumulating copies.
    /// </remarks>
    private static string ComputeRunId(IReadOnlyList<SweepJob> jobs, string[] runConfiguration)
    {
        var material = new StringBuilder();

        foreach (string line in runConfiguration)
        {
            material.Append(line).Append('\n');
        }

        foreach (SweepJob job in jobs.OrderBy(job => job.Name, StringComparer.Ordinal))
        {
            material.Append(job.Name).Append('\n');
        }

        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(material.ToString()));

        // Four bytes of a cryptographic hash, which is what eight characters can carry.  This identifies a sweep among
        // the handful a person keeps; it is not being relied on to be unforgeable.
        return Convert.ToHexStringLower(hash.AsSpan(0, 4));
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
