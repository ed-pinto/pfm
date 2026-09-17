namespace Pfm;

/// <summary>
/// Represents the arguments for the pfm CLI.
/// </summary>
public class Arguments
{
    /// <summary>
    /// The command to be executed by the pfm CLI.
    /// </summary>
    public required Commands Command { get; init; }

    /// <summary>
    /// The path of the workbook the command was pointed at by --file-path, or null for a command that is pointed at no
    /// workbook, ex. <see cref="Commands.Coalesce"/>, which reads what a finished sweep left on disk.
    /// </summary>
    /// <remarks>
    /// Which workbook that is belongs to the command rather than to this property, and the two kinds are not
    /// interchangeable: the commands that drive the model are pointed at PortfolioManager.xlsx, and
    /// <see cref="Commands.ApplyAnalysis"/> is pointed at the workbook of one coalesced sweep, ex.
    /// PortfolioSimData.48a0f144.xlsx.  It is one property because it is one option, declared by each command in its
    /// own words and read, validated and logged the same way for all of them; a second property would hold the same
    /// kind of value, leave one of the two empty on every run, and still not say which workbook was meant without
    /// naming the command.  The analysis template is a property of its own, <see cref="TemplatePath"/>, because it is
    /// an option of its own.
    /// </remarks>
    public string? FilePath { get; init; }

    /// <summary>
    /// The path to the analysis template workbook a sweep is analysed against, or null for a command that applies no
    /// analysis.
    /// </summary>
    /// <remarks>
    /// The template holds the whole of the analysis, expressed in its own formulas, tables and charts.  It is never
    /// written to: the command that applies it copies it and writes the sweep into the copy.
    /// </remarks>
    public string? TemplatePath { get; init; }

    /// <summary>
    /// The directory holding the run directories of the sweep being coalesced.  Defaults to the current directory.
    /// </summary>
    public string InputPath { get; init; } = "./";

    /// <summary>
    /// The directory the coalesced workbook, or the analysis of one, is written to.  Defaults to output under the
    /// current directory.
    /// </summary>
    public string OutputPath { get; init; } = "./output";

    /// <summary>
    /// Indicates whether the operation should report how long each of its phases took.  This is off unless the caller
    /// asked for it, ex. to find out where a slow run spent its time.
    /// </summary>
    public bool Timing { get; init; }

    /// <summary>
    /// The number of jobs a simulation sweep is split across.  One means the sweep is not being run in parallel.
    /// </summary>
    /// <remarks>
    /// A sweep takes hours, so it is divided into slices that run as separate processes.  Every job of one sweep is
    /// given the same job count and its own job index; nothing coordinates them beyond that, so they may be started on
    /// one machine or several.
    /// </remarks>
    public int JobCount { get; init; } = 1;

    /// <summary>
    /// The zero based index of the slice of a simulation sweep this process runs.  It is less than
    /// <see cref="JobCount"/>.
    /// </summary>
    public int JobIndex { get; init; }

    /// <summary>
    /// The greatest number of simulations the whole sweep runs, or null when it runs every one the workbook offers.
    /// </summary>
    /// <remarks>
    /// This shortens the sweep, ex. to the first hundred iterations of a Monte Carlo campaign, rather than limiting
    /// what a job does: a sweep of a hundred split across four jobs is still a hundred simulations.  It is what makes
    /// a trial run of a sweep that takes hours cost minutes instead.
    /// </remarks>
    public int? SimulationCount { get; init; }
}
