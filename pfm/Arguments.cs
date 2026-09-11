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
    /// The file path to the PortfolioManager.xlsx file on which to operate, or null for a command that does not
    /// operate on a workbook, ex. <see cref="Commands.Coalesce"/>.
    /// </summary>
    public string? FilePath { get; init; }

    /// <summary>
    /// The directory holding the run directories of the sweep being coalesced.  Defaults to the current directory.
    /// </summary>
    public string InputPath { get; init; } = "./";

    /// <summary>
    /// The directory the coalesced workbook is written to.  Defaults to output under the current directory.
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
