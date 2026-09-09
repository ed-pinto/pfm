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
    /// The file path to the PortfolioManager.xlsx file on which to operate.
    /// </summary>
    public required string FilePath { get; init; }

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
}
