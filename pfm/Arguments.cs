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
}
