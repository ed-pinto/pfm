using System.Runtime.Serialization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace Pfm;

/// <summary>
/// Represents the available commands for the pfm CLI.
/// </summary>
[JsonConverter(typeof(StringEnumConverter))]
public enum Commands
{
    /// <summary>
    /// Iterates the materialized series of the PortfolioManager workbook until each agrees with the live value it
    /// stands for.  Ex. 15_TaxPayments.TaxProvision and 16_PortfolioState.PortfolioState.
    /// </summary>
    [EnumMember(Value = "Iterate")]
    Iterate,

    /// <summary>
    /// Closes the current month in the PortfolioManager workbook.  When the month is the last of the semester
    /// it also closes the semester.  See D048 and 91_CloseRunbook.  This is the implementation of 91_CloseRunbook.
    /// </summary>
    [EnumMember(Value = "ClosePeriod")]
    ClosePeriod,

    /// <summary>
    /// Performs a Monte Carlo simulation on the PortfolioManager workbook.
    /// </summary>
    [EnumMember(Value = "MonteCarlo")]
    MonteCarlo,

    /// <summary>
    /// Performs a backtest on the PortfolioManager workbook.
    /// </summary>
    [EnumMember(Value = "BackTest")]
    BackTest,

    /// <summary>
    /// Gathers the run directories of one simulation sweep into a single Excel workbook: the configuration the sweep
    /// exercised on one sheet, and the concatenated results of every job on another.
    /// </summary>
    [EnumMember(Value = "Coalesce")]
    Coalesce,

    /// <summary>
    /// There are many configurations in the workbook affected by the year in which tax residency changes to Canada.
    /// Ex. When IRA positions stop being purchased.
    /// This command updates all related configurations including 10_Parameters.
    /// </summary>
    [EnumMember(Value = "SetCanadianMoveYear")]
    SetCanadianMoveYear
}
