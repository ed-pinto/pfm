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
    /// Iterates calculations in the PortfolioManager workbook which require iterative updates.
    /// Ex. 15_TaxPayments.TaxProvision.
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
    /// There are many configurations in the workbook affected by the year in which tax residency changes to Canada.
    /// Ex. When IRA positions stop being purchased.
    /// This command updates all related configurations including 10_Parameters.
    /// </summary>
    [EnumMember(Value = "SetCanadianMoveYear")]
    SetCanadianMoveYear
}
