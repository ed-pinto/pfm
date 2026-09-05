param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Command,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$RemainingArgs
)

$PfmExe = Join-Path $PSScriptRoot "..\pfm\bin\Debug\net10.0\pfm.exe"
$PortfolioFile = "H:\OneDrive\Documents\Financials\Planning\Portfolio\PortfolioManager.xlsx"

& $PfmExe $Command -f $PortfolioFile @RemainingArgs
