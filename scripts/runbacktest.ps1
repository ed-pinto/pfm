$PfmExe = Join-Path $PSScriptRoot "..\pfm\bin\Debug\net10.0\pfm.exe"
$PortfolioFile = "H:\OneDrive\Documents\Financials\Planning\Portfolio\PortfolioManager.xlsx"

0..3 | ForEach-Object {
    Start-Process $PfmExe -ArgumentList "back-test", "-f", $PortfolioFile, "--job-count", "4", "--job-index", $_
}