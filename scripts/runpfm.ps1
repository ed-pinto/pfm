<#
.SYNOPSIS
    Runs a pfm command against the PortfolioManager workbook.

.DESCRIPTION
    This is the one script that knows where pfm.exe and the workbook live.  Every other script in this directory
    calls it rather than invoking pfm.exe itself.

    Anything beyond the parameters below is handed to pfm untouched, ex. -t.  The parameters are deliberately
    declared without [Parameter()] attributes: that keeps the script out of PowerShell's advanced binding, where
    common parameters would capture pass through options such as -i and -o before pfm ever saw them.

    The workbook is the one named by $PfmPortfolioFile when the session has set that variable, and the default
    below otherwise.

.EXAMPLE
    .\runpfm.ps1 iterate -t

.EXAMPLE
    $PfmPortfolioFile = "C:\Users\pinto\OneDrive\Documents\Financials\Planning\Portfolio\PortfolioManager.xlsx"
    .\runpfm.ps1 iterate
#>
param(
    [string]$Command,

    # Starts pfm in a process of its own and returns immediately, rather than running it to completion here.  This
    # is what lets a caller fan a sweep out across several processes.
    [switch]$Background
)

if (-not $Command) {
    throw "A pfm command is required, ex. iterate, back-test, monte-carlo or coalesce."
}

$PfmExe = Join-Path $PSScriptRoot "..\pfm\bin\Debug\net10.0\pfm.exe"

# The workbook path comes from $PfmPortfolioFile when the session has set one, so a machine whose OneDrive sits
# somewhere else needs no edit here.  PowerShell looks a variable up through the enclosing scopes, so this reads
# whatever the prompt or a calling script has defined and falls back to the default when nothing has.
$DefaultPortfolioFile = "H:\OneDrive\Documents\Financials\Planning\Portfolio\PortfolioManager.xlsx"
$PortfolioFile = if ($PfmPortfolioFile) { $PfmPortfolioFile } else { $DefaultPortfolioFile }

# coalesce reads what a finished sweep left on disk and has no workbook to be pointed at, so it is the one command
# that is not given -f.  Naming the exceptions rather than the workbook commands means a command added later gets
# the workbook without a change here.
$CommandsWithoutWorkbook = @("coalesce")

$PfmArgs = @($Command)

if ($CommandsWithoutWorkbook -notcontains $Command) {
    $PfmArgs += @("-f", $PortfolioFile)
}

if ($args) {
    $PfmArgs += $args
}

if ($Background) {
    # Start-Process joins its argument list with spaces without quoting any of it, so a path holding a space has to
    # be quoted here.
    $QuotedArgs = $PfmArgs | ForEach-Object { if ($_ -match "\s") { '"' + $_ + '"' } else { $_ } }

    Start-Process -FilePath $PfmExe -ArgumentList $QuotedArgs
}
else {
    & $PfmExe @PfmArgs
}
