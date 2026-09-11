<#
.SYNOPSIS
    Runs a back test sweep, split across the given number of parallel jobs.

.DESCRIPTION
    Each job is a pfm process of its own, started through runpfm.ps1.  Anything beyond the job count is handed to
    every job untouched, ex. -t.

.EXAMPLE
    .\backtest.ps1 8 -t
#>
param(
    # The number of processes the sweep is split across.
    [ValidateRange(1, 64)]
    [int]$JobCount = 6
)

$RunPfm = Join-Path $PSScriptRoot "runpfm.ps1"

# The ForEach-Object script block has an $args of its own, so what this script was handed has to be captured here
# to be splatted inside the loop.
$PassThruArgs = $args

0..($JobCount - 1) | ForEach-Object {
    & $RunPfm back-test -Background --job-count $JobCount --job-index $_ @PassThruArgs
}
