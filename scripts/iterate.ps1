<#
.SYNOPSIS
    Iterates the calculations in the PortfolioManager workbook that need iterative updates.

.DESCRIPTION
    Arguments are handed to pfm untouched, ex. -t.

.EXAMPLE
    .\iterate.ps1 -t
#>
$RunPfm = Join-Path $PSScriptRoot "runpfm.ps1"

& $RunPfm iterate @args
