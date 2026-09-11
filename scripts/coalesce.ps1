<#
.SYNOPSIS
    Gathers the run directories of one simulation sweep into a single workbook.

.DESCRIPTION
    Arguments are handed to pfm untouched, ex. -i, -o and -t.

.EXAMPLE
    .\coalesce.ps1 -i .\runs -o .\output
#>
$RunPfm = Join-Path $PSScriptRoot "runpfm.ps1"

& $RunPfm coalesce @args
