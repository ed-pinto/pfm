<#
.SYNOPSIS
    Applies the analysis of the template workbook to one coalesced sweep.

.DESCRIPTION
    The sweep is the workbook coalesce wrote, PortfolioSimData.<run-id>.xlsx.  The template is the analysis workbook,
    which is the one named by $PfmAnalysisTemplate when the session has set that variable and the default below
    otherwise.  Anything beyond the parameters below is handed to pfm untouched, ex. -o and -t.

.EXAMPLE
    .\applyanalysis.ps1 .\output\PortfolioSimData.48a0f144.xlsx

.EXAMPLE
    .\applyanalysis.ps1 .\output\PortfolioSimData.48a0f144.xlsx -o .\output -t
#>
param(
    # The workbook holding the sweep to analyse.
    [string]$Sweep,

    # The analysis template.  Defaults to $PfmAnalysisTemplate, or to the template beside the model workbook.
    [string]$Template
)

if (-not $Sweep) {
    throw "The workbook of one coalesced sweep is required, ex. .\output\PortfolioSimData.48a0f144.xlsx."
}

$RunPfm = Join-Path $PSScriptRoot "runpfm.ps1"

$DefaultTemplate = "H:\OneDrive\Documents\Financials\Planning\Portfolio\PortfolioSimAnalysis.xlsx"
$AnalysisTemplate = if ($Template) { $Template } elseif ($PfmAnalysisTemplate) { $PfmAnalysisTemplate } else { $DefaultTemplate }

& $RunPfm apply-analysis -f $Sweep --template $AnalysisTemplate @args
