# pfm Readme
pfm is a command line interface for performing operations against the PortfolioManager Excel workbook.

## iterate

Iterates the tax provision on `15_TaxPayments` until it agrees with the modelled total tax on `55_Tax`, which is what
check K39 on `90_Checks` tests. It works on the workbook you point it at, and saves it unless that workbook was already
open in Excel, in which case it leaves the saving to you.

```
pfm iterate -f <workbook>
```

## back-test

Projects the workbook from every historical period in `21_EconomyHistorical` long enough to project over, and records
the outcome of each. One simulation sets `BackTestingStartYear` and `BackTestingStartSemester`, which moves
`30_MarketSimulation` onto a different run of historical semesters and reprojects the whole model from it; the tax
provision is then iterated to convergence, and the outcome metrics are harvested from the `Projected` years of
`SummaryAnnual` on `56_Summary`. The starting point advances one semester per simulation.

```
pfm back-test -f <workbook>
pfm back-test -f <workbook> --job-count 4 --job-index 0
```

The command never touches the workbook you point it at. Each run creates a directory `pfm.<yyyyMMddHHmm>.<index>_of_<count>`
under the current directory, copies the workbook into it, drives the copy in an Excel process of its own, and writes
`backtest.<yyyyMMddHHmm>.<index>_of_<count>.csv` beside it.

### Run configuration

Before the first simulation, the job writes `RunConfiguration.txt` into its run directory: one line for each row of
the `Synopsis` table on `00_Overview`, which is where the workbook states the configuration a run exercises. The copy
of the workbook is deleted when the job ends, so this is what a results file is read against afterwards. Writing it
first also means an interrupted sweep still says what it was running.

### Running a sweep in parallel

A sweep is split into contiguous slices, one per job. Every job of one sweep is given the same `--job-count` and its own
`--job-index`, which is **zero based** and less than the job count; nothing coordinates the jobs beyond that, so they can
be started from one shell or several. Concatenating the data rows of the results files, in job index order, reproduces
what a single job would have written; `SimulationIndex` identifies each simulation within the whole sweep.

```powershell
0..3 | ForEach-Object {
    Start-Process pfm -ArgumentList "back-test", "-f", $workbook, "--job-count", "4", "--job-index", $_
}
```

Each job runs its own Excel process, verified by process id at startup rather than assumed, because two jobs sharing an
Excel process would recalculate each other's workbook.

### Results

One row per simulation. The columns are:

| Columns | Meaning |
| --- | --- |
| `SimulationIndex`, `StartYear`, `StartSemester` | Which simulation the row is, and the historical period it was driven from. |
| `<DataElement>_<Year>` | One column per simulated year, repeated for each harvested element, so an element reads across as a contiguous block of years. The elements are `InflationIndex`, `NetIncome`, `EndingPortfolioValue`, `TotalTax`, `PortfolioConsumedRatio`, `RatioBondCash`, `RatioBondLadder`, `RatioEquityConcentrated`, `RatioEquityCore` and `RatioEquityInternational`. |
| `ConvergencePasses`, `FinalTaxDrift`, `DriftWithinTolerance` | How the tax provision behind that simulation settled. `DriftWithinTolerance` is TRUE when the drift is within `TaxProvisionTolerance`, which is the condition K39 tests. |

Values are written to import into Excel as numbers and booleans rather than text: invariant culture, no grouping
separators or currency and percent formatting, and the shortest representation that reads back as the same value. A
simulation whose tax never settled is still recorded, with `DriftWithinTolerance` FALSE, rather than dropped, because
dropping it would bias the sweep towards the periods that were easy to settle.

Rows are flushed as each simulation finishes, so a sweep that is interrupted leaves behind everything it had finished.
