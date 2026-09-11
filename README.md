# pfm Readme
pfm is a command line interface for performing operations against the PortfolioManager Excel workbook.

## iterate

Iterates the workbook's materialized series until each agrees with the live value it stands for. It works on the
workbook you point it at, and saves it unless that workbook was already open in Excel, in which case it leaves the
saving to you.

```
pfm iterate -f <workbook>
```

### The materialized series

The workbook materializes a value wherever a live link would be circular at range level, which Excel does not report as
an error: it silently freezes the column instead (D009). Each such column is a static series refreshed from a live
column beside it, over the rows a status column marks, within a tolerance the workbook states, and guarded by a check
on `90_Checks` that compares the two.

| Series | Refreshed | From | Over rows where | Tolerance | Check |
| --- | --- | --- | --- | --- | --- |
| `15_TaxPayments[TaxProvision]` | `ProvisionAmount` | `TotalTaxLive` | `ProvisionSource` is `ProvisionModelled` | `TaxProvisionTolerance` | K39 |
| `16_PortfolioState[PortfolioState]` | `OpeningPortfolioValue` | `OpeningPortfolioLive` | `Regime` is `StatusProjected` | `PortfolioStateTolerance` | K53 |

They are settled together in one loop rather than one after another, because they are coupled: `ProvisionAmount` feeds
the tax funding `55_Tax` draws on, and `OpeningPortfolioValue` feeds the income throttle `51_Targets` draws on, which
`50_Assets` and so the opening portfolio itself respond to. Settling one and then the other would leave the first stale
again. A pass therefore writes every stale series and pays for one recalculation rather than one per series, which is
also what step 16 of the workbook's own close runbook asks for.

Adding a series to the loop is a matter of adding one entry to `ConvergenceIterator.Definitions`: the loop, the
outcome and the results file are all stated in terms of that list.

## back-test

Projects the workbook from every historical period in `21_EconomyHistorical` long enough to project over, and records
the outcome of each. One simulation sets `BackTestingStartYear` and `BackTestingStartSemester`, which moves
`30_MarketSimulation` onto a different run of historical semesters and reprojects the whole model from it; the
materialized series are then iterated to convergence as `iterate` does, and the outcome metrics are harvested from the
`Projected` years of `SummaryAnnual` on `56_Summary`. The starting point advances one semester per simulation.

```
pfm back-test -f <workbook>
pfm back-test -f <workbook> --job-count 4 --job-index 0
pfm back-test -f <workbook> --simulation-count 10
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

### Running part of a sweep

`--simulation-count` caps the number of simulations the **whole sweep** runs. It shortens the sweep rather than
sampling it, so the simulations that run are the first ones: the earliest historical periods of a back test, the lowest
iterations of a Monte Carlo campaign. This is what makes a trial run of a sweep that takes hours cost minutes.

```
pfm monte-carlo -f <workbook> --simulation-count 100
pfm monte-carlo -f <workbook> --simulation-count 100 --job-count 4 --job-index 0
```

The limit is applied before the sweep is divided, so the second line above is four jobs of 25 simulations rather than
of 100. `SimulationIndex` still identifies a simulation within the whole sweep, and a simulation is the same
simulation whether or not the sweep it belongs to was shortened: iteration 7 of a campaign limited to 100 is iteration
7 of the campaign limited to nothing. A limit larger than what the workbook offers is not an error — it asks for at
most that many simulations, and a workbook with fewer has already answered.

### Results

One row per simulation. The columns are:

| Columns | Meaning |
| --- | --- |
| `SimulationIndex`, `StartYear`, `StartSemester` | Which simulation the row is, and the historical period it was driven from. |
| `<DataElement>_<Year>` | One column per simulated year, repeated for each harvested element, so an element reads across as a contiguous block of years. The elements are `NetIncome`, `NetIncomeCurrentDollars`, `EndingPortfolioValue`, `EndingPortfolioValueCurrentDollars`, `TotalTax`, `TotalTaxCurrentDollars`, `PortfolioConsumedRatio`, `RatioBondCash`, `RatioBondLadder`, `RatioEquityConcentrated`, `RatioEquityCore` and `RatioEquityInternational`. |
| `ConvergencePasses`, `Final<Series>Drift`, `DriftWithinTolerance` | How the materialized series behind that simulation settled: `FinalTaxProvisionDrift` and `FinalPortfolioStateDrift` today, one per series in the order the [table above](#the-materialized-series) states them. The pass count is shared because one loop settles them together, and `DriftWithinTolerance` is TRUE only when every series is within its own tolerance, which is the condition K39 and K53 test between them. |

Values are written to import into Excel as numbers and booleans rather than text: invariant culture, no grouping
separators or currency and percent formatting, and the shortest representation that reads back as the same value. A
simulation whose series never settled is still recorded, with `DriftWithinTolerance` FALSE, rather than dropped, because
dropping it would bias the sweep towards the periods that were easy to settle.

Rows are flushed as each simulation finishes, so a sweep that is interrupted leaves behind everything it had finished.

## monte-carlo

Projects the workbook down every Monte Carlo path `22_MCSeeds` holds seeds for, and records the outcome of each. One
simulation sets `SimulationMode` to `MonteCarlo` and `MCIteration` to the iteration it runs, which is the whole of the
contract the workbook states for an external runner: the iteration selects the column of seeds the block bootstrap on
`30_MarketSimulation` draws its block start rows from, and the model reprojects down that path. The materialized
series are then iterated to convergence as `iterate` does, and the outcome metrics are harvested from the `Projected`
years of `SummaryAnnual` on `56_Summary`, exactly as a back test harvests them.

```
pfm monte-carlo -f <workbook>
pfm monte-carlo -f <workbook> --job-count 4 --job-index 0
pfm monte-carlo -f <workbook> --simulation-count 100
```

A Monte Carlo path is four spliced runs of real history rather than a sequence of independent draws, and it is
reproducible: the seed sheet is static, nothing in the driver set is random, and the iteration is all that selects a
path. Running iteration 7 twice therefore produces the same projection twice.

The mode is set once, before the first simulation, because it is what makes the mode selector on
`30_MarketSimulation` read the `MC_` columns at all; `MCIteration` is the only thing an individual simulation stamps.

### Which iterations are run

The sweep is one simulation per seed column of `MCSeeds`, numbered from one as those columns are, which is 1 to 1000
as the table stands today. The columns are counted rather than the range being stated in the tool, because
`GetMCSeed` reads the seed of a year from column `MCIteration + 1` of that table: the seed columns *are* the
iterations the workbook can be driven through, so widening the table is a change to the workbook alone. `SimulationIndex`
is the zero based position within the sweep, as it is for a back test, and `MCIteration` is that position plus one.

`--simulation-count` shortens the campaign to its first *n* iterations, so `--simulation-count 100` runs iterations 1
to 100 and leaves the other 900 paths unrun. See [Running part of a sweep](#running-part-of-a-sweep).

Everything else is what the back test does, and is described above: the command copies the workbook rather than
driving the one you point it at, writes [`RunConfiguration.txt`](#run-configuration) into the run directory before the
first simulation, is [split across jobs](#running-a-sweep-in-parallel) the same way, and writes the same
[results](#results) but for the identifying columns, which are `SimulationIndex` and `MCIteration` rather than
`SimulationIndex`, `StartYear` and `StartSemester`. The results file is `montecarlo.<yyyyMMddHHmm>.<index>_of_<count>.csv`.

Harvesting the same elements over the same years is deliberate: the two sweeps ask the same questions of different
market paths, and their results are only comparable if they are measured the same way.

## coalesce

Gathers the run directories one sweep left behind into a single workbook: the configuration the sweep exercised on one
sheet, and the results of every job, concatenated in job index order, on another.

```
pfm coalesce
pfm coalesce --input-path <directory> --output-path <directory>
```

`--input-path` defaults to the current directory, which is where a sweep leaves its run directories, and
`--output-path` to `output` beneath it. The output directory is created if it is not there. This command reads a sweep
and never writes to it, and it is the one command that does not take `--file-path`: everything it needs is what the
sweep left on disk.

### The input must be one whole sweep

The input directory must hold the run directory of every job of one sweep, `pfm.<stamp>.<index>_of_<count>` for each
index from zero to one less than the count. A sweep is only meaningful whole: the jobs cover contiguous slices of one
range of simulations, so a missing job is a gap in the middle of the results rather than a shorter run, and there is
nothing in a results file that would say so afterwards. Anything short of the full set is reported rather than
coalesced, as are two sweeps found in one directory.

The jobs stamp their directories as they start, so two that straddle a minute boundary carry different stamps; it is
the job count they agree on that identifies the sweep.

### The workbook

The workbook is written as `<output-path>/PortfolioSimAnalysis.<run-id>.xlsx`, where the run id is eight hexadecimal
characters derived from `RunConfiguration.txt` and the run directory names. The configuration is most of what the
identifier is for, since it is the plan the sweep exercised and what a reader comparing two of these workbooks is
comparing; the directory names are folded in as well, because the configuration alone would give one name to two sweeps
of an unchanged plan and the second would be written over the first. Coalescing the same sweep twice therefore produces
the same name, and an existing workbook of that name is reported rather than replaced.

| Sheet | Contents |
| --- | --- |
| `RunConfiguration` | The lines of `RunConfiguration.txt`, one per row of column A, read from the first job. Every job of a sweep drives a copy of one workbook, so they all record the same configuration. The column is formatted as text so the synopsis reads back exactly as the workbook wrote it. |
| `BackTestData` or `MonteCarloData` | The header row once, then the data rows of every job in job index order, which is the order that reproduces what a single job would have written. The sheet is named after the simulation that produced the results. |

Fields are imported as the types they were written from: numbers as numbers, `TRUE` and `FALSE` as booleans, and an
empty field as an empty cell, so the sheet can be charted and filtered without converting anything first. The full
precision of each value survives the round trip.

Every job's header is checked against the first job's. Two that disagree are results of different plans, or of
different projection horizons, and concatenating them would produce a sheet whose columns meant different things in
different rows.
