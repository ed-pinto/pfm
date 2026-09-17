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

Before the first simulation, the job records the configuration of the workbook it is about to drive. The copy of the
workbook is deleted when the job ends, so this is what a results file is read against afterwards. Writing it first also
means an interrupted sweep still says what it was running.

It is recorded twice, because the two forms are for different readers:

`RunConfiguration.txt` is the synopsis, one line for each row of the `Synopsis` table on `00_Overview`, which is where
the workbook states its configuration as prose.

Beside it, one CSV file per table of the plan itself, written as values rather than sentences, so that the tables can
be sorted, filtered and compared column by column:

| File | Copied from |
| --- | --- |
| `RunConfiguration.Expenditures.csv` | `Expenditures` on `52_Expenditures` |
| `RunConfiguration.TargetNetIncomeEras.csv` | `TargetNetIncomeEras` on `51_Targets` |
| `RunConfiguration.AllocationTargets.csv` | `AllocationTargets` on `13_Allocations` |
| `RunConfiguration.AllocationFloors.csv` | `AllocationFloors` on `13_Allocations` |
| `RunConfiguration.DirectedDraws.csv` | `DirectedDraws` on `13_Allocations` |
| `RunConfiguration.Parameters.csv` | The defined names below, as `Name` and `Value` |

Each table is copied whole, with the headings the workbook gives its columns and the values it last calculated. A table
the plan left empty is written as its heading row alone, which is what distinguishes a plan that directs no draws from
a sweep whose configuration was never recorded.

`Parameters` is the same idea for the scalar parameters, which `10_Parameters` states as defined names rather than as a
table: `CdnResidencyYear`, `YearlyMaxDisposal`, `EquityGainsRebalanceThreshold`, `ConsumptionStressStart`,
`ConsumptionStressFull`, `BootstrapBlockYears` and `BootstrapMatchPool`. A name the workbook does not define fails the
run rather than being recorded empty, so a sweep that takes hours cannot leave an incomplete record of what it ran.

The files are written the way the results are, and read back the same way: the invariant culture throughout, `TRUE` or
`FALSE` for a flag, and an empty field for an empty cell. A line break inside a cell, ex. a note typed with `Alt+Enter`,
becomes a space.

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
driving the one you point it at, records [the run configuration](#run-configuration) in the run directory before the
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

The workbook is written as `<output-path>/PortfolioSimData.<run-id>.xlsx`, where the run id is eight hexadecimal
characters derived from [the run configuration](#run-configuration) and the run directory names. The configuration is
most of what the identifier is for, since it is the plan the sweep exercised and what a reader comparing two of these
workbooks is comparing; both the synopsis and the tables are folded in, because the synopsis is a summary and two plans
that differ only in a row it does not spell out are still two plans. The directory names are folded in as well, because
the configuration alone would give one name to two sweeps of an unchanged plan and the second would be written over the
first. Coalescing the same sweep twice therefore produces the same name, and an existing workbook of that name is
reported rather than replaced.

| Sheet | Contents |
| --- | --- |
| `RunConfiguration` | The configuration the sweep was driven from, read from the first job: the synopsis at the top, one line per row of column A, and then each of the [configuration tables](#run-configuration) beneath it, stacked down the sheet with a blank row between them. Every job of a sweep drives a copy of one workbook, so they all record the same configuration. The synopsis cells are formatted as text so that it reads back exactly as the workbook wrote it; each table is labelled with its name and defined as an Excel table, so it can be filtered and referred to by name. |
| `SimData` | The header row once, then the data rows of every job in job index order, which is the order that reproduces what a single job would have written. The sheet carries one name whichever simulation produced it: which one it was is stated on `RunConfiguration`, and a name that varied would buy a branch in every reader of the sheet and nothing for a person. |

Fields are imported as the types they were written from: numbers as numbers, `TRUE` and `FALSE` as booleans, and an
empty field as an empty cell, so the sheet can be charted and filtered without converting anything first. The full
precision of each value survives the round trip.

Every job's header is checked against the first job's. Two that disagree are results of different plans, or of
different projection horizons, and concatenating them would produce a sheet whose columns meant different things in
different rows.
## apply-analysis

Applies the analysis of a template workbook to one coalesced sweep, and writes the result as a workbook of values.

```
pfm apply-analysis -f <sweep workbook> --template <analysis template>
pfm apply-analysis -f <sweep workbook> --template <analysis template> --output-path <directory>
```

`--file-path` is the workbook [coalesce](#coalesce) wrote, `PortfolioSimData.<run-id>.xlsx`, and `--template` is the
analysis template, `PortfolioSimAnalysis.xlsx`. Neither is written to. The output is
`<output-path>/PortfolioSimAnalysis.<run-id>.xlsx`, named for the same run as the data it analyses, so the two sort
together in a listing; `--output-path` defaults to `output` beneath the current directory.

### The analysis lives in the template

The template is a workbook, authored by hand, that holds the whole of the analysis: some 3400 formulas, 81 defined
names, 15 charts, the conditional formatting, and an `AnalysisSpec` worksheet stating why each of them is what it is.
This command does not describe any of that a second time. It copies the template, writes the sweep into the copy, lets
the template compute, checks what it computed, and flattens the result to values.

What the command knows about the template is three references wide, which is the whole of the coupling between a dataset
and the analysis of it:

| The template reads | Formulas | For |
| --- | --- | --- |
| `SimData`, the table over the results | 1579 | Every metric of every simulation, addressed by header text rather than by position |
| `TargetNetIncomeEras`, a table on `RunConfiguration` | 79 | The income target and floor of each projected year |
| `RunConfiguration!A1` | 1 | The title line of the sweep, echoed under the heading of the analysis |

No formula and no chart series names the worksheet the results are on, which is why one analysis reads a back test and a
Monte Carlo campaign without a change.

### What it does

1. Copies the template to the output name. The copy is the workbook that is driven; the template is only ever read.
2. Reads the sweep's configuration and results out of the sweep workbook, opened beside the copy and closed again
   before anything is recalculated.
3. Writes the configuration onto the copy's `RunConfiguration`, resizing each table to the rows this plan states and
   moving the blocks below it by inserting or deleting whole worksheet rows.
4. Writes the results onto the copy's `SimData`, resizing the table over them and clearing whatever the last sweep left
   beyond them.
5. Recalculates every formula, not only the ones Excel marked dirty: the spills the analysis is built on resize against
   the table that has just been redefined.
6. Checks the analysis (below). A failure names what broke and leaves no workbook behind.
7. Stamps the provenance line into `SimAnalysis!B2`: the template version, the run id, the workbook the sweep was read
   from, and the time of the run.
8. Flattens `SimAnalysis` and then `SimAnalysisCalc` to values, drops the defined names that only mean anything while
   those formulas are live, removes `AnalysisSpec`, and saves.

The tables are resized rather than replaced throughout, because a table that is dissolved takes its references with it:
Excel rewrites every formula that read it structurally into the cell range it happened to occupy, so the next sweep
would be read through the last one's geometry with nothing reported. For the same reason the results are written as
values into the cells of the table rather than pasted over it, and the heading row is written on its own: a single write
covering a table's whole block, heading row included, replaces the table the way a paste does.

### The output workbook

Four worksheets, all values: `RunConfiguration`, `SimData`, `SimAnalysisCalc` and `SimAnalysis`. Nothing in it can
recalculate, which is the point of it: the inputs of the analysis were decided in the template, and a question about how
a figure was derived is answered by opening the template rather than by editing the output.

What survives the flatten is what a reader needs: the number formats, the conditional formatting, the tables `SimData`
can still be filtered through, the `#N/A` of a trial slot with no trial to plot, and the charts, which need no attention
because every series reads a range that still holds the same values. What does not survive is the names anchored to a
spill, which refer to nothing once the spill is a block of values.

`SimAnalysis` and `SimAnalysisCalc` are flattened in that order. The names the analysis is built on are anchored to the
spills on the calculation sheet, so flattening the calculation sheet first would leave every figure that reads one of
those names saved as `#REF!`.

An analysis re-applied to a sweep already analysed resolves to the name already taken, and the workbook that is there is
reported rather than replaced, as `coalesce` does with its own output. Which version of the analysis produced a workbook
is stated in the workbook, not in its name.

### The checks

Every check runs before any of them is reported, so a template or a harvest that needs fixing is described once rather
than one line at a time. They are the validation list of the `AnalysisSpec` worksheet:

| Check | What it catches |
| --- | --- |
| Every metric key the analysis states resolves to a block of `<MetricKey>_<YYYY>` columns the results carry | A harvest that renamed or dropped a metric. This is the one check that fails on a dataset the template cannot analyse at all, and it names the metric and the column it could not find |
| `SimulationCount` equals the rows written; `ProjectionYearCount` and `FirstProjectionYear` agree with the headers | A table that does not cover the sweep, ex. one left at the previous sweep's geometry |
| The sweep projects at most 39 years | A horizon the analysis has no room for. The blocks occupy `B:AN`, so a longer one is a change to the template rather than a dataset it can be applied to |
| A simulation is identified by at most 5 columns | A sweep whose per-simulation diagnostics would reach into the block beside them. A back test identifies a simulation by three columns and a Monte Carlo campaign by two, and each diagnostic on the calculation sheet is one column wider than that. Excel reports the collision as `#SPILL!` on one cell and says nothing about the sweep that provoked it |
| No defined name refers to anything bracketed | A name carried in from another workbook, which resolves against whatever copy of that file is open and reads the wrong sweep without saying so |
| No cell holds an error other than the `#N/A` of a blank trial slot, and no cell's value is a string beginning with `=` | A formula that did not take |
| Every year header row holds 39 years rather than one | A year header row formatted as text before its formula was written, which spills nothing and fails silently while the chart above it still renders |
| No chart's bottom edge passes the top of the next non-blank row | A chart covering the block below it, which a reader of a workbook with no formulas cannot fix |
| The count of simulations that fell below the income floor is bracketed by what the `IncomeShortfall` columns report | An analysis reading the wrong columns. This is the one check that reads the results themselves rather than what the analysis says about them |
