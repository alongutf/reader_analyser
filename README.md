# reader_analyser

Interactive plots of Tecan plate-reader kinetic data (i-control Excel exports).

`tecan_plot.py` reads a workbook in which **each worksheet is a separate run** of the
reader, places all runs on one common time axis using each sheet's `Start Time:` stamp,
and writes a self-contained HTML page with one time-course trace per well.

## Setup

```bash
conda env create -f environment.yml
conda activate reader_data_analysis
```

To update an existing environment: `conda env update -f environment.yml`.

## Usage

```bash
python tecan_plot.py path/to/experiment.xlsx --open
```

This writes, next to the input file:

| File | Contents |
| --- | --- |
| `<name>_plot.html` | Interactive plot (works offline; open in any browser) |
| `<name>_long.csv` | Tidy data: one row per well per time point |

### Options

| Option | Description |
| --- | --- |
| `-o, --out FILE` | Output HTML path |
| `--csv FILE` | Output CSV path |
| `--no-csv` | Don't write the CSV |
| `--blank WELL` | Background well: preselected in the plot, and adds a `value_bgsub` column to the CSV |
| `--cdn` | Load plotly.js from a CDN (≈4 MB smaller HTML, needs internet to view) |
| `--open` | Open the HTML in the browser when done |

## The plot

- **Time axis**: hours since the start of the earliest run. Each run's start is marked with
  a dotted vertical line; gaps between runs are shown as gaps, not joined lines.
- **Well picker**: a plate-layout grid above the plot. Click a well to show/hide it, a row
  letter or column number to toggle that whole row/column, or *Show all* / *Hide all*.
  The Plotly legend also works (click to toggle, double-click to isolate) and stays in sync.
- **Background well**: choose a well from the drop-down to subtract it from every well.
  Values are matched by run and cycle; points from a run in which the background well was
  not measured are left blank. Choose *none* to return to raw values.
- **Linear / Log** y-axis buttons, hover details (clock time, sheet, cycle, temperature),
  and a camera button to export the figure as SVG.

## Input format

The script expects the standard Tecan i-control export:

- One worksheet per run, each containing a `Start Time:` row (falls back to the
  `Date:` + `Time:` header rows).
- A kinetic table starting with a `Cycle Nr.` row, in either orientation
  (wells as columns, `Cycle Nr. | Time [s] | Temp. [°C] | A1 | A2 …`, or wells as rows).
- Multiple measurement labels (e.g. `Label: OD` and `Label: GFP`) each get their own plot.

Handled automatically: empty sheets, unfilled cycle rows from runs that were stopped
early, non-numeric cells (`OVER`, `Wait (Plate)`, …) treated as missing, and different
sets of wells in different runs. Sheet order does not matter — runs are ordered by
their timestamps.
