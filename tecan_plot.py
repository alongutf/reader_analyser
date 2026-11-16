"""Plot Tecan i-control kinetic exports as an interactive, per-well time course.

Each worksheet in the workbook is one run (one initiation of the reader). Runs
are placed on a common time axis using the "Start Time:" stamp on each sheet,
so gaps between runs show up as real gaps.

Usage:
    python tecan_plot.py data.xlsx                # writes data_plot.html + data_long.csv
    python tecan_plot.py data.xlsx -o out.html --cdn --open

The HTML has a plate-layout well picker (click a well, a row letter, a column
number, or All/None) in addition to Plotly's own legend toggling, and a
"Background well" menu that subtracts the chosen well from every well
(matched by run and cycle). --blank WELL preselects it and adds a
background-subtracted column to the CSV.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import webbrowser
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
import plotly.graph_objects as go

WELL_RE = re.compile(r"^([A-P])(\d{1,2})$")


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #
def _to_float(v):
    if isinstance(v, bool):
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return np.nan  # "OVER", "Wait (Plate)", blanks, ...


def _parse_timestamp(value) -> pd.Timestamp | None:
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return pd.Timestamp(value)
    try:
        return pd.Timestamp(pd.to_datetime(str(value).strip()))
    except (ValueError, TypeError):
        return None


def _sheet_start(rows: list[tuple]) -> pd.Timestamp | None:
    """Run start: 'Start Time:' cell, else 'Date:' + 'Time:' header cells."""
    date = time = None
    for r in rows:
        key = str(r[0]).strip() if r and r[0] is not None else ""
        val = next((c for c in r[1:] if c is not None), None)
        if key == "Start Time:":
            ts = _parse_timestamp(val)
            if ts is not None:
                return ts
        elif key == "Date:":
            date = val
        elif key == "Time:":
            time = val
    if date is not None and time is not None:
        d = date.date() if isinstance(date, dt.datetime) else pd.to_datetime(str(date)).date()
        t = time if isinstance(time, dt.time) else pd.to_datetime(str(time)).time()
        return pd.Timestamp(dt.datetime.combine(d, t))
    return None


def _parse_block(rows: list[tuple], hdr: int) -> pd.DataFrame | None:
    """Parse one kinetic table whose 'Cycle Nr.' cell is at rows[hdr][0].

    Supports both export orientations: wells as columns (row header
    'Cycle Nr. | Time [s] | Temp | A1 | A2 ...') and wells as rows.
    """
    header = [str(c).strip() if c is not None else "" for c in rows[hdr]]
    if len(header) > 1 and header[1].startswith("Time"):
        # Wells as columns.
        cols = {i: h for i, h in enumerate(header) if WELL_RE.match(h)}
        records = []
        for r in rows[hdr + 1:]:
            if not r or r[0] is None or not isinstance(r[0], (int, float)):
                break
            t = _to_float(r[1]) if len(r) > 1 else np.nan
            if np.isnan(t):
                continue  # cycle placeholder rows of an aborted run
            temp = _to_float(r[2]) if len(r) > 2 else np.nan
            for i, well in cols.items():
                records.append((int(r[0]), t, temp, well, _to_float(r[i]) if i < len(r) else np.nan))
    else:
        # Wells as rows: 'Cycle Nr.' row, 'Time [s]' row, 'Temp' row, then one row per well.
        block = {}
        for r in rows[hdr:]:
            key = str(r[0]).strip() if r and r[0] is not None else ""
            if not key:
                break
            block[key] = r[1:]
        time_key = next((k for k in block if k.startswith("Time")), None)
        temp_key = next((k for k in block if k.startswith("Temp")), None)
        if time_key is None:
            return None
        records = []
        for j, cyc in enumerate(block["Cycle Nr."]):
            t = _to_float(block[time_key][j]) if j < len(block[time_key]) else np.nan
            if cyc is None or np.isnan(t):
                continue
            temp = _to_float(block[temp_key][j]) if temp_key and j < len(block[temp_key]) else np.nan
            for well, vals in block.items():
                if WELL_RE.match(well):
                    records.append((int(cyc), t, temp, well, _to_float(vals[j]) if j < len(vals) else np.nan))
    if not records:
        return None
    return pd.DataFrame(records, columns=["cycle", "time_s", "temp_c", "well", "value"])


def parse_sheet(ws) -> list[pd.DataFrame]:
    rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
    start = _sheet_start(rows)
    out = []
    label = "Value"
    for i, r in enumerate(rows):
        a = str(r[0]).strip() if r and r[0] is not None else ""
        if a.startswith("Label:"):
            label = a.split(":", 1)[1].strip() or label
        elif a == "Cycle Nr.":
            df = _parse_block(rows, i)
            if df is not None:
                df.insert(0, "label", label)
                out.append(df)
    if out and start is None:
        print(f"warning: sheet '{ws.title}' has data but no start time stamp; skipped", file=sys.stderr)
        return []
    for df in out:
        df.insert(0, "sheet", ws.title)
        df.insert(1, "run_start", start)
    return out


def load_workbook(path: Path) -> pd.DataFrame:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    frames = [df for ws in wb.worksheets for df in parse_sheet(ws)]
    wb.close()
    if not frames:
        raise SystemExit(f"No kinetic data found in {path}")
    data = pd.concat(frames, ignore_index=True)
    data["clock_time"] = data["run_start"] + pd.to_timedelta(data["time_s"], unit="s")
    t0 = data["run_start"].min()
    data["time_h"] = (data["clock_time"] - t0).dt.total_seconds() / 3600.0
    data = data.sort_values(["label", "well", "clock_time"]).reset_index(drop=True)
    return data


# --------------------------------------------------------------------------- #
# Plotting
# --------------------------------------------------------------------------- #
def _well_key(w: str):
    m = WELL_RE.match(w)
    return (m.group(1), int(m.group(2))) if m else (w, 0)


def build_figure(data: pd.DataFrame, label: str, runs: pd.DataFrame) -> go.Figure:
    d = data[data["label"] == label]
    wells = sorted(d["well"].unique(), key=_well_key)
    fig = go.Figure()
    for well in wells:
        w = d[d["well"] == well]
        xs, ys, cd = [], [], []
        # Break the line between runs so separate runs are not joined across the gap.
        for _, g in w.groupby("run_start", sort=True):
            xs += g["time_h"].tolist() + [None]
            ys += g["value"].tolist() + [None]
            cd += list(zip(g["sheet"], g["clock_time"].dt.strftime("%Y-%m-%d %H:%M:%S"),
                           g["cycle"], g["temp_c"])) + [(None, None, None, None)]
        fig.add_trace(go.Scatter(
            x=xs, y=ys, customdata=cd, name=well, mode="lines+markers",
            marker=dict(size=4), line=dict(width=1.5), connectgaps=False,
            hovertemplate=(f"<b>{well}</b><br>t = %{{x:.2f}} h<br>{label} = %{{y:.4f}}"
                           "<br>%{customdata[1]}<br>%{customdata[0]}, cycle %{customdata[2]}"
                           "<br>%{customdata[3]} °C<extra></extra>"),
        ))

    t0 = data["run_start"].min()
    for _, run in runs.iterrows():
        x = (run["run_start"] - t0).total_seconds() / 3600.0
        fig.add_vline(x=x, line=dict(color="grey", dash="dot", width=1))
        fig.add_annotation(x=x, y=1, yref="paper", yanchor="top", xanchor="right", textangle=-90,
                           text=f"{run['sheet']} · {run['run_start']:%H:%M}", showarrow=False,
                           font=dict(size=10, color="grey"))

    fig.update_layout(
        title=f"{label} vs time (t = 0 at {t0:%Y-%m-%d %H:%M:%S})",
        xaxis_title="Time [h]", yaxis_title=label, template="plotly_white",
        hovermode="closest", legend=dict(title="Well (click to toggle,<br>double-click to isolate)"),
        margin=dict(t=80), height=650,
        updatemenus=[dict(type="buttons", direction="left", x=1, y=1.12, xanchor="right", yanchor="bottom",
                          buttons=[dict(label="Linear", method="relayout", args=[{"yaxis.type": "linear"}]),
                                   dict(label="Log", method="relayout", args=[{"yaxis.type": "log"}])])],
    )
    return fig


PAGE = """<!doctype html>
<html><head><meta charset="utf-8"><title>{title}</title>
<style>
 body {{ font-family: system-ui, sans-serif; margin: 16px; color: #222; }}
 h1 {{ font-size: 18px; }}
 .picker {{ display: inline-block; margin: 6px 0 4px; }}
 .picker table {{ border-collapse: collapse; }}
 .picker td {{ padding: 2px; text-align: center; font-size: 12px; }}
 .picker button {{ font: inherit; cursor: pointer; border: 1px solid #bbb; border-radius: 4px; background: #fff; }}
 .picker .well {{ width: 40px; height: 26px; }}
 .picker .well.on {{ color: #fff; border-color: transparent; }}
 .picker .well:disabled {{ visibility: hidden; }}
 .picker .hdr {{ background: #f2f2f2; min-width: 26px; }}
 .controls button {{ padding: 3px 10px; margin-right: 6px; }}
 .runs {{ font-size: 12px; color: #555; }}
</style></head><body>
<h1>{title}</h1>
<div class="runs">{runs}</div>
{sections}
<script>
function makePicker(divId, pickerId, raw, label, initialBg) {{
  const gd = document.getElementById(divId), root = document.getElementById(pickerId);
  // Background subtraction: point j of each trace is matched to the background
  // well's reading from the same run and cycle (key "sheet|cycle").
  const bgSel = root.querySelector('select.bg');
  const applyBg = () => {{
    const bg = bgSel.value, ref = bg ? raw.traces[bg] : null, lut = {{}};
    if (ref) ref.k.forEach((k, j) => {{ if (k !== null) lut[k] = ref.y[j]; }});
    const ys = [], hts = [];
    gd.data.forEach(t => {{
      const r = raw.traces[t.name];
      ys.push(r.y.map((v, j) => {{
        if (!ref) return v;
        const b = lut[r.k[j]];
        return (v === null || b === undefined || b === null) ? null : v - b;
      }}));
      hts.push(ref ? r.ht.replace(label + ' =', label + ' − ' + bg + ' =') : r.ht);
    }});
    Plotly.update(gd, {{y: ys, hovertemplate: hts}}, {{'yaxis.title.text': ref ? label + ' − ' + bg : label}});
  }};
  bgSel.addEventListener('change', applyBg);
  if (initialBg && raw.traces[initialBg]) {{ bgSel.value = initialBg; applyBg(); }}
  const colors = gd._fullData.map(t => t.line.color);
  const idx = {{}}; gd.data.forEach((t, i) => idx[t.name] = i);
  const isOn = i => gd.data[i].visible === undefined || gd.data[i].visible === true;
  const set = (ids, on) => {{ if (ids.length) Plotly.restyle(gd, {{visible: on ? true : 'legendonly'}}, ids); }};
  const refresh = () => root.querySelectorAll('button.well').forEach(b => {{
    const i = idx[b.dataset.well]; if (i === undefined) return;
    const on = isOn(i); b.classList.toggle('on', on); b.style.background = on ? colors[i] : '#fff';
  }});
  root.addEventListener('click', e => {{
    const b = e.target.closest('button'); if (!b) return;
    let ids = [];
    if (b.dataset.well) ids = [idx[b.dataset.well]];
    else if (b.dataset.row) ids = Object.keys(idx).filter(w => w.match(/^[A-P]/)[0] === b.dataset.row).map(w => idx[w]);
    else if (b.dataset.col) ids = Object.keys(idx).filter(w => w.slice(1) === b.dataset.col).map(w => idx[w]);
    else if (b.dataset.all) {{ set(Object.values(idx), b.dataset.all === '1'); return; }}
    // Toggle group: if any is shown, hide all of them; otherwise show all.
    set(ids, !ids.some(isOn));
  }});
  gd.on('plotly_restyle', refresh);
  refresh();
}}
{init}
</script></body></html>
"""


def _picker_html(pid: str, wells: list[str]) -> str:
    keys = [_well_key(w) for w in wells]
    rows = sorted({r for r, _ in keys})
    cols = sorted({c for _, c in keys})
    present = set(wells)
    h = [f'<div class="picker" id="{pid}"><table><tr><td></td>']
    h += [f'<td><button class="hdr" data-col="{c}" title="toggle column {c}">{c}</button></td>' for c in cols]
    h.append("</tr>")
    for r in rows:
        h.append(f'<tr><td><button class="hdr" data-row="{r}" title="toggle row {r}">{r}</button></td>')
        for c in cols:
            w = f"{r}{c}"
            dis = "" if w in present else " disabled"
            h.append(f'<td><button class="well" data-well="{w}"{dis}>{w}</button></td>')
        h.append("</tr>")
    h.append('</table><div class="controls" style="margin-top:6px">'
             '<button data-all="1">Show all</button><button data-all="0">Hide all</button>'
             '<label style="margin-left:12px">Background well: <select class="bg">'
             '<option value="">none</option>'
             + "".join(f'<option value="{w}">{w}</option>' for w in wells)
             + '</select></label></div></div>')
    return "".join(h)


def _raw_traces(fig: go.Figure) -> dict:
    """Unsubtracted y values, run|cycle match keys and hover templates, per well."""
    def clean(v):
        return None if v is None or (isinstance(v, float) and np.isnan(v)) else float(v)
    return {"traces": {
        t.name: {"y": [clean(v) for v in t.y],
                 "k": [f"{c[0]}|{c[2]}" if c[0] is not None else None for c in t.customdata],
                 "ht": t.hovertemplate}
        for t in fig.data}}


def write_html(data: pd.DataFrame, out: Path, title: str, cdn: bool, blank: str | None = None) -> None:
    runs = (data.groupby(["sheet", "run_start"], as_index=False)
            .agg(cycles=("cycle", "nunique"), last_s=("time_s", "max"))
            .sort_values("run_start"))
    run_txt = " &nbsp;|&nbsp; ".join(
        f"{r.sheet}: {r.run_start:%Y-%m-%d %H:%M:%S}, {r.cycles} cycles, {r.last_s / 3600:.2f} h"
        for r in runs.itertuples())
    sections, inits = [], []
    first = True
    for k, label in enumerate(sorted(data["label"].unique())):
        fig = build_figure(data, label, runs)
        div_id, pid = f"plot{k}", f"picker{k}"
        plot_html = fig.to_html(full_html=False, div_id=div_id,
                                include_plotlyjs=("cdn" if cdn else True) if first else False,
                                config={"displaylogo": False, "toImageButtonOptions": {"format": "svg"}})
        first = False
        wells = sorted(data.loc[data["label"] == label, "well"].unique(), key=_well_key)
        sections.append(f"<h2 style='font-size:16px'>{label}</h2>{_picker_html(pid, wells)}{plot_html}")
        inits.append(f"makePicker({json.dumps(div_id)}, {json.dumps(pid)}, {json.dumps(_raw_traces(fig))}, "
                     f"{json.dumps(label)}, {json.dumps(blank)});")
    out.write_text(PAGE.format(title=title, runs=run_txt, sections="\n".join(sections),
                               init="window.addEventListener('load', () => {" + " ".join(inits) + "});"),
                   encoding="utf-8")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx", type=Path, help="Tecan i-control Excel export (one sheet per run)")
    ap.add_argument("-o", "--out", type=Path, help="output HTML (default: <xlsx>_plot.html)")
    ap.add_argument("--csv", type=Path, help="output long-format CSV (default: <xlsx>_long.csv)")
    ap.add_argument("--no-csv", action="store_true", help="don't write the CSV")
    ap.add_argument("--cdn", action="store_true", help="load plotly.js from CDN (smaller file, needs internet)")
    ap.add_argument("--open", action="store_true", help="open the HTML in the browser when done")
    ap.add_argument("--blank", metavar="WELL",
                    help="background well: preselected in the plot, and adds a value_bgsub column to the CSV")
    args = ap.parse_args(argv)

    data = load_workbook(args.xlsx)
    if args.blank:
        args.blank = args.blank.upper()
        if args.blank not in set(data["well"]):
            raise SystemExit(f"--blank {args.blank}: no such well in the data")
        bg = (data[data["well"] == args.blank][["label", "sheet", "cycle", "value"]]
              .rename(columns={"value": "_bg"}))
        data = data.merge(bg, on=["label", "sheet", "cycle"], how="left")
        data["value_bgsub"] = data["value"] - data.pop("_bg")
    stem = args.xlsx.with_suffix("")
    out = args.out or Path(f"{stem}_plot.html")
    write_html(data, out, title=args.xlsx.name, cdn=args.cdn, blank=args.blank)
    print(f"wrote {out}")
    if not args.no_csv:
        csv = args.csv or Path(f"{stem}_long.csv")
        data.to_csv(csv, index=False)
        print(f"wrote {csv}")
    for (sheet, start), g in data.groupby(["sheet", "run_start"], sort=False):
        print(f"  {sheet}: start {start}, {g['cycle'].nunique()} cycles, wells {', '.join(sorted(g['well'].unique(), key=_well_key))}")
    if args.open:
        webbrowser.open(out.resolve().as_uri())


if __name__ == "__main__":
    main()
