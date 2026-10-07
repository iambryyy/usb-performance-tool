"""Formatted Excel workbook for a test run, plus a matching blank template.

The Summary and Analysis sheets are driven by formulas that read the Results
sheet, so the same workbook layout works two ways: filled in by the tool after
a run, or filled in by hand from the blank template (type rows into Results
and every total and chart updates). Each formula is also written with its
calculated value, so numbers and charts show correctly in viewers that don't
recalculate (e.g. Excel's Protected View on a downloaded file).
"""
import os
import re
from dataclasses import asdict, is_dataclass
from datetime import datetime

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

RESULTS_FILE = "usb_performance_results.xlsx"
TEMPLATE_FILE = "USB_Performance_Results_TEMPLATE.xlsx"

NAVY = "#0B1F3A"
BLUE = "#1F3F73"
ORANGE = "#C9A227"
GREEN = "#70AD47"
RED = "#E15759"
AMBER = "#FFC000"
GRAY_TXT = "#7F7F7F"
PANEL = "#F6F3EA"
INPUT_FILL = "#FFF9E5"
FONT = "Calibri"

INTERFACE_ORDER = ["1Gbps", "2.5Gbps", "10Gbps", "SFP Port", "2.4GHz", "5GHz", "6GHz", "MLO"]
FORMAT_ORDER = ["FAT32", "NTFS", "exFAT", "HFS+", "EXT4"]
FILE_ORDER = ["small_1mb.bin", "document_10mb.bin", "image_50mb.bin", "video_100mb.bin", "large_500mb.bin"]
ROLE_ORDER = ["Router", "Extender"]
CONNECTION_ORDER = ["wired", "wireless"]
TEMPLATE_ROUNDS = 3
SHEET_ORDER = ("Summary", "Results", "Issues", "Analysis", "Notes")

# Results sheet geometry (0-indexed rows). Formulas read rows up to the
# capacity below, so rows added later (or typed into the template) count.
RES_TITLE_ROW = 0
RES_HINT_ROW = 1
RES_BANNER_ROW = 2
RES_HEADER_ROW = 3
RES_FIRST = 4
MIN_CAPACITY = 5000  # result rows the formulas read; grows with the data so a big run is never silently under-counted
_state = {"capacity": MIN_CAPACITY}

# key, header, width, kind, group
COLUMNS = [
    ("dut_role", "DUT Role", 9, "center", "case"),
    ("configured_interface", "Interface", 11, "center", "case"),
    ("connection_type", "Connection", 11, "center", "case"),
    ("usb_format", "USB Format", 11, "center", "case"),
    ("direction", "Direction", 10, "center", "case"),
    ("file_name", "File", 19, "text", "case"),
    ("iteration", "Round", 7, "int", "case"),
    ("throughput_mbps", "Throughput (MB/s)", 15, "num", "measure"),
    ("duration_seconds", "Duration (s)", 11, "num3", "measure"),
    ("average_mbps", "Average (MB/s)", 13, "num", "measure"),
    ("min_mbps", "Min (MB/s)", 11, "num", "measure"),
    ("max_mbps", "Max (MB/s)", 11, "num", "measure"),
    ("stddev_mbps", "Std Dev", 9, "num", "measure"),
    ("status", "Result", 12, "center", "result"),
    ("failure_reason", "Failure Reason", 58, "wrap", "result"),
    ("timestamp", "Time", 19, "datetime", "trace"),
    ("usb_share_path", "USB Share", 22, "text", "trace"),
    ("actual_ssid", "SSID", 14, "text", "trace"),
    ("actual_band", "Actual Band", 11, "center", "trace"),
    ("actual_channel", "Channel", 9, "center", "trace"),
    ("wifi_signal_percent", "Signal", 8, "center", "trace"),
    ("bssid", "BSSID", 18, "text", "trace"),
    ("client_ip", "Client IP", 15, "text", "trace"),
    ("detected_gateway_ip", "Gateway IP", 15, "text", "trace"),
    ("file_verification", "File Check", 16, "text", "trace"),
    ("file_set_version", "File Set", 9, "center", "trace"),
    ("actual_usb_format", "Detected USB Format", 15, "center", "trace"),
    ("dut_model", "DUT Model", 11, "center", "trace"),
    ("dut_name", "DUT Name", 14, "text", "trace"),
    ("firmware_version", "Firmware", 32, "text", "trace"),
    ("detected_hardware_version", "HW Version", 11, "center", "trace"),
    ("topology", "Topology", 38, "text", "trace"),
    ("client_type", "Client Type", 12, "text", "trace"),
    ("client_os", "Client OS", 12, "text", "trace"),
    ("project_name", "Project", 30, "text", "trace"),
    ("interface_match", "Interface Matched", 11, "center", "trace"),
    ("role_assumed", "Role Assumed", 11, "center", "trace"),
    ("dut_reachable", "DUT Reachable", 11, "center", "trace"),
    ("share_write_access", "Share Writable", 11, "center", "trace"),
]
COL = {key: i for i, (key, *_rest) in enumerate(COLUMNS)}
LAST_VISIBLE_COL = COL["failure_reason"]
BOOLEAN_KEYS = {"interface_match", "role_assumed", "dut_reachable", "share_write_access"}
MEASURE_KEYS = {"throughput_mbps", "duration_seconds", "average_mbps", "min_mbps", "max_mbps", "stddev_mbps"}


class _Styles:
    """Cache of formats so identical looks share one format object."""

    def __init__(self, workbook):
        self.workbook = workbook
        self._cache = {}

    def __call__(self, **props):
        props = {"font_name": FONT, "font_size": 10, "valign": "vcenter", **props}
        key = tuple(sorted(props.items()))
        if key not in self._cache:
            self._cache[key] = self.workbook.add_format(props)
        return self._cache[key]


def _as_dict(result):
    if isinstance(result, dict):
        return dict(result)
    if is_dataclass(result):
        return asdict(result)
    return dict(vars(result))


def _ordered(values, canonical):
    present = {v for v in values if v}
    return [c for c in canonical if c in present] + sorted(v for v in present if v not in canonical)


def _parse_time(text):
    try:
        return datetime.fromisoformat(str(text))
    except (TypeError, ValueError):
        return None


def _rng(key):
    """Results-sheet range for a column, e.g. Results!$O$5:$O$2004."""
    letter = xl_col_to_name(COL[key])
    return f"Results!${letter}${RES_FIRST + 1}:${letter}${RES_FIRST + _state['capacity']}"


# ---------------------------------------------------------------- calculations
# These mirror the sheet formulas exactly; they supply each formula's cached value.

def _ran(rows):
    return [r for r in rows if r["status"] != "SKIPPED"]


NA = "–"  # shown instead of 0 when there is nothing to average (e.g. a band that was never tested)


def _avg(rows, direction, **where):
    vals = [
        float(r["throughput_mbps"] or 0) for r in rows
        if r["direction"] == direction and float(r["duration_seconds"] or 0) > 0
        and all(r[k] == v for k, v in where.items())
    ]
    return sum(vals) / len(vals) if vals else NA


def _peak(rows, direction):
    vals = [float(r["throughput_mbps"] or 0) for r in rows if r["direction"] == direction]
    return max(vals) if vals else NA


def _size_label(name):
    """'document_10mb.bin' -> '10 MB' (short enough to sit under a chart bar)."""
    m = re.search(r"(\d+(?:\.\d+)?)\s*(mb|gb|kb)", name or "", re.IGNORECASE)
    return f"{float(m.group(1)):g} {m.group(2).upper()}" if m else os.path.splitext(name or "")[0]


def _verdict(rounds, failed, skipped):
    if rounds + skipped == 0:
        return "NO DATA"
    if failed > 0:
        return "FAIL"
    return "INCOMPLETE" if skipped > 0 else "PASS"


def _threshold_lines(config):
    """(label, text) per line: by file size where a table is configured, else the flat values."""
    config = config or {}
    table = config.get("thresholds_by_file_mbps") or {}
    flat = config.get("thresholds_mbps") or {}
    lines = []
    for connection in ("wired", "wireless"):
        by_size = {d: {float(k): float(v) for k, v in ((table.get(d) or {}).get(connection) or {}).items()}
                   for d in ("read", "write")}
        if any(by_size.values()):
            sizes = sorted({s for per in by_size.values() for s in per})
            row = lambda d: " / ".join(f"{by_size[d][s]:g}" if s in by_size[d] else "-" for s in sizes)
            lines.append((f"Pass thresholds ({connection})",
                          f"By file size ({' / '.join(f'{s:g}' for s in sizes)} MB):  "
                          f"read ≥ {row('read')}  ·  write ≥ {row('write')} MB/s"))
        else:
            r, w = (flat.get("read") or {}).get(connection), (flat.get("write") or {}).get(connection)
            if r is not None and w is not None:
                lines.append((f"Pass thresholds ({connection})", f"Read ≥ {r:g}  ·  Write ≥ {w:g} MB/s"))
    if config.get("thresholds_note"):
        lines.append(("Threshold basis", str(config["thresholds_note"])))
    return lines


def _threshold_text(config):
    t = (config or {}).get("thresholds_mbps") or {}
    read, write = t.get("read") or {}, t.get("write") or {}

    def part(label, d):
        wired, wireless = d.get("wired"), d.get("wireless")
        if wired is None and wireless is None:
            return ""
        if wired == wireless:
            return f"{label} ≥ {wired:g} MB/s"
        return f"{label} ≥ {wired:g} MB/s wired, {wireless:g} MB/s wireless"

    return "  ·  ".join(p for p in (part("Read", read), part("Write", write)) if p)


def _span_text(rows):
    times = sorted(t for t in (_parse_time(r["timestamp"]) for r in rows) if t)
    if not times:
        return ""
    first, last = times[0], times[-1]
    minutes = max(1, round((last - first).total_seconds() / 60))
    end = last.strftime("%H:%M") if first.date() == last.date() else last.strftime("%Y-%m-%d %H:%M")
    return f"{first:%Y-%m-%d  %H:%M} – {end}  ({minutes} min)"


def _unique(rows, key):
    seen = []
    for r in rows:
        v = r.get(key)
        if v not in (None, "") and v not in seen:
            seen.append(v)
    return seen


# ---------------------------------------------------------------------- charts

CHART_PITCH = 15  # sheet rows per chart row (20 px each); the chart itself is 290 px tall


def _style_chart(chart, title):
    chart.set_title({"name": title, "name_font": {"name": FONT, "size": 12, "bold": True, "color": NAVY}})
    chart.set_chartarea({"border": {"color": "#D9D9D9"}, "fill": {"color": "#FFFFFF"}})
    chart.set_plotarea({"fill": {"none": True}})
    chart.set_size({"width": 570, "height": 290})


def _column_chart(wb, title, cats, read, write, n_categories, peak=None):
    """peak = tallest bar (MB/s) so the axis labels can be whole numbers; None = no data yet (blank template)."""
    chart = wb.add_chart({"type": "column"})
    gap = 70 if n_categories >= 4 else 150  # fat bars look clumsy when there are only one or two groups
    for name, values, color in (("Read (MB/s)", read, BLUE), ("Write (MB/s)", write, ORANGE)):
        chart.add_series({
            "name": name, "categories": cats, "values": values,
            "fill": {"color": color}, "border": {"none": True},
            "gap": gap, "overlap": -5,
            "data_labels": {"value": True, "num_format": "0.0;;;", "font": {"name": FONT, "size": 8, "color": "#404040"}},
        })
    _style_chart(chart, title)
    chart.set_x_axis({"num_font": {"name": FONT, "size": 9}, "line": {"color": "#BFBFBF"}, "major_tick_mark": "none"})
    # Blank template: the axis must work for any speed the user later enters, so no fixed max or step; this
    # number format shows 0.0-0.9 with a decimal (the empty chart's 0-1 scale) and whole numbers above that.
    num_format = "[<1]0.0;0" if peak is None else ("0" if peak >= 10 else "0.0")
    chart.set_y_axis({
        "name": "MB/s", "name_font": {"name": FONT, "size": 9, "bold": False},
        "num_font": {"name": FONT, "size": 9}, "min": 0, "line": {"none": True}, "num_format": num_format,
        "major_gridlines": {"visible": True, "line": {"color": "#E7E6E6", "width": 0.75}},
    })
    chart.set_legend({"position": "bottom", "font": {"name": FONT, "size": 9}})
    return chart


def _outcome_chart(wb, cats, values):
    chart = wb.add_chart({"type": "doughnut"})
    chart.add_series({
        "name": "Outcome", "categories": cats, "values": values,
        "points": [{"fill": {"color": c}, "border": {"color": "white", "width": 1.5}} for c in (GREEN, RED, AMBER)],
        "data_labels": {"value": True, "num_format": "0;;;", "font": {"name": FONT, "size": 12, "bold": True, "color": "#1F1F1F"}},
    })
    chart.set_hole_size(55)
    _style_chart(chart, "Test Outcome (rounds)")
    chart.set_legend({"position": "bottom", "font": {"name": FONT, "size": 9}})
    return chart


# ----------------------------------------------------------------------- sheets

def _page_setup(ws, landscape=True):
    if landscape:
        ws.set_landscape()
    ws.set_paper(1)
    ws.fit_to_pages(1, 0)


def _title_bar(ws, st, row, c1, c2, text):
    ws.merge_range(row, c1, row, c2, text, st(bold=True, font_size=14, font_color="#FFFFFF", bg_color=NAVY, indent=1))
    ws.set_row(row, 28)


def _build_results_sheet(ws, st, rows):
    ws.set_tab_color(BLUE)
    ws.outline_settings(True, True, False, False)  # the [+] button sits left of the grouped columns
    _title_bar(ws, st, RES_TITLE_ROW, 0, LAST_VISIBLE_COL, "TEST RESULTS — DETAILED DATA")
    ws.merge_range(RES_HINT_ROW, 0, RES_HINT_ROW, LAST_VISIBLE_COL,
                   "Use the arrows on the header row to filter or sort. Time, SSID, BSSID, firmware and other traceability "
                   "columns are grouped — click the [+] above 'Failure Reason' to show them.",
                   st(italic=True, font_color=GRAY_TXT, indent=1))

    groups = {"case": ("TEST CASE", "#DFE5F0", NAVY), "measure": ("MEASUREMENT", "#F3EAC8", "#6B5412"),
              "result": ("RESULT", "#E8DFCF", "#5A3E1B")}
    for gkey, (label, fill, color) in groups.items():
        cols = [i for i, c in enumerate(COLUMNS) if c[4] == gkey]
        ws.merge_range(RES_BANNER_ROW, cols[0], RES_BANNER_ROW, cols[-1], label,
                       st(bold=True, font_size=9, align="center", bg_color=fill, font_color=color, border=1, border_color="#FFFFFF"))

    line = {"bottom": 1, "bottom_color": "#E7E6E6"}  # thin row separator; the table style is off so striping stays light
    cell_fmts = {
        "text": st(**line), "center": st(align="center", **line), "wrap": st(text_wrap=True, valign="top", **line),
        "int": st(align="center", num_format="0", **line), "num": st(num_format="0.00", **line),
        "num3": st(num_format="0.000", **line), "datetime": st(num_format="yyyy-mm-dd hh:mm:ss", align="center", **line),
    }
    header_fmt = st(bold=True, font_color="#FFFFFF", bg_color=NAVY, text_wrap=True, align="center", border=1, border_color="#FFFFFF")
    table_cols = [{"header": h, "format": cell_fmts[kind], "header_format": header_fmt} for _k, h, _w, kind, _g in COLUMNS]

    data = []
    for r in rows:
        skipped = r["status"] == "SKIPPED"
        line = []
        for key, *_rest in COLUMNS:
            v = r.get(key)
            if key in BOOLEAN_KEYS:
                v = "Yes" if v else "No"
            elif key == "timestamp":
                v = _parse_time(v) or v
            elif (skipped and key in MEASURE_KEYS) or (key == "iteration" and (skipped or not v)):
                v = None
            line.append(v if v != "" else None)
        data.append(line)

    ws.add_table(RES_HEADER_ROW, 0, RES_HEADER_ROW + max(1, len(data)), len(COLUMNS) - 1, {
        "name": "ResultsTable", "style": None, "columns": table_cols, "data": data,
    })
    ws.set_row(RES_HEADER_ROW, 32)

    for i, (_key, _h, width, _kind, group) in enumerate(COLUMNS):
        if group == "trace":
            ws.set_column(i, i, width, None, {"level": 1, "hidden": True})
        elif i == LAST_VISIBLE_COL:
            ws.set_column(i, i, width, None, {"collapsed": True})
        else:
            ws.set_column(i, i, width)

    last = RES_FIRST + _state["capacity"] - 1
    status_col = COL["status"]
    status_letter = xl_col_to_name(status_col)
    for value, bg in (("PERF_FAIL", "#FDECEA"), ("ERROR", "#FDECEA"), ("SKIPPED", "#FFF8E1")):
        ws.conditional_format(RES_FIRST, 0, last, LAST_VISIBLE_COL, {
            "type": "formula", "criteria": f'=${status_letter}{RES_FIRST + 1}="{value}"', "format": st(bg_color=bg),
        })
    for value, bg, fg in (("PASS", "#C6EFCE", "#006100"), ("PERF_FAIL", "#FFC7CE", "#9C0006"),
                          ("ERROR", "#FFC7CE", "#9C0006"), ("SKIPPED", "#FFEB9C", "#7F6000")):
        ws.conditional_format(RES_FIRST, status_col, last, status_col, {
            "type": "cell", "criteria": "==", "value": f'"{value}"', "format": st(bg_color=bg, font_color=fg, bold=True),
        })
    thr = COL["throughput_mbps"]
    ws.conditional_format(RES_FIRST, thr, last, thr, {
        "type": "data_bar", "bar_color": "#A9B8D6", "bar_solid": True, "min_type": "num", "min_value": 0,
    })
    # Faint zebra striping, created last so the outcome tints above win on failed/skipped rows.
    ws.conditional_format(RES_FIRST, 0, last, LAST_VISIBLE_COL, {
        "type": "formula", "criteria": f"=AND($A{RES_FIRST + 1}<>\"\",MOD(ROW(),2)=0)", "format": st(bg_color="#F7F5EE"),
    })

    def pick(key, options):
        c = COL[key]
        ws.data_validation(RES_FIRST, c, last, c, {
            "validate": "list", "source": options, "error_type": "warning",
            "error_title": "Unexpected value", "error_message": "That isn't one of the usual values. Keep it anyway?",
        })

    pick("dut_role", ROLE_ORDER)
    pick("connection_type", CONNECTION_ORDER)
    pick("direction", ["read", "write"])
    pick("usb_format", FORMAT_ORDER)
    pick("configured_interface", INTERFACE_ORDER)
    pick("status", ["PASS", "PERF_FAIL", "ERROR", "SKIPPED"])

    ws.freeze_panes(RES_FIRST, 2)
    ws.repeat_rows(RES_HEADER_ROW)
    _page_setup(ws)


def _build_analysis_sheet(wb, ws, st, rows, template):
    """Tables the Summary charts read from. Returns where each table landed."""
    ws.set_tab_color("#C9A227")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 22)
    ws.set_column(1, 7, 13)
    ws.set_column(8, 8, 12)
    _title_bar(ws, st, 0, 0, 7, "ANALYSIS — THE NUMBERS BEHIND THE CHARTS")
    ws.merge_range(1, 0, 1, 7, "Every value is a formula reading the Results sheet. Change a label in the first column "
                   "to analyse a different value.", st(italic=True, font_color=GRAY_TXT, indent=1))

    ran = _ran(rows)
    pass_fmt, num_fmt = st(num_format="0%", align="center"), st(num_format="0.00", align="center")
    count_fmt, label_fmt = st(num_format="0", align="center"), st(bold=True, indent=1)
    head_fmt = st(bold=True, font_color="#FFFFFF", bg_color=NAVY, text_wrap=True, align="center", border=1, border_color="#FFFFFF")
    sect_fmt = st(bold=True, font_size=11, font_color=NAVY, bottom=2, bottom_color=NAVY)
    thr, dur, status, direction = _rng("throughput_mbps"), _rng("duration_seconds"), _rng("status"), _rng("direction")
    layout = {}
    row = 3

    def dimension_block(name, title, first_head, key, labels):
        nonlocal row
        dim = _rng(key)
        counted = {lab: sum(1 for r in ran if r[key] == lab) for lab in labels}
        labels = [l for l in labels if counted[l] > 0] + [l for l in labels if counted[l] == 0]
        ws.merge_range(row, 0, row, 7, title, sect_fmt)
        for c, h in enumerate((first_head, "Rounds Run", "Passed", "Failed", "Skipped", "Pass Rate", "Avg Read (MB/s)", "Avg Write (MB/s)")):
            ws.write(row + 1, c, h, head_fmt)
        ws.set_row(row + 1, 30)
        first = row + 2
        for i, lab in enumerate(labels):
            r0, n = first + i, first + i + 1
            sel = [r for r in rows if r[key] == lab]
            rounds = len([r for r in sel if r["status"] != "SKIPPED"])
            passed = len([r for r in sel if r["status"] == "PASS"])
            skipped = len([r for r in sel if r["status"] == "SKIPPED"])
            ws.write(r0, 0, lab, label_fmt)
            ws.write_formula(r0, 1, f'=COUNTIFS({dim},$A{n},{status},"<>SKIPPED")', count_fmt, rounds)
            ws.write_formula(r0, 2, f'=COUNTIFS({dim},$A{n},{status},"PASS")', count_fmt, passed)
            ws.write_formula(r0, 3, f"=B{n}-C{n}", count_fmt, rounds - passed)
            ws.write_formula(r0, 4, f'=COUNTIFS({dim},$A{n},{status},"SKIPPED")', count_fmt, skipped)
            ws.write_formula(r0, 5, f'=IFERROR(C{n}/B{n},"{NA}")', pass_fmt, passed / rounds if rounds else NA)
            for c, d in ((6, "read"), (7, "write")):
                ws.write_formula(r0, c, f'=IFERROR(AVERAGEIFS({thr},{dim},$A{n},{direction},"{d}",{dur},">0"),"{NA}")',
                                 num_fmt, _avg(sel, d))
            if key == "file_name":
                ws.write(r0, 8, _size_label(lab), st(font_color=GRAY_TXT, italic=True, align="center"))
        last = first + len(labels) - 1
        if key == "file_name":
            ws.write(row + 1, 8, "Chart Label", head_fmt)
        ws.conditional_format(first, 5, last, 5, {
            "type": "3_color_scale", "min_type": "num", "min_value": 0, "mid_type": "num", "mid_value": 0.5,
            "max_type": "num", "max_value": 1, "min_color": "#F8CBAD", "mid_color": "#FFE699", "max_color": "#C6E0B4"})
        for c, color in ((6, "#A9B8D6"), (7, "#E8D28A")):
            ws.conditional_format(first, c, last, c, {"type": "data_bar", "bar_color": color, "bar_solid": True,
                                                      "min_type": "num", "min_value": 0})
        # Charts only cover labels that actually ran, so a skipped band isn't drawn as a zero-speed bar.
        measured = sum(1 for l in labels if counted[l] > 0) or len(labels)
        speeds = [v for lab in labels[:measured] for d in ("read", "write")
                  for v in [_avg([r for r in rows if r[key] == lab], d)] if v != NA]
        layout[name] = {"first": first, "last": first + measured - 1, "labels": labels[:measured],
                        "peak": max(speeds) if speeds else None, "start": row}
        row = last + 3

    def present(key, canonical, source):
        return canonical if template else _ordered([r[key] for r in source], canonical)

    dimension_block("interface", "BY INTERFACE — wired port speed or wireless band", "Interface", "configured_interface",
                    present("configured_interface", INTERFACE_ORDER, rows))
    dimension_block("format", "BY USB FORMAT — file system on the drive", "USB Format", "usb_format",
                    present("usb_format", FORMAT_ORDER, ran))
    dimension_block("file", "BY FILE — transfer size", "File", "file_name", present("file_name", FILE_ORDER, ran))
    dimension_block("role", "BY DUT ROLE — Router vs Extender", "DUT Role", "dut_role", present("dut_role", ROLE_ORDER, rows))
    dimension_block("connection", "BY CONNECTION — wired vs wireless", "Connection", "connection_type",
                    present("connection_type", CONNECTION_ORDER, rows))

    ws.merge_range(row, 0, row, 7, "OUTCOME — every round that was planned", sect_fmt)
    ws.write(row + 1, 0, "Outcome", head_fmt)
    ws.write(row + 1, 1, "Rounds", head_fmt)
    passed_n = sum(1 for r in ran if r["status"] == "PASS")
    first_outcome = row + 2
    for i, (lab, formula, val) in enumerate((
        ("Passed", f'=COUNTIF({status},"PASS")', passed_n),
        ("Failed", f'=COUNTA({status})-COUNTIF({status},"SKIPPED")-B{first_outcome + 1}', len(ran) - passed_n),
        ("Skipped", f'=COUNTIF({status},"SKIPPED")', len(rows) - len(ran)),
    )):
        ws.write(first_outcome + i, 0, lab, label_fmt)
        ws.write_formula(first_outcome + i, 1, formula, count_fmt, val)
    layout["outcome"] = {"first": first_outcome, "last": first_outcome + 2}
    row = first_outcome + 5

    rounds_n = TEMPLATE_ROUNDS if template else max([int(r["iteration"] or 0) for r in rows] + [1])
    ws.merge_range(row, 0, row, 7, "PER ROUND — average across every file, interface and format", sect_fmt)
    for c, h in enumerate(("Round", "#", "Avg Read (MB/s)", "Avg Write (MB/s)")):
        ws.write(row + 1, c, h, head_fmt)
    rr_first = row + 2
    iteration = _rng("iteration")
    for i in range(rounds_n):
        r0, n = rr_first + i, rr_first + i + 1
        ws.write(r0, 0, f"Round {i + 1}", label_fmt)
        ws.write_number(r0, 1, i + 1, st(font_color=GRAY_TXT, align="center"))
        sel = [r for r in rows if int(r["iteration"] or 0) == i + 1]
        for c, d in ((2, "read"), (3, "write")):
            ws.write_formula(r0, c, f'=IFERROR(AVERAGEIFS({thr},{direction},"{d}",{iteration},$B{n},{dur},">0"),"{NA}")',
                             num_fmt, _avg(sel, d))
    per_round = [v for i in range(rounds_n) for d in ("read", "write")
                 for v in [_avg([r for r in rows if int(r["iteration"] or 0) == i + 1], d)] if v != NA]
    layout["rounds"] = {"first": rr_first, "last": rr_first + rounds_n - 1, "count": rounds_n,
                        "peak": max(per_round) if per_round else None}
    # Page 1 = interface, format, file; page 2 = the rest. A fixed scale keeps these manual breaks in effect.
    ws.set_h_pagebreaks([layout["role"]["start"]])
    ws.set_landscape()
    ws.set_paper(1)
    ws.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.set_print_scale(85)
    return layout


def _build_issues_sheet(ws, st, rows, template):
    ws.set_tab_color(RED)
    ws.hide_gridlines(2)
    heads = [("Result", 12), ("DUT Role", 10), ("Interface", 12), ("USB Format", 11), ("Direction", 10),
             ("File", 20), ("Round", 8), ("Reason", 95)]
    for c, (_h, w) in enumerate(heads):
        ws.set_column(c, c, w)
    _title_bar(ws, st, 0, 0, len(heads) - 1, "FAILED, SKIPPED AND ABNORMAL CASES")

    issues = []
    for r in rows:
        if r["status"] in ("ERROR", "PERF_FAIL", "NOT_RUN"):
            issues.append((0 if r["status"] == "ERROR" else 1, r["status"], r))
        elif r["status"] == "SKIPPED":
            issues.append((3, "SKIPPED", r))
        elif r.get("interface_match") is False:
            issues.append((2, "ABNORMAL", r))
    issues.sort(key=lambda t: t[0])

    clean = not template and not issues
    if template:
        note = "List failed, skipped or abnormal rounds here (or filter the Results sheet by Result)."
    elif clean:
        note = "None. Every configured case ran and met its threshold."
    else:
        counts = {}
        for _sev, label, _r in issues:
            counts[label] = counts.get(label, 0) + 1
        note = f"{len(issues)} case(s) need attention:  " + ",  ".join(
            f"{n} {k.lower().replace('_', ' ')}" for k, n in counts.items())
    note_fmt = st(indent=1, bold=True, font_color="#375623", bg_color="#E2EFDA") if clean else st(indent=1, italic=True, font_color=GRAY_TXT)
    ws.merge_range(1, 0, 1, len(heads) - 1, note, note_fmt)
    if clean:
        _page_setup(ws)
        return

    data = []
    for _sev, label, r in issues:
        reason = r.get("failure_reason") or (
            f"Connected band {r.get('actual_band') or '?'} does not match configured interface {r.get('configured_interface')}")
        data.append([label, r["dut_role"] or None, r["configured_interface"] or None, r["usb_format"] or None,
                     r["direction"] or None, r["file_name"] or None, r["iteration"] or None, reason])
    head_fmt = st(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", border=1, border_color="#FFFFFF")
    line = {"bottom": 1, "bottom_color": "#E7E6E6"}
    center, wrap = st(align="center", **line), st(text_wrap=True, valign="top", **line)
    cols = [{"header": h, "header_format": head_fmt, "format": wrap if h == "Reason" else (st(**line) if h == "File" else center)}
            for h, _w in heads]
    last = 3 + max(1, len(data))
    ws.add_table(3, 0, last, len(heads) - 1, {"name": "IssuesTable", "style": None, "columns": cols, "data": data})
    for value, bg, fg in (("PERF_FAIL", "#FFC7CE", "#9C0006"), ("ERROR", "#FFC7CE", "#9C0006"),
                          ("ABNORMAL", "#FFEB9C", "#7F6000"), ("SKIPPED", "#FFEB9C", "#7F6000")):
        ws.conditional_format(4, 0, last, 0, {"type": "cell", "criteria": "==", "value": f'"{value}"',
                                              "format": st(bg_color=bg, font_color=fg, bold=True)})
    ws.freeze_panes(4, 0)
    _page_setup(ws)


def _build_notes_sheet(ws, st, template):
    ws.set_tab_color("#A6A6A6")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 2, 100)
    _title_bar(ws, st, 0, 1, 2, "NOTES — HOW TO READ THIS WORKBOOK")
    sect = st(bold=True, font_size=11, font_color=NAVY, bottom=2, bottom_color=NAVY)
    key_fmt, text_fmt = st(bold=True, valign="top", indent=1), st(text_wrap=True, valign="top")
    row = 2

    def code_fmt(bg, fg):
        return st(bold=True, valign="top", indent=1, bg_color=bg, font_color=fg)

    def section(title, items):
        nonlocal row
        ws.merge_range(row, 1, row, 2, title, sect)
        row += 1
        for key, text, fmt in items:
            ws.write(row, 1, key, fmt or key_fmt)
            ws.write(row, 2, text, text_fmt)
            row += 1
        row += 1

    if template:
        section("USING THIS TEMPLATE", [
            ("1. Results sheet", "Enter one row per test round (one file, one direction, one round). Fill in the columns up to 'Failure "
                                 "Reason'; the grouped columns after it (time, SSID, firmware...) are optional traceability data.", None),
            ("2. Summary sheet", "Fill in the yellow cells: project, date, topology, client and devices under test. The totals, "
                                 "tiles and charts update by themselves from the Results sheet.", None),
            ("3. Analysis sheet", "The label in the first column of each table (e.g. '5GHz', 'NTFS') is what gets counted. Change or "
                                  "add labels there if your test uses different interfaces, formats or files.", None),
            ("4. Result column", "Use PASS, PERF_FAIL, ERROR or SKIPPED. Skipped rounds are counted separately, never as zero speed.", None),
            ("5. Issues sheet", "Copy failed, skipped or abnormal rounds here, or filter the Results sheet by Result.", None),
        ])
    section("SHEETS", [
        ("Summary", "One-page dashboard: overall result, headline numbers, charts, test environment and devices under test.", None),
        ("Results", "Every test round, one per row, with filters. Rows are tinted by outcome and the bar in the throughput "
                    "column shows speed at a glance.", None),
        ("Issues", "Only the rounds that failed, were skipped or looked abnormal, with the reason.", None),
        ("Analysis", "The tables behind the charts: averages, pass rates and counts by interface, USB format, file, role and connection.", None),
    ])
    section("RESULT CODES", [
        ("PASS", "The round ran and met the throughput threshold.", code_fmt("#C6EFCE", "#006100")),
        ("PERF_FAIL", "The round ran but was slower than the threshold.", code_fmt("#FFC7CE", "#9C0006")),
        ("ERROR", "The round could not complete (timeout, share error, no space...). The reason is in 'Failure Reason'.", code_fmt("#FFC7CE", "#9C0006")),
        ("SKIPPED", "Not tested at all (a band couldn't be joined, or the USB share wasn't reachable). Never counted as 0 MB/s.", code_fmt("#FFEB9C", "#7F6000")),
    ])
    section("HOW THE NUMBERS WORK", [
        ("Overall result", "FAIL if any round failed or errored. INCOMPLETE if nothing failed but something was skipped. PASS otherwise.", None),
        ("Averages", "Read and write averages only include rounds that actually ran; skipped and errored rounds are left out.", None),
        ("Pass rate", "Passed rounds divided by rounds that ran.", None),
        ("Thresholds", "Each round is judged against the pass thresholds listed on the Summary sheet. Very small files finish so "
                       "quickly that connection overhead dominates, so they miss speed thresholds more often than large files.", None),
    ])
    ws.write(row, 1, f"Generated {datetime.now():%Y-%m-%d %H:%M} by the USB Performance Automation Tool",
             st(italic=True, font_color=GRAY_TXT, indent=1))
    _page_setup(ws)


def _build_summary_sheet(wb, ws, st, rows, config, template, analysis):
    ws.set_tab_color(NAVY)
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 12, 13)
    ws.set_column(13, 13, 2)
    ran = _ran(rows)
    first = ran[0] if ran else (rows[0] if rows else {})
    cfg = config or {}

    ws.merge_range(0, 1, 0, 12, "USB PERFORMANCE TEST REPORT" + ("  —  BLANK TEMPLATE" if template else ""),
                   st(bold=True, font_size=20, font_color="#FFFFFF", bg_color=NAVY, indent=1))
    ws.set_row(0, 38)
    project = first.get("project_name") or cfg.get("project_name") or "USB Performance Tool"
    when = _span_text(ran or rows)
    sub = "Project: ______________________      Date: ______________" if template else f"{project}    ·    {when}"
    ws.merge_range(1, 1, 1, 12, sub, st(font_size=11, font_color="#FFFFFF", bg_color=BLUE, indent=1))
    ws.set_row(1, 22)

    roles = _ordered([r["dut_role"] for r in rows], ROLE_ORDER)
    devices = []
    for role in roles:
        pick = next((r for r in rows if r["dut_role"] == role and r["status"] != "SKIPPED" and r["dut_model"]),
                    next((r for r in rows if r["dut_role"] == role), {}))
        devices.append((role, pick))
    if template:
        dut_line = "Device under test: model, name and firmware — fill in the Devices Under Test table below."
    else:
        def device_text(role, p):
            name = p.get("dut_name")
            named = f' "{name}"' if name and name != "-" else ""
            return f"{role}: {p.get('dut_model') or 'Unknown model'}{named}  ·  FW {p.get('firmware_version') or 'Unknown'}"

        dut_line = "    |    ".join(device_text(role, p) for role, p in devices) or "Device details not available"
    ws.merge_range(2, 1, 2, 12, dut_line, st(italic=True, font_color="#FFFFFF", bg_color=BLUE, indent=1))
    ws.set_row(3, 8)

    status, dur, direction, thr = _rng("status"), _rng("duration_seconds"), _rng("direction"), _rng("throughput_mbps")
    rounds = len(ran)
    passed = sum(1 for r in ran if r["status"] == "PASS")
    skipped = len(rows) - len(ran)
    failed = rounds - passed
    ws.set_row(5, 36)
    ws.set_row(8, 34)
    # Tile value cells: Rounds=D6, Passed=F6, Failed=H6, Skipped=J6 (the verdict formula points at these).
    _tile(ws, st, 4, 5, 1, 2, "OVERALL RESULT", '=IF(D6+J6=0,"NO DATA",IF(H6>0,"FAIL",IF(J6>0,"INCOMPLETE","PASS")))',
          _verdict(rounds, failed, skipped), "General", NAVY, NAVY)
    for value, bg, fg in (("PASS", "#C6EFCE", "#006100"), ("FAIL", "#FFC7CE", "#9C0006"), ("INCOMPLETE", "#FFEB9C", "#7F6000")):
        ws.conditional_format(5, 1, 5, 2, {"type": "cell", "criteria": "==", "value": f'"{value}"',
                                           "format": st(bg_color=bg, font_color=fg)})
    _tile(ws, st, 4, 5, 3, 4, "ROUNDS RUN", f'=COUNTA({status})-COUNTIF({status},"SKIPPED")', rounds, "0", NAVY, NAVY)
    _tile(ws, st, 4, 5, 5, 6, "PASSED", f'=COUNTIF({status},"PASS")', passed, "0", "#548235", GREEN)
    _tile(ws, st, 4, 5, 7, 8, "FAILED", "=D6-F6", failed, "0", "#C00000", RED)
    _tile(ws, st, 4, 5, 9, 10, "SKIPPED", f'=COUNTIF({status},"SKIPPED")', skipped, "0", "#BF8F00", AMBER)
    _tile(ws, st, 4, 5, 11, 12, "PASS RATE", f'=IFERROR(F6/D6,"{NA}")', passed / rounds if rounds else NA, "0%", NAVY, NAVY)
    ws.set_row(6, 6)
    mbps = '0.0 "MB/s"'
    _tile(ws, st, 7, 8, 1, 3, "AVG READ", f'=IFERROR(AVERAGEIFS({thr},{direction},"read",{dur},">0"),"{NA}")',
          _avg(ran, "read"), mbps, BLUE, BLUE)
    _tile(ws, st, 7, 8, 4, 6, "AVG WRITE", f'=IFERROR(AVERAGEIFS({thr},{direction},"write",{dur},">0"),"{NA}")',
          _avg(ran, "write"), mbps, ORANGE, ORANGE)
    for c1, c2, label, d, color in ((7, 9, "PEAK READ", "read", BLUE), (10, 12, "PEAK WRITE", "write", ORANGE)):
        _tile(ws, st, 7, 8, c1, c2, label,
              f'=IF(COUNTIFS({direction},"{d}")=0,"{NA}",SUMPRODUCT(MAX(({direction}="{d}")*{thr})))',
              _peak(ran, d), mbps, color, color)
    ws.set_row(9, 10)

    def ref(name, col):
        a = analysis[name]
        return ["Analysis", a["first"], col, a["last"], col]

    def speed_chart(name, title, label_col=0):
        n = analysis[name]["last"] - analysis[name]["first"] + 1
        return _column_chart(wb, title, ref(name, label_col), ref(name, 6), ref(name, 7), n, analysis[name]["peak"])

    charts = [
        speed_chart("interface", "Average Throughput by Interface"),
        speed_chart("format", "Average Throughput by USB Format"),
        _outcome_chart(wb, ref("outcome", 0), ref("outcome", 1)),
        speed_chart("file", "Average Throughput by File Size", label_col=8),
    ]
    if template or len(analysis["role"]["labels"]) > 1:
        charts.append(speed_chart("role", "Router vs Extender"))
    if template or len(analysis["connection"]["labels"]) > 1:
        charts.append(speed_chart("connection", "Wired vs Wireless"))
    rounds_info = analysis["rounds"]
    if template or rounds_info["count"] > 1:  # with a single round this would just repeat the overall average
        r1, r2 = rounds_info["first"], rounds_info["last"]
        charts.append(_column_chart(wb, "Average Throughput per Round", ["Analysis", r1, 0, r2, 0],
                                    ["Analysis", r1, 2, r2, 2], ["Analysis", r1, 3, r2, 3],
                                    rounds_info["count"], rounds_info["peak"]))
    chart_top = 10
    for i, chart in enumerate(charts):
        ws.insert_chart(chart_top + (i // 2) * CHART_PITCH, 1 if i % 2 == 0 else 7, chart, {"x_offset": 2, "y_offset": 2})
    chart_rows = (len(charts) + 1) // 2
    row = chart_top + chart_rows * CHART_PITCH + 1
    # Two chart rows fit on a printed page; break there (and before the details when charts need a third page)
    # so no chart is sliced across two pages.
    breaks = [chart_top + 2 * CHART_PITCH] if chart_rows >= 3 else []
    breaks.append(row if chart_rows < 3 or chart_rows >= 4 else None)
    ws.set_h_pagebreaks([b for b in breaks if b])

    sect = st(bold=True, font_size=11, font_color=NAVY, bottom=2, bottom_color=NAVY)
    key_fmt = st(bold=True, font_color=GRAY_TXT, bg_color=PANEL, indent=1)
    in_fmt = st(bg_color=INPUT_FILL, border=1, border_color="#E6D9A8", indent=1)
    value_fmt = in_fmt if template else st(bottom=4, bottom_color="#D9D9D9", indent=1)

    def kv(r, c, key, value, val_cols):
        ws.merge_range(r, c, r, c + 1, key, key_fmt)
        ws.merge_range(r, c + 2, r, c + 1 + val_cols, "" if template else (value or ""), value_fmt)

    ws.merge_range(row, 1, row, 12, "TEST ENVIRONMENT", sect)
    row += 1
    measured = _ordered([r["configured_interface"] for r in ran], INTERFACE_ORDER)
    unmeasured = [i for i in _ordered([r["configured_interface"] for r in rows if r["status"] == "SKIPPED"], INTERFACE_ORDER)
                  if i not in measured]
    tested = ", ".join(measured) + (f"   (not tested: {', '.join(unmeasured)})" if unmeasured else "")
    iterations = max([int(r["iteration"] or 0) for r in rows] + [0])
    left = [("Project", project), ("Test date", when), ("Topology", first.get("topology")),
            ("Connection", " + ".join(_unique(ran, "connection_type")))]
    right = [("Client", ", ".join(x for x in (first.get("client_type"), first.get("client_os")) if x)),
             ("Client IP", ", ".join(_unique(ran, "client_ip"))),
             ("Rounds per case", str(iterations) if iterations else ""), ("Interfaces tested", tested)]
    link = cfg.get("client_link") or {}
    link_text = f"{link.get('adapter', '')}  -  {link['mbps']:g} Mbps" if link.get("mbps") else ""
    left.append(("PC network link", link_text))
    right.append(("Log file", cfg.get("log_file") or ""))
    for i, ((lk, lv), (rk, rv)) in enumerate(zip(left, right)):
        kv(row + i, 1, lk, lv, 4)
        kv(row + i, 7, rk, rv, 4)
    row += len(left)
    # A report merged from several runs (e.g. Router then Extender) has more than one topology; the
    # row above shows only the first, so list them all across the full width.
    topologies = _unique(ran, "topology")
    if len(topologies) > 1:
        kv(row, 1, "All topologies", "   |   ".join(topologies), 10)
        row += 1
    kv(row, 1, "Test file set", cfg.get("file_set_summary") or "", 10)
    threshold_rows = _threshold_lines(cfg) or [("Pass thresholds", _threshold_text(cfg))]
    for i, (label, text) in enumerate(threshold_rows):
        kv(row + 1 + i, 1, label, text, 10)
    row += 2 + len(threshold_rows)
    # A slow or slower-than-the-port PC link changes how the speeds should be read.
    if not template:
        for text, fill, edge, color in ((cfg.get("client_link_warning"), "#FFF4E5", "#E08A00", "#6B4200"),
                                        (cfg.get("client_link_note"), "#EEF1F8", "#1F3F73", "#0B1F3A")):
            if text:
                ws.merge_range(row, 1, row, 12, text, st(text_wrap=True, valign="vcenter", bg_color=fill, font_color=color,
                                                         border=1, border_color=edge, indent=1, font_size=10))
                ws.set_row(row, 34)
                row += 1
        row += 1

    ws.merge_range(row, 1, row, 12, "DEVICES UNDER TEST", sect)
    row += 1
    head = st(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", border=1, border_color="#FFFFFF")
    spans = ((1, 1, "Role"), (2, 3, "Name"), (4, 5, "Model"), (6, 9, "Firmware"), (10, 10, "HW Version"), (11, 12, "USB Share"))
    for c1, c2, label in spans:
        ws.write(row, c1, label, head) if c1 == c2 else ws.merge_range(row, c1, row, c2, label, head)
    row += 1
    cell = in_fmt if template else st(border=1, border_color="#D9D9D9", indent=1)
    role_fmt = in_fmt if template else st(border=1, border_color="#D9D9D9", align="center", bold=True)
    for role, p in ([("Router", {}), ("Extender", {})] if template else devices):
        vals = [None] * 5 if template else [p.get("dut_name"), p.get("dut_model"), p.get("firmware_version"),
                                            p.get("detected_hardware_version"), p.get("usb_share_path")]
        ws.write(row, 1, role, role_fmt)
        for (c1, c2, _label), v in zip(spans[1:], vals):
            ws.write(row, c1, v or "", cell) if c1 == c2 else ws.merge_range(row, c1, row, c2, v or "", cell)
        row += 1
    # A fixed print scale (not fit-to-page) so the manual page breaks above are honoured.
    ws.set_landscape()
    ws.set_paper(1)
    ws.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.set_print_scale(78)
    ws.set_zoom(100)


def _tile(ws, st, label_row, value_row, c1, c2, label, formula, value, num_format, color, accent):
    ws.merge_range(label_row, c1, label_row, c2, label,
                   st(font_size=9, bold=True, font_color=GRAY_TXT, bg_color=PANEL, align="center"))
    value_fmt = st(font_size=20, bold=True, font_color=color, bg_color=PANEL, align="center",
                   num_format=num_format, bottom=5, bottom_color=accent)
    ws.merge_range(value_row, c1, value_row, c2, "", value_fmt)
    ws.write_formula(value_row, c1, formula, value_fmt, value)


# ------------------------------------------------------------------- public API

def _build(path, rows, config, template):
    wb = xlsxwriter.Workbook(path)
    wb.set_properties({
        "title": "USB Performance Test Results",
        "subject": "USB read/write throughput through a router or mesh device",
        "author": "USB Performance Automation Tool",
        "comments": "Blank template" if template else "Generated after a test run",
    })
    _state["capacity"] = max(MIN_CAPACITY, len(rows) + 1000)
    st = _Styles(wb)
    sheets = {name: wb.add_worksheet(name) for name in SHEET_ORDER}  # tab order is creation order
    _build_results_sheet(sheets["Results"], st, rows)
    layout = _build_analysis_sheet(wb, sheets["Analysis"], st, rows, template)
    _build_issues_sheet(sheets["Issues"], st, rows, template)
    _build_notes_sheet(sheets["Notes"], st, template)
    _build_summary_sheet(wb, sheets["Summary"], st, rows, config, template, layout)
    sheets["Summary"].activate()
    wb.close()


def _normalise(results):
    rows = []
    for item in results or []:
        d = _as_dict(item)
        for key, *_rest in COLUMNS:
            d.setdefault(key, "")
        d.setdefault("status", "")
        rows.append(d)
    return rows


def _write(path, rows, config, template):
    try:
        _build(path, rows, config, template)
        return path
    except (OSError, xlsxwriter.exceptions.FileCreateError):
        # Most often the file is open in Excel. Write a timestamped copy rather than fail the run.
        stem, ext = os.path.splitext(path)
        alt = f"{stem}_{datetime.now():%Y%m%d_%H%M%S}{ext}"
        _build(alt, rows, config, template)
        return alt


def write_workbook(out_dir, results, config=None, filename=RESULTS_FILE):
    """Write the formatted results workbook; returns the path actually written."""
    os.makedirs(out_dir, exist_ok=True)
    return _write(os.path.join(out_dir, filename), _normalise(results), config, False)


def write_template(out_dir, filename=TEMPLATE_FILE, overwrite=False):
    """Write the blank template (same layout, no data). An existing file is kept unless overwrite=True."""
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, filename)
    if os.path.exists(path) and not overwrite:
        return path
    return _write(path, [], {}, True)
