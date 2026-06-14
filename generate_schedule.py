"""
Employee Weekly Shift Schedule Generator
Produces a formatted .xlsx file ready to open in Excel or Google Sheets.
"""

from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_NUMBER_00

# ─────────────────────────── Configuration ───────────────────────────────────

WEEK_START = date(2026, 6, 15)          # Monday of the schedule week

EMPLOYEES = [
    "Alice Johnson",
    "Bob Martinez",
    "Carol Smith",
    "David Lee",
    "Emma Wilson",
    "Frank Brown",
    "Grace Kim",
    "Henry Davis",
]

# Each shift: label, start time string, end time string, hours worked
SHIFTS = {
    "M":  ("Morning",   "06:00", "14:00", 8),
    "A":  ("Afternoon", "14:00", "22:00", 8),
    "N":  ("Night",     "22:00", "06:00", 8),
    "D":  ("Day",       "08:00", "16:00", 8),
    "S":  ("Split",     "10:00", "18:00", 7),
    "OFF":("Off",       "",      "",      0),
}

# Sample schedule — rows = employees, cols = Mon..Sun
# Edit freely; use shift codes above or leave "" for blank (to fill in manually)
SCHEDULE = [
    #         Mon   Tue   Wed   Thu   Fri   Sat   Sun
    ["Alice",  "M",  "M",  "M",  "M",  "M",  "OFF","OFF"],
    ["Bob",    "A",  "A",  "OFF","A",  "A",  "A",  "OFF"],
    ["Carol",  "D",  "OFF","D",  "D",  "D",  "D",  "OFF"],
    ["David",  "OFF","N",  "N",  "N",  "N",  "OFF","N"  ],
    ["Emma",   "S",  "S",  "S",  "OFF","S",  "S",  "OFF"],
    ["Frank",  "D",  "D",  "OFF","D",  "D",  "OFF","D"  ],
    ["Grace",  "M",  "OFF","M",  "M",  "OFF","M",  "M"  ],
    ["Henry",  "A",  "A",  "A",  "OFF","A",  "A",  "A"  ],
]

# ──────────────────────────── Color palette ───────────────────────────────────

CLR = {
    "header_bg":  "1F4E79",   # dark blue header
    "header_fg":  "FFFFFF",
    "title_bg":   "2E75B6",
    "subhdr_bg":  "D6E4F0",
    "subhdr_fg":  "1F4E79",
    "weekend_bg": "F2F2F2",
    "M":          "C6EFCE",   # morning — green
    "A":          "FFEB9C",   # afternoon — yellow
    "N":          "9DC3E6",   # night — blue
    "D":          "E2EFDA",   # day — light green
    "S":          "FCE4D6",   # split — peach
    "OFF":        "F2F2F2",   # off — grey
    "total_bg":   "DAEEF3",
    "legend_hdr": "2E75B6",
}

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# ─────────────────────────── Helpers ─────────────────────────────────────────

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="000000", size=11, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")

def align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def thin_border(top=True, bottom=True, left=True, right=True):
    s = Side(style="thin", color="BFBFBF")
    n = None
    return Border(
        top=s if top else n,
        bottom=s if bottom else n,
        left=s if left else n,
        right=s if right else n,
    )

def thick_border():
    s = Side(style="medium", color="1F4E79")
    return Border(top=s, bottom=s, left=s, right=s)

def date_range_str(start: date) -> str:
    end = start + timedelta(days=6)
    return f"{start.strftime('%B %d')} – {end.strftime('%B %d, %Y')}"

# ─────────────────────────── Builder ─────────────────────────────────────────

def build_schedule(wb: Workbook):
    ws = wb.active
    ws.title = "Weekly Schedule"

    # ── Column widths ──────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 20   # employee name
    for col in range(2, 10):               # Mon-Sun + Total
        ws.column_dimensions[get_column_letter(col)].width = 15

    # ── Row 1: title ──────────────────────────────────────────────────
    ws.row_dimensions[1].height = 36
    ws.merge_cells("A1:I1")
    c = ws["A1"]
    c.value = "EMPLOYEE WEEKLY SHIFT SCHEDULE"
    c.fill = fill(CLR["header_bg"])
    c.font = font(bold=True, color=CLR["header_fg"], size=16)
    c.alignment = align()

    # ── Row 2: week range ─────────────────────────────────────────────
    ws.row_dimensions[2].height = 22
    ws.merge_cells("A2:I2")
    c = ws["A2"]
    c.value = f"Week of  {date_range_str(WEEK_START)}"
    c.fill = fill(CLR["title_bg"])
    c.font = font(bold=False, color="FFFFFF", size=12, italic=True)
    c.alignment = align()

    # ── Row 3: blank spacer ───────────────────────────────────────────
    ws.row_dimensions[3].height = 8

    # ── Row 4: column headers ─────────────────────────────────────────
    ws.row_dimensions[4].height = 40
    headers = ["Employee"] + DAYS + ["Total Hrs"]
    for col_idx, label in enumerate(headers, start=1):
        c = ws.cell(row=4, column=col_idx, value=label)
        c.fill = fill(CLR["subhdr_bg"])
        c.font = font(bold=True, color=CLR["subhdr_fg"], size=11)
        c.alignment = align(wrap=True)
        c.border = thin_border()

    # ── Row 5: date sub-headers ───────────────────────────────────────
    ws.row_dimensions[5].height = 22
    ws.cell(row=5, column=1, value="").fill = fill(CLR["subhdr_bg"])
    for d in range(7):
        day_date = WEEK_START + timedelta(days=d)
        c = ws.cell(row=5, column=d + 2, value=day_date.strftime("%b %d"))
        c.fill = fill(CLR["subhdr_bg"])
        c.font = font(italic=True, color="555555", size=10)
        c.alignment = align()
        c.border = thin_border()
    ws.cell(row=5, column=9, value="").fill = fill(CLR["subhdr_bg"])

    # ── Employee rows ─────────────────────────────────────────────────
    for row_offset, row_data in enumerate(SCHEDULE):
        r = 6 + row_offset
        ws.row_dimensions[r].height = 28
        emp_name = EMPLOYEES[row_offset] if row_offset < len(EMPLOYEES) else row_data[0]

        # Employee name cell
        c = ws.cell(row=r, column=1, value=emp_name)
        c.font = font(bold=True, size=11)
        c.alignment = align(h="left")
        c.border = thin_border()
        if row_offset % 2 == 1:
            c.fill = fill("F9F9F9")

        total_hrs = 0
        for d in range(7):
            col = d + 2
            code = row_data[1 + d] if len(row_data) > 1 + d else "OFF"
            shift_info = SHIFTS.get(code, SHIFTS["OFF"])
            label, start, end, hrs = shift_info
            total_hrs += hrs

            display = f"{label}\n{start}–{end}" if start else label
            cell = ws.cell(row=r, column=col, value=display)

            bg = CLR.get(code, "FFFFFF")
            if d >= 5:  # weekend overlay
                cell.fill = fill("EDEDF5") if code == "OFF" else fill(bg)
            else:
                cell.fill = fill(bg)

            cell.font = font(size=10)
            cell.alignment = align(wrap=True)
            cell.border = thin_border()

        # Total hours
        tc = ws.cell(row=r, column=9, value=total_hrs)
        tc.fill = fill(CLR["total_bg"])
        tc.font = font(bold=True, size=11)
        tc.alignment = align()
        tc.border = thin_border()

    # ── Thick outer border around data block ─────────────────────────
    last_row = 5 + len(SCHEDULE)
    for r in range(4, last_row + 1):
        for col in range(1, 10):
            cell = ws.cell(row=r, column=col)
            top    = Side(style="medium", color="1F4E79") if r == 4 else None
            bottom = Side(style="medium", color="1F4E79") if r == last_row else None
            left   = Side(style="medium", color="1F4E79") if col == 1 else None
            right  = Side(style="medium", color="1F4E79") if col == 9 else None
            if any([top, bottom, left, right]):
                existing = cell.border
                cell.border = Border(
                    top=top or existing.top,
                    bottom=bottom or existing.bottom,
                    left=left or existing.left,
                    right=right or existing.right,
                )

    # ── Legend ────────────────────────────────────────────────────────
    legend_start_row = last_row + 3
    ws.row_dimensions[legend_start_row].height = 22
    ws.merge_cells(f"A{legend_start_row}:I{legend_start_row}")
    lc = ws[f"A{legend_start_row}"]
    lc.value = "SHIFT LEGEND"
    lc.fill = fill(CLR["legend_hdr"])
    lc.font = font(bold=True, color="FFFFFF", size=12)
    lc.alignment = align()

    for i, (code, (label, start, end, hrs)) in enumerate(SHIFTS.items()):
        lr = legend_start_row + 1 + i
        ws.row_dimensions[lr].height = 22
        time_str = f"{start} – {end}" if start else "—"
        hrs_str  = f"{hrs}h" if hrs else "—"

        code_cell = ws.cell(row=lr, column=1, value=f"  {code}")
        code_cell.fill = fill(CLR.get(code, "FFFFFF"))
        code_cell.font = font(bold=True, size=11)
        code_cell.alignment = align(h="left")
        code_cell.border = thin_border()

        label_cell = ws.cell(row=lr, column=2, value=label)
        label_cell.fill = fill(CLR.get(code, "FFFFFF"))
        label_cell.font = font(size=11)
        label_cell.alignment = align(h="left")
        label_cell.border = thin_border()

        time_cell = ws.cell(row=lr, column=3, value=time_str)
        time_cell.fill = fill(CLR.get(code, "FFFFFF"))
        time_cell.font = font(size=11)
        time_cell.alignment = align()
        time_cell.border = thin_border()

        hrs_cell = ws.cell(row=lr, column=4, value=hrs_str)
        hrs_cell.fill = fill(CLR.get(code, "FFFFFF"))
        hrs_cell.font = font(size=11)
        hrs_cell.alignment = align()
        hrs_cell.border = thin_border()

    # ── Freeze panes below headers ────────────────────────────────────
    ws.freeze_panes = "B6"

    # ── Print settings ────────────────────────────────────────────────
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.print_title_rows = "1:5"


def build_blank_template(wb: Workbook):
    """A second sheet with an empty, editable template."""
    ws = wb.create_sheet("Blank Template")

    ws.column_dimensions["A"].width = 20
    for col in range(2, 10):
        ws.column_dimensions[get_column_letter(col)].width = 15

    ws.row_dimensions[1].height = 36
    ws.merge_cells("A1:I1")
    c = ws["A1"]
    c.value = "EMPLOYEE WEEKLY SHIFT SCHEDULE"
    c.fill = fill(CLR["header_bg"])
    c.font = font(bold=True, color="FFFFFF", size=16)
    c.alignment = align()

    ws.row_dimensions[2].height = 22
    ws.merge_cells("A2:I2")
    c = ws["A2"]
    c.value = "Week of:  ___________________________"
    c.fill = fill(CLR["title_bg"])
    c.font = font(color="FFFFFF", size=12, italic=True)
    c.alignment = align()

    ws.row_dimensions[3].height = 8

    ws.row_dimensions[4].height = 40
    headers = ["Employee"] + DAYS + ["Total Hrs"]
    for col_idx, label in enumerate(headers, start=1):
        c = ws.cell(row=4, column=col_idx, value=label)
        c.fill = fill(CLR["subhdr_bg"])
        c.font = font(bold=True, color=CLR["subhdr_fg"], size=11)
        c.alignment = align()
        c.border = thin_border()

    ws.row_dimensions[5].height = 22
    for col in range(1, 10):
        c = ws.cell(row=5, column=col)
        c.fill = fill(CLR["subhdr_bg"])
        c.border = thin_border()
        if col > 1 and col < 9:
            c.font = font(italic=True, color="555555", size=10)
            c.alignment = align()
            c.value = "mm/dd"

    for row_offset in range(15):
        r = 6 + row_offset
        ws.row_dimensions[r].height = 28
        for col in range(1, 10):
            cell = ws.cell(row=r, column=col)
            cell.border = thin_border()
            cell.alignment = align()
            if col == 9:
                cell.fill = fill(CLR["total_bg"])
            elif row_offset % 2 == 1:
                cell.fill = fill("F9F9F9")
        ws.cell(row=r, column=1).alignment = align(h="left")

    ws.freeze_panes = "B6"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1


# ─────────────────────────── Visual timeline ─────────────────────────────────

# 30-minute slots from 8 AM to 9 PM
_TSLOTS = list(range(8 * 60, 21 * 60, 30))   # 26 slots

def _slot_color(slot_min, s_min, e_min, ls_min, le_min):
    """Return hex color for a timeline slot, or None if unworked."""
    if s_min is None or e_min is None:
        return None
    if slot_min < s_min or slot_min >= e_min:
        return None
    if ls_min is not None and ls_min <= slot_min < le_min:
        return "FFEB9C"   # lunch — yellow
    return "4472C4"       # on shift — blue

def _shift_times(code):
    """Convert a shift code to (start_min, end_min, lunch_start_min, lunch_end_min).
    Clamps to the 8 AM–9 PM timeline window."""
    info = SHIFTS.get(code, SHIFTS["OFF"])
    _, start_str, end_str, hrs = info
    if not start_str or not end_str:
        return None, None, None, None

    def t(s): h, m = map(int, s.split(":")); return h * 60 + m

    s_min = t(start_str)
    e_min = t(end_str)
    if e_min <= s_min:       # overnight shift — push end to next day
        e_min += 24 * 60

    TL_S, TL_E = 8 * 60, 21 * 60
    s_min = max(s_min, TL_S)
    e_min = min(e_min, TL_E)
    if s_min >= e_min:
        return None, None, None, None

    # Estimate 30-min lunch at midpoint for long shifts
    if hrs >= 6:
        mid = ((s_min + e_min) // 2) // 30 * 30   # snap to slot boundary
        ls_min, le_min = mid, mid + 30
    else:
        ls_min = le_min = None

    return s_min, e_min, ls_min, le_min


def build_visual_timeline(wb: Workbook):
    """Visual Gantt-style timeline: one row per employee per day (Mon–Fri).
    Columns = 30-min time slots 8 AM–9 PM.  Scroll vertically to compare employees."""
    ws = wb.create_sheet("Visual Timeline")

    N_SLOTS  = len(_TSLOTS)    # 26
    SLOT_COL = 3               # time slot columns start at C

    ws.column_dimensions["A"].width = 20   # employee name
    ws.column_dimensions["B"].width = 12   # day label
    for ci in range(SLOT_COL, SLOT_COL + N_SLOTS):
        ws.column_dimensions[get_column_letter(ci)].width = 2.6

    # Row 1: title
    last_col = get_column_letter(SLOT_COL + N_SLOTS - 1)
    ws.merge_cells(f"A1:{last_col}1")
    ws.row_dimensions[1].height = 26
    c = ws["A1"]
    c.value = (f"Staff Coverage Timeline  —  Week of {WEEK_START.strftime('%B %d, %Y')}"
               "  (Mon–Fri  |  8 AM – 9 PM  |  30-min slots)")
    c.fill      = PatternFill("solid", fgColor="1F4E79")
    c.font      = Font(bold=True, color="FFFFFF", size=12, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center")

    # Row 2: headers + time labels
    ws.row_dimensions[2].height = 24
    thin = lambda: Border(
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
    )
    for col, txt in [(1, "Employee"), (2, "Day")]:
        c = ws.cell(row=2, column=col, value=txt)
        c.fill      = PatternFill("solid", fgColor="D6E4F0")
        c.font      = Font(bold=True, color="1F4E79", size=10, name="Calibri")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = thin()

    for si, slot in enumerate(_TSLOTS):
        h, m = slot // 60, slot % 60
        label = str(h) if m == 0 else "30"
        c = ws.cell(row=2, column=SLOT_COL + si, value=label)
        c.fill      = PatternFill("solid", fgColor="EBF3FB")
        c.font      = Font(size=7, color="1F4E79", name="Calibri", bold=(m == 0))
        c.alignment = Alignment(horizontal="center", vertical="bottom", text_rotation=90)
        c.border    = Border(
            bottom=Side(style="thin", color="BFBFBF"),
            left=Side(style="hair", color="DDDDDD"),
        )

    DAYS_WD = ["Monday","Tuesday","Wednesday","Thursday","Friday"]
    DAYS_SH = ["Mon","Tue","Wed","Thu","Fri"]

    cur = 3
    for ei, emp_row in enumerate(SCHEDULE):
        # emp_row: ["Alice", "M", "A", "OFF", ...]  (index 0=name, 1-7=Mon-Sun)
        emp_name   = EMPLOYEES[ei] if ei < len(EMPLOYEES) else emp_row[0]
        emp_start  = cur
        row_bg     = "F2F6FA" if ei % 2 == 0 else "FAFAFA"
        name_bg    = "E8F0F8" if ei % 2 == 0 else "EBF3FB"

        for di in range(5):   # Mon=0, Fri=4
            ws.row_dimensions[cur].height = 16
            code      = emp_row[1 + di] if len(emp_row) > 1 + di else "OFF"
            day_date  = WEEK_START + __import__("datetime").timedelta(days=di)
            day_label = f"{DAYS_SH[di]}  {day_date.strftime('%m/%d')}"
            is_last   = (di == 4)

            c = ws.cell(row=cur, column=2, value=day_label)
            c.font      = Font(size=9, bold=(di == 0), name="Calibri")
            c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            c.fill      = PatternFill("solid", fgColor=row_bg)
            c.border    = Border(
                bottom=Side(style="medium" if is_last else "hair",
                            color="BFBFBF" if is_last else "E0E0E0"),
                left=Side(style="thin", color="BFBFBF"),
                right=Side(style="thin", color="BFBFBF"),
            )

            s_min, e_min, ls_min, le_min = _shift_times(code)

            for si, slot in enumerate(_TSLOTS):
                col   = SLOT_COL + si
                color = _slot_color(slot, s_min, e_min, ls_min, le_min)
                cell  = ws.cell(row=cur, column=col)
                cell.fill   = PatternFill("solid", fgColor=color if color else row_bg)
                cell.border = Border(
                    bottom=Side(style="medium" if is_last else "hair",
                                color="BFBFBF" if is_last else "E0E0E0"),
                    left=Side(style="hair", color="D8D8D8"),
                )
            cur += 1

        # Merge employee name across 5 day rows
        ws.merge_cells(f"A{emp_start}:A{cur - 1}")
        c = ws.cell(row=emp_start, column=1, value=emp_name)
        c.font      = Font(bold=True, size=10, name="Calibri")
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c.fill      = PatternFill("solid", fgColor=name_bg)
        c.border    = Border(
            top=Side(style="medium", color="1F4E79"),
            bottom=Side(style="medium", color="1F4E79"),
            left=Side(style="medium", color="1F4E79"),
            right=Side(style="thin", color="BFBFBF"),
        )

        # Thin separator between employees
        ws.row_dimensions[cur].height = 4
        cur += 1

    # Legend
    leg = cur + 1
    ws.row_dimensions[leg].height = 16
    ws.cell(row=leg, column=1, value="Legend:").font = Font(bold=True, size=9, name="Calibri")
    for i, (color, lbl) in enumerate([
            ("4472C4", "On Shift"),
            ("FFEB9C", "Lunch (estimated midpoint)"),
            ("F2F6FA", "Off / Not Working")]):
        c = ws.cell(row=leg, column=3 + i * 3)
        c.fill   = PatternFill("solid", fgColor=color)
        c.border = thin()
        c = ws.cell(row=leg, column=4 + i * 3, value=f"  {lbl}")
        c.font      = Font(size=9, name="Calibri")
        c.alignment = Alignment(horizontal="left", vertical="center")

    ws.freeze_panes = "C3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1


# ─────────────────────────── Main ────────────────────────────────────────────

def main():
    wb = Workbook()
    build_schedule(wb)
    build_blank_template(wb)
    build_visual_timeline(wb)

    out_path = "employee_weekly_schedule.xlsx"
    wb.save(out_path)
    print(f"Saved → {out_path}")


if __name__ == "__main__":
    main()
