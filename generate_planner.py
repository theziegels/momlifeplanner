"""
Ultimate Mom Life Planner – v3
Aug 2026 – Dec 2027  (17 planner months)
Year-at-a-Glance: Jul 2026 – Feb 2028 = 20 months, 5 × 4 grid
All pages: landscape letter, fit to one page
"""

import calendar
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Brand Colors ──────────────────────────────────────────────────────────────
CREAM        = "FAF7F4"
WARM_LIGHT   = "F0EAE4"
WARM_MED     = "E2D5CC"
WARM_DARK    = "C4B5A8"
ACCENT       = "CE8282"
ACCENT_LIGHT = "EDD5D5"
TEXT_DARK    = "3E342A"
TEXT_MED     = "7A6B62"
TEXT_LIGHT   = "A89890"

# ── Fonts ─────────────────────────────────────────────────────────────────────
def tf(size=11, bold=False, color=TEXT_DARK, italic=False):
    return Font(name="Cormorant Garamond", size=size, bold=bold,
                color=color, italic=italic)

def mf(size=9, bold=False, color=TEXT_DARK, italic=False):
    return Font(name="Montserrat", size=size, bold=bold, color=color, italic=italic)

# ── Helpers ───────────────────────────────────────────────────────────────────
def fill(c):    return PatternFill("solid", fgColor=c)
def sd(style="thin", color=WARM_DARK): return Side(style=style, color=color)

def bdr(left=None, right=None, top=None, bottom=None):
    return Border(left=left or Side(style=None), right=right or Side(style=None),
                  top=top or Side(style=None),   bottom=bottom or Side(style=None))

def al(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

C = al("center"); L = al("left")

def bg(ws, r1, c1, r2, c2, color):
    f = fill(color)
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row: cell.fill = f

def bdr_range(ws, r1, c1, r2, c2, b):
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row: cell.border = b

def mc(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    return ws.cell(r1, c1)

def ruled(ws, r1, r2, c1, c2, color=WARM_MED, style="hair"):
    b = bdr(bottom=sd(style, color))
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            ws.cell(r, c).border = b

def landscape_all(ws, scale=80):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 1
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale     = scale
    ws.page_margins.left   = 0.4
    ws.page_margins.right  = 0.4
    ws.page_margins.top    = 0.35
    ws.page_margins.bottom = 0.35

# ── Calendar data ─────────────────────────────────────────────────────────────
MONTHS_PLANNER = (
    [(2026, m) for m in range(8, 13)] +
    [(2027, m) for m in range(1, 13)]
)

MN = ["","January","February","March","April","May","June",
      "July","August","September","October","November","December"]
MA = ["","Jan","Feb","Mar","Apr","May","Jun",
      "Jul","Aug","Sep","Oct","Nov","Dec"]
DF = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

def weeks_of_month(year, month):
    first = date(year, month, 1)
    last  = date(year, month, calendar.monthrange(year, month)[1])
    start = first - timedelta(days=first.weekday())
    weeks, d = [], start
    while d <= last:
        weeks.append([d + timedelta(days=i) for i in range(7)])
        d += timedelta(weeks=1)
    return weeks


# ══════════════════════════════════════════════════════════════════════════════
# COVER
# ══════════════════════════════════════════════════════════════════════════════
def make_cover(wb):
    ws = wb.create_sheet("Cover")
    landscape_all(ws)
    for c in range(1, 17): ws.column_dimensions[get_column_letter(c)].width = 7
    for r in range(1, 48):  ws.row_dimensions[r].height = 14

    bg(ws, 1, 1, 48, 16, CREAM)
    bg(ws, 1, 1, 3,  16, ACCENT_LIGHT)
    bg(ws, 1, 1, 48, 1,  WARM_MED)

    c = mc(ws, 8,  2, 8,  16); c.value = "THE ULTIMATE"
    c.font = tf(14, italic=True, color=TEXT_MED); c.alignment = C

    c = mc(ws, 11, 2, 13, 16); c.value = "Mom Life"
    c.font = tf(48, bold=True, color=TEXT_DARK); c.alignment = C

    c = mc(ws, 14, 2, 14, 16); c.value = "P  L  A  N  N  E  R"
    c.font = mf(12, color=ACCENT); c.alignment = C

    for col in range(3, 16):
        ws.cell(16, col).border = bdr(bottom=sd("medium", ACCENT))

    c = mc(ws, 18, 2, 18, 16); c.value = "August 2026 – December 2027"
    c.font = mf(10, italic=True, color=TEXT_MED); c.alignment = C

    for box in [(21,4,23,6),(21,8,23,10),(21,12,23,14)]:
        bg(ws, *box, WARM_LIGHT)
        bdr_range(ws, *box, bdr(left=sd(),right=sd(),top=sd(),bottom=sd()))

    c = mc(ws, 31, 2, 32, 16)
    c.value = '"You are doing better than you think."'
    c.font = tf(14, italic=True, color=TEXT_MED); c.alignment = C

    c = mc(ws, 43, 2, 43, 16); c.value = "plan  ·  dream  ·  thrive"
    c.font = mf(9, italic=True, color=TEXT_LIGHT); c.alignment = C

    bg(ws, 46, 1, 48, 16, WARM_LIGHT)
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# HOW TO USE
# ══════════════════════════════════════════════════════════════════════════════
def make_how_to_use(wb):
    ws = wb.create_sheet("How To Use")
    landscape_all(ws)
    for c in range(1, 21): ws.column_dimensions[get_column_letter(c)].width = 6.5
    for r in range(1, 50):  ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 50, 20, CREAM)
    bg(ws, 1, 1, 4,  20, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, 20)
    c.value = "How to Use This Planner"; c.font = tf(22, bold=True); c.alignment = C

    items = [
        (6,  "Year at a Glance",
             "20 months on one landscape page (July 2026 – February 2028). "
             "Mark holidays, vacations, and big life moments at a glance."),
        (11, "Monthly Layout",
             "One landscape page per month. Full calendar grid with generous "
             "writing space per day, plus a notes column on the right."),
        (16, "Weekly Layout",
             "One landscape page per week. Left: Mon–Sun in a 2-column grid "
             "with ruled writing lines. Right: Priorities, To-Do list, Notes."),
        (21, "Goals Pages",
             "Separate pages for Family, Personal, and Professional goals — "
             "big goals, quarterly breakdown, and action steps."),
        (26, "Brain Map & Brainstorm",
             "Brain Map: central idea branching outward. "
             "Brainstorm: spine-and-branch layout for free-flow thinking."),
        (31, "Vision Board + Extras",
             "Vision Board, Bucket List, Self-Care Tracker, and Important Contacts."),
    ]
    for row, title, body in items:
        bg(ws, row, 2, row, 2, ACCENT); ws.row_dimensions[row].height = 18
        c = mc(ws, row, 3, row, 20)
        c.value = title; c.font = tf(13, bold=True, color=ACCENT); c.alignment = L
        c = mc(ws, row+1, 3, row+3, 20)
        c.value = body; c.font = mf(9, color=TEXT_MED); c.alignment = al("left","top")

    c = mc(ws, 44, 2, 45, 19)
    c.value = "This planner is yours — write in it, doodle in it, make it beautiful and real."
    c.font = tf(11, italic=True, color=TEXT_MED); c.alignment = C
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# IMPORTANT CONTACTS
# ══════════════════════════════════════════════════════════════════════════════
def make_important_contacts(wb):
    ws = wb.create_sheet("Important Contacts")
    landscape_all(ws)
    for c in range(1, 21): ws.column_dimensions[get_column_letter(c)].width = 6.5
    for r in range(1, 50):  ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 50, 20, CREAM)
    bg(ws, 1, 1, 4,  20, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, 20)
    c.value = "Important Contacts"; c.font = tf(22, bold=True); c.alignment = C

    cats = ["Emergency","Pediatrician","Dentist","School / Teacher",
            "Babysitter / Childcare","Husband / Partner","Mom / In-Laws",
            "Plumber / Electrician","Pharmacy","Other"]
    hdrs = ["Name","Phone","Email","Notes"]
    hcols = [2, 7, 12, 16]

    r = 7; ws.row_dimensions[r].height = 16
    for lbl, hc in zip(hdrs, hcols):
        span = 4 if hc < 16 else 4
        c = mc(ws, r, hc, r, hc+span-1)
        c.value = lbl; c.font = mf(8, bold=True, color=ACCENT); c.alignment = L
        bg(ws, r, hc, r, hc+span-1, WARM_LIGHT)

    r = 9
    for cat in cats:
        ws.row_dimensions[r].height = 18
        ws.cell(r, 1).value = cat
        ws.cell(r, 1).font  = mf(8, bold=True, color=TEXT_MED)
        for hc, span in zip(hcols, [4,4,4,4]):
            cell = mc(ws, r, hc, r, hc+span-1)
            cell.border = bdr(bottom=sd("thin", WARM_DARK))
        r += 2
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# YEAR AT A GLANCE  –  5 × 4 = 20 months, landscape
# ══════════════════════════════════════════════════════════════════════════════
def make_year_view(wb):
    ws = wb.create_sheet("Year at a Glance")
    landscape_all(ws, scale=80)

    ym_all = (
        [(2026, m) for m in range(7, 13)] +
        [(2027, m) for m in range(1, 13)] +
        [(2028, m) for m in range(1, 3)]
    )  # 20 months

    planner_months = set(MONTHS_PLANNER)

    # ── Columns: 5 cals × 7 day-cols + 4 gap-cols = 39 ──────────────────────
    DAY_W, GAP_W = 3.9, 1.6
    # col_starts[i] = first col of i-th calendar (0-indexed)
    col_starts = []
    c = 1
    for i in range(5):
        col_starts.append(c)
        c += 7
        if i < 4: c += 1

    for ci in range(1, 40):
        is_gap = any(ci == col_starts[k]+7 for k in range(4))
        ws.column_dimensions[get_column_letter(ci)].width = GAP_W if is_gap else DAY_W

    # ── Rows: 3 header + 4 groups × (1 name + 1 hdr + 6 weeks + 1 gap) ──────
    ROWS_PER = 9
    for r in range(1, 42): ws.row_dimensions[r].height = 13
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 6
    for g in range(4):
        ws.row_dimensions[3 + (g+1)*ROWS_PER].height = 5  # gap between groups

    bg(ws, 1, 1, 41, 39, CREAM)
    bg(ws, 1, 1, 2,  39, WARM_LIGHT)

    c = mc(ws, 1, 1, 1, 39)
    c.value = "Year at a Glance  ·  2026 – 2027"
    c.font  = tf(20, bold=True, italic=True); c.alignment = al("center")

    c = mc(ws, 2, 1, 2, 39)
    c.value = "July 2026 – February 2028  ·  Planner months: August 2026 – December 2027"
    c.font  = mf(7, italic=True, color=TEXT_MED); c.alignment = al("center")

    for col in range(1, 40):
        ws.cell(3, col).border = bdr(bottom=sd("medium", ACCENT))

    DAY_LTRS = ["S","M","T","W","T","F","S"]

    for idx, (yr, mo) in enumerate(ym_all):
        gcol = idx % 5
        grow = idx // 5
        cal_c = col_starts[gcol]
        cal_r = 4 + grow * ROWS_PER

        is_planner = (yr, mo) in planner_months
        name_color = ACCENT if is_planner else TEXT_LIGHT
        num_color  = TEXT_DARK if is_planner else TEXT_LIGHT
        hdr_bg     = ACCENT_LIGHT if is_planner else WARM_LIGHT

        # Month name
        ws.row_dimensions[cal_r].height = 14
        c = mc(ws, cal_r, cal_c, cal_r, cal_c+6)
        label = MN[mo] if yr == 2027 else f"{MN[mo]} '{str(yr)[2:]}"
        c.value = label; c.font = tf(8, bold=True, color=name_color)
        c.alignment = al("center")
        bg(ws, cal_r, cal_c, cal_r, cal_c+6, hdr_bg)

        # Day-letter headers
        ws.row_dimensions[cal_r+1].height = 10
        for d, ltr in enumerate(DAY_LTRS):
            cell = ws.cell(cal_r+1, cal_c+d)
            cell.value = ltr
            cell.font  = mf(6, bold=True,
                            color=ACCENT if d in (0,6) and is_planner else TEXT_LIGHT)
            cell.alignment = al("center")

        # Dates (Sun-first)
        for wi, week in enumerate(calendar.monthcalendar(yr, mo)):
            sun = week[6]; row_days = [sun] + week[:6]
            ws.row_dimensions[cal_r+2+wi].height = 12
            for d, day_num in enumerate(row_days):
                cell = ws.cell(cal_r+2+wi, cal_c+d)
                if day_num:
                    cell.value = day_num
                    cell.font  = mf(6, color=ACCENT if d==0 and is_planner else num_color)
                cell.alignment = al("center")

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# MONTHLY LAYOUT  –  landscape, one page per month
# ══════════════════════════════════════════════════════════════════════════════
def make_monthly_layout(wb, year, month):
    ws = wb.create_sheet(f"{MA[month]} {year}")
    landscape_all(ws, scale=80)

    # 7 day-cols + divider + 2 notes cols
    for d in range(1, 8):  ws.column_dimensions[get_column_letter(d)].width = 12
    ws.column_dimensions["H"].width = 0.8
    ws.column_dimensions["I"].width = 15
    ws.column_dimensions["J"].width = 15

    # Row heights: header rows 1-4, day-name row 5, then week blocks
    LINES = 5; BLOCK = 1 + LINES  # date-row + writing rows = 6

    for r in range(1, 50): ws.row_dimensions[r].height = 12
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 14
    ws.row_dimensions[4].height = 8
    ws.row_dimensions[5].height = 14

    bg(ws, 1, 1, 50, 10, CREAM)
    bg(ws, 1, 1, 4,  10, WARM_LIGHT)
    bg(ws, 1, 8, 50,  8, WARM_MED)

    # Header
    c = mc(ws, 1, 1, 2, 7)
    c.value = MN[month]; c.font = tf(26, bold=True, italic=True); c.alignment = L
    c = mc(ws, 3, 1, 3, 7)
    c.value = str(year); c.font = mf(10, color=TEXT_MED); c.alignment = L
    for col in range(1, 8):
        ws.cell(4, col).border = bdr(bottom=sd("medium", ACCENT))

    # Day-of-week headers
    day_names = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
    for d, dn in enumerate(day_names):
        cell = ws.cell(5, d+1)
        cell.value = dn
        cell.font  = mf(8, bold=True, color=ACCENT if d==0 else TEXT_MED)
        cell.alignment = al("center")
        bg(ws, 5, d+1, 5, d+1, WARM_LIGHT)

    # Calendar grid
    row = 6
    for week in calendar.monthcalendar(year, month):
        sun = week[6]; week_days = [sun] + week[:6]
        ws.row_dimensions[row].height = 14
        for d, day_num in enumerate(week_days):
            col = d+1
            if day_num:
                ws.cell(row, col).value = day_num
                ws.cell(row, col).font  = mf(9, bold=True,
                    color=ACCENT if d==0 else TEXT_DARK)
                ws.cell(row, col).alignment = al("left","top")
                ws.cell(row, col).border = bdr(top=sd("thin",WARM_MED),
                                               left=sd("hair",WARM_MED))
                for ln in range(1, LINES+1):
                    ws.row_dimensions[row+ln].height = 12
                    ws.cell(row+ln, col).border = bdr(bottom=sd("hair",WARM_MED),
                                                       left=sd("hair",WARM_MED))
            else:
                for lr in range(row, row+BLOCK):
                    bg(ws, lr, col, lr, col, WARM_LIGHT)
        row += BLOCK

    # Notes sidebar
    c = mc(ws, 1, 9, 2, 10)
    c.value = "Monthly Notes"; c.font = tf(14, bold=True); c.alignment = L
    bg(ws, 1, 9, 4, 10, WARM_LIGHT)

    note_secs = [(6,"Appointments"),(14,"Birthdays & Occasions"),
                 (22,"Meal Ideas"),(30,"Don't Forget")]
    for sr, slbl in note_secs:
        c = mc(ws, sr, 9, sr, 10)
        c.value = slbl; c.font = mf(8, bold=True, color=ACCENT); c.alignment = L
        bg(ws, sr, 9, sr, 10, ACCENT_LIGHT)
        for lr in range(sr+1, sr+8):
            ws.row_dimensions[lr].height = 13
            for lc in range(9, 11):
                ws.cell(lr, lc).border = bdr(bottom=sd("hair", WARM_MED))

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# WEEKLY LAYOUT  –  landscape, one page per week
#
# Left (cols 1–13): header + 3 pairs (Mon/Tue, Wed/Thu, Fri/Sat) + Sun full-width
# Right (cols 15–26): Priorities box | To-Do (single col, lots of lines) | Notes
# ══════════════════════════════════════════════════════════════════════════════
def make_weekly_layout(wb, week_dates, month_name, year):
    start_d = week_dates[0]  # Monday
    end_d   = week_dates[6]  # Sunday
    ws = wb.create_sheet(f"Wk {start_d.strftime('%b %-d')}")
    landscape_all(ws, scale=80)

    # ── Columns ───────────────────────────────────────────────────────────────
    # Left section: cols 1-6 (left-day), col 7 (mid-gap), cols 8-13 (right-day)
    # Spine:        col 14
    # Right section: cols 15-26 (priorities + todo + notes)
    for c in list(range(1,7)) + list(range(8,14)):
        ws.column_dimensions[get_column_letter(c)].width = 10.8
    ws.column_dimensions["G"].width  = 1.2   # mid gap between day columns
    ws.column_dimensions["N"].width  = 1.2   # spine
    for c in range(15, 27):
        ws.column_dimensions[get_column_letter(c)].width = 8.0

    # ── Row heights ───────────────────────────────────────────────────────────
    # Rows 1-3: header   Rows 4-56: content
    # Day blocks: each pair gets 13 rows (1 hdr + 10 lines + 2 gap)
    # Sun: rows 43-56 = 14 rows (1 hdr + 12 lines + 1 gap)
    LINES      = 10   # writing lines per day in paired blocks
    SUN_LINES  = 12   # writing lines for Sunday full-width
    DAY_HDR_H  = 15
    LINE_H     = 13.0
    GAP_H      = 4

    # Set all content rows first
    for r in range(1, 58):
        ws.row_dimensions[r].height = LINE_H

    # Header
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 7    # accent underline row

    # Day pair row starts (after header = row 4)
    # Pair 0 (Mon/Tue): rows 4-16  (1 hdr + 10 lines + 2 gap = 13)
    # Pair 1 (Wed/Thu): rows 17-29
    # Pair 2 (Fri/Sat): rows 30-42
    # Sun:              rows 43-56
    PAIR_BLOCK = 1 + LINES + 2  # 13 rows per pair
    SUN_BLOCK  = 1 + SUN_LINES + 1  # 14 rows

    pair_starts = [4, 4 + PAIR_BLOCK, 4 + PAIR_BLOCK*2]  # 4, 17, 30
    sun_start   = 4 + PAIR_BLOCK*3                         # 43

    for rs in pair_starts:
        ws.row_dimensions[rs].height = DAY_HDR_H
        for ln in range(1, LINES+1):
            ws.row_dimensions[rs+ln].height = LINE_H
        ws.row_dimensions[rs+LINES+1].height = GAP_H
        ws.row_dimensions[rs+LINES+2].height = GAP_H

    ws.row_dimensions[sun_start].height = DAY_HDR_H
    for ln in range(1, SUN_LINES+1):
        ws.row_dimensions[sun_start+ln].height = LINE_H
    ws.row_dimensions[sun_start+SUN_LINES+1].height = GAP_H

    # ── Backgrounds ───────────────────────────────────────────────────────────
    bg(ws, 1, 1,  57, 13, CREAM)
    bg(ws, 1, 14, 57, 14, WARM_MED)    # spine strip
    bg(ws, 1, 15, 57, 26, CREAM)

    # ── LEFT: Header (rows 1-2, then accent line row 3) ───────────────────────
    # Draw bg BEFORE any merges so fill lands on the right cells
    bg(ws, 1, 1, 2, 13, WARM_LIGHT)

    # Write header text AFTER bg
    c = mc(ws, 1, 1, 2, 8)
    c.value = f"{month_name.lower()}  {year}"
    c.font  = tf(18, bold=True, italic=True, color=TEXT_DARK)
    c.alignment = al("left", "center")

    c = mc(ws, 1, 9, 2, 13)
    c.value = f"{start_d.strftime('%-m/%-d')} – {end_d.strftime('%-m/%-d')}"
    c.font  = mf(8, italic=True, color=TEXT_MED)
    c.alignment = al("right", "center")

    for col in range(1, 14):
        ws.cell(3, col).border = bdr(bottom=sd("medium", ACCENT))

    # ── LEFT: Draw each day pair then Sunday ──────────────────────────────────
    def draw_day_block(day_idx, rs, c_start, c_end):
        """Draw one day: header row at rs, LINES ruled rows below."""
        d      = week_dates[day_idx]
        is_wkd = d.weekday() >= 5   # Sat or Sun

        ws.cell(rs, c_start).value = d.day
        ws.cell(rs, c_start).font  = mf(9, bold=True,
                                         color=ACCENT if is_wkd else TEXT_DARK)
        ws.cell(rs, c_start).alignment = al("left", "center")

        # Day name (merged across remaining cols)
        cell = mc(ws, rs, c_start+1, rs, c_end)
        cell.value = DF[day_idx].upper()
        cell.font  = mf(7, color=TEXT_MED)
        cell.alignment = al("left", "center")

        # Thin underline under header
        for col in range(c_start, c_end+1):
            ws.cell(rs, col).border = bdr(bottom=sd("thin", WARM_MED))

        # Ruled writing lines
        n_lines = SUN_LINES if day_idx == 6 else LINES
        for ln in range(1, n_lines+1):
            for col in range(c_start, c_end+1):
                ws.cell(rs+ln, col).border = bdr(bottom=sd("hair", WARM_MED))

    pairs = [(0,1), (2,3), (4,5)]
    for (li, ri), rs in zip(pairs, pair_starts):
        draw_day_block(li, rs, 1,  6)
        draw_day_block(ri, rs, 8, 13)

    # Sunday full width
    draw_day_block(6, sun_start, 1, 13)

    # ── RIGHT: Header (rows 1-2) ──────────────────────────────────────────────
    bg(ws, 1, 15, 2, 26, WARM_LIGHT)

    # ── RIGHT: PRIORITIES (rows 4-16) ─────────────────────────────────────────
    # Header label (row 4)
    ws.row_dimensions[4].height = 14
    c = mc(ws, 4, 15, 4, 26)
    c.value = "PRIORITIES"; c.font = mf(7, bold=True, color=ACCENT)
    c.alignment = al("left", "center")
    bg(ws, 4, 15, 4, 26, ACCENT_LIGHT)

    # Priorities box (rows 5-15) with outer border + inner ruled lines
    bdr_range(ws, 5, 15, 15, 26,
              bdr(left=sd("thin",WARM_DARK), right=sd("thin",WARM_DARK),
                  top=sd("thin",WARM_DARK),  bottom=sd("thin",WARM_DARK)))
    for r in range(6, 16):
        ws.row_dimensions[r].height = 14
        ruled(ws, r, r, 15, 26, WARM_MED, "hair")

    # ── RIGHT: TO-DO (rows 18 onward) ─────────────────────────────────────────
    TODO_START = 18
    ws.row_dimensions[TODO_START].height = 14
    c = mc(ws, TODO_START, 15, TODO_START, 26)
    c.value = "TO-DO"; c.font = mf(7, bold=True, color=TEXT_MED)
    c.alignment = al("left", "center")
    bg(ws, TODO_START, 15, TODO_START, 26, WARM_LIGHT)

    # 22 to-do lines (rows 19-40)
    TODO_LINES = 22
    for ti in range(TODO_LINES):
        tr = TODO_START + 1 + ti
        ws.row_dimensions[tr].height = 14

        # Circle bullet
        ws.cell(tr, 15).value     = "○"
        ws.cell(tr, 15).font      = Font(name="Montserrat", size=8, color=ACCENT)
        ws.cell(tr, 15).alignment = al("center")

        # Ruled line across cols 16-26
        for lc in range(16, 27):
            ws.cell(tr, lc).border = bdr(bottom=sd("hair", WARM_MED))

    # ── RIGHT: NOTES (rows 42 onward) ─────────────────────────────────────────
    NOTES_START = TODO_START + TODO_LINES + 2
    ws.row_dimensions[NOTES_START].height = 14
    c = mc(ws, NOTES_START, 15, NOTES_START, 26)
    c.value = "NOTES"; c.font = mf(7, bold=True, color=ACCENT)
    c.alignment = al("left", "center")
    bg(ws, NOTES_START, 15, NOTES_START, 26, ACCENT_LIGHT)

    # Notes ruled lines through end of page (rows NOTES_START+1 to 57)
    for nr in range(NOTES_START+1, 58):
        ws.row_dimensions[nr].height = 14
        ruled(ws, nr, nr, 15, 26, WARM_MED, "hair")

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# GOALS PAGES
# ══════════════════════════════════════════════════════════════════════════════
GOAL_TITLES = {
    "Family":       "My Big Goals for Our Family",
    "Personal":     "My Big Personal Goals",
    "Professional": "My Big Professional Goals",
}

def make_goals_page(wb, goal_type):
    ws = wb.create_sheet(f"{goal_type} Goals")
    landscape_all(ws)
    for c in range(1, 22): ws.column_dimensions[get_column_letter(c)].width = 6
    for r in range(1, 50):  ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 50, 21, CREAM)
    bg(ws, 1, 1, 4,  21, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, 21)
    c.value = f"{goal_type} Goals"; c.font = tf(24, bold=True); c.alignment = C

    # Section 1: Big Goals (goal_type specific title)
    S1 = 6
    c = mc(ws, S1, 2, S1, 21)
    c.value = GOAL_TITLES[goal_type]
    c.font = tf(13, bold=True, color=ACCENT); c.alignment = L
    bg(ws, S1, 2, S1, 21, ACCENT_LIGHT)

    for i, lr in enumerate(range(S1+2, S1+13)):
        ws.cell(lr, 2).value = f"{i+1}."
        ws.cell(lr, 2).font  = mf(9, color=ACCENT)
        for lc in range(3, 22):
            ws.cell(lr, lc).border = bdr(bottom=sd("hair", WARM_MED))
        ws.row_dimensions[lr].height = 18

    # Section 2: Quarter by Quarter
    S2 = S1 + 16
    c = mc(ws, S2, 2, S2, 21)
    c.value = "Quarter by Quarter"
    c.font = tf(13, bold=True, color=ACCENT); c.alignment = L
    bg(ws, S2, 2, S2, 21, ACCENT_LIGHT)

    q_labels = ["Q1  Jan–Mar","Q2  Apr–Jun","Q3  Jul–Sep","Q4  Oct–Dec"]
    q_cols   = [2, 7, 12, 17]
    for q_lbl, qc in zip(q_labels, q_cols):
        c = mc(ws, S2+2, qc, S2+2, qc+3)
        c.value = q_lbl; c.font = mf(8, bold=True, color=TEXT_MED); c.alignment = L
        for lr in range(S2+3, S2+11):
            for lc in range(qc, qc+4):
                ws.cell(lr, lc).border = bdr(bottom=sd("hair", WARM_MED))
            ws.row_dimensions[lr].height = 16

    # Section 3: Action Steps
    S3 = S2 + 13
    c = mc(ws, S3, 2, S3, 21)
    c.value = "Action Steps & Milestones"
    c.font = tf(13, bold=True, color=ACCENT); c.alignment = L
    bg(ws, S3, 2, S3, 21, ACCENT_LIGHT)

    for lr in range(S3+2, min(S3+13, 50)):
        for lc in range(2, 22):
            ws.cell(lr, lc).border = bdr(bottom=sd("hair", WARM_MED))
        ws.row_dimensions[lr].height = 18

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# BRAIN MAP  (to be revisited for visual design)
# ══════════════════════════════════════════════════════════════════════════════
def make_brain_map(wb):
    ws = wb.create_sheet("Brain Map")
    landscape_all(ws)
    for c in range(1, 29): ws.column_dimensions[get_column_letter(c)].width = 4.5
    for r in range(1, 50):  ws.row_dimensions[r].height = 13
    bg(ws, 1, 1, 50, 28, CREAM)
    bg(ws, 1, 1, 3,  28, WARM_LIGHT)

    c = mc(ws, 2, 1, 2, 28)
    c.value = "Brain Map"; c.font = tf(22, bold=True); c.alignment = C
    c = mc(ws, 3, 1, 3, 28)
    c.value = "★  To be redesigned  ★"
    c.font = mf(8, italic=True, color=TEXT_LIGHT); c.alignment = C

    # Central circle
    cr, cc = 26, 13
    bg(ws, cr-2, cc-1, cr+2, cc+3, ACCENT_LIGHT)
    cell = mc(ws, cr-2, cc-1, cr+2, cc+3)
    cell.value = "Central\nIdea"; cell.font = tf(12, bold=True, color=ACCENT)
    cell.alignment = al("center")
    bdr_range(ws, cr-2, cc-1, cr+2, cc+3,
              bdr(left=sd("medium",ACCENT),right=sd("medium",ACCENT),
                  top=sd("medium",ACCENT), bottom=sd("medium",ACCENT)))

    branch_boxes = [
        (8,3,6),(8,12,15),(8,21,24),
        (26,2,5),(26,18,21),
        (40,3,6),(40,12,15),(40,21,24),
    ]
    for (br_r, bc1, bc2) in branch_boxes:
        bg(ws, br_r, bc1, br_r+3, bc2, WARM_LIGHT)
        cell = mc(ws, br_r, bc1, br_r+3, bc2)
        cell.value = ""; cell.alignment = al("center")
        bdr_range(ws, br_r, bc1, br_r+3, bc2, bdr(left=sd(),right=sd(),top=sd(),bottom=sd()))
        for sub_r in range(br_r, br_r+4):
            for sub_c in range(max(1,bc1-3), bc1):
                ws.cell(sub_r, sub_c).border = bdr(bottom=sd("hair",WARM_MED))
            for sub_c in range(bc2+1, min(28,bc2+4)):
                ws.cell(sub_r, sub_c).border = bdr(bottom=sd("hair",WARM_MED))
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# BRAINSTORM  (to be revisited for visual design)
# ══════════════════════════════════════════════════════════════════════════════
def make_brainstorm(wb):
    ws = wb.create_sheet("Brainstorm")
    landscape_all(ws)
    for c in range(1, 29): ws.column_dimensions[get_column_letter(c)].width = 4.5
    for r in range(1, 50):  ws.row_dimensions[r].height = 13
    bg(ws, 1, 1, 50, 28, CREAM)
    bg(ws, 1, 1, 3,  28, WARM_LIGHT)

    c = mc(ws, 2, 1, 2, 28)
    c.value = "Brainstorm"; c.font = tf(22, bold=True); c.alignment = C
    c = mc(ws, 3, 1, 3, 28)
    c.value = "★  To be redesigned  ★"
    c.font = mf(8, italic=True, color=TEXT_LIGHT); c.alignment = C

    spine_row = 26
    for col in range(2, 28):
        ws.cell(spine_row, col).border = bdr(bottom=sd("medium", TEXT_MED))

    bg(ws, spine_row-1, 1, spine_row+1, 3, ACCENT_LIGHT)
    cell = mc(ws, spine_row-1, 1, spine_row+1, 3)
    cell.value = "Main\nIdea"; cell.font = tf(10, bold=True, color=ACCENT)
    cell.alignment = al("center")
    bdr_range(ws, spine_row-1, 1, spine_row+1, 3,
              bdr(left=sd("medium",ACCENT),right=sd("medium",ACCENT),
                  top=sd("medium",ACCENT),bottom=sd("medium",ACCENT)))

    for bc in [5, 9, 13, 17, 21, 25]:
        for arm in range(1, 4):
            for sign in [-1, 1]:
                br = spine_row + sign * arm * 5
                if 5 <= br <= 48:
                    ws.cell(br, bc).border = bdr(bottom=sd("thin", WARM_DARK))
                    mc(ws, br, bc, br, bc+2).border = bdr(bottom=sd("thin",WARM_DARK))
                    for leaf in range(1, 3):
                        lc = bc - leaf
                        if lc >= 1:
                            ws.cell(br, lc).border = bdr(bottom=sd("hair",WARM_MED))
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# VISION BOARD
# ══════════════════════════════════════════════════════════════════════════════
def make_vision_board(wb):
    ws = wb.create_sheet("Vision Board")
    landscape_all(ws)
    for c in range(1, 23): ws.column_dimensions[get_column_letter(c)].width = 6.2
    for r in range(1, 50):  ws.row_dimensions[r].height = 13
    bg(ws, 1, 1, 50, 22, CREAM)
    bg(ws, 1, 1, 4,  22, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, 22)
    c.value = "Vision Board"; c.font = tf(26, bold=True); c.alignment = C
    c = mc(ws, 4, 1, 4, 22)
    c.value = "Paste images  ·  write words  ·  capture feelings  ·  dream boldly"
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    words = ["FAITH","FAMILY","HEALTH","JOY","GROWTH","PEACE",
             "ADVENTURE","HOME","PURPOSE","LOVE","ABUNDANCE","FREEDOM"]
    for i, word in enumerate(words):
        wr = 6 + (i//6)*3; wc = 1 + (i%6)*3 + 1
        cell = mc(ws, wr, wc, wr+1, wc+1)
        cell.value = word; cell.font = mf(8, bold=True, color=TEXT_LIGHT)
        cell.alignment = al("center")
        bg(ws, wr, wc, wr+1, wc+1, WARM_LIGHT)
        bdr_range(ws, wr, wc, wr+1, wc+1, bdr(left=sd(),right=sd(),top=sd(),bottom=sd()))

    boxes = [
        (14, 1, 30, 11, "paste an image or write your dream here"),
        (14,12, 30, 22, "paste an image or write your dream here"),
        (32, 1, 42, 11, "this year I will..."),
        (32,12, 42, 22, "I am grateful for..."),
        (44, 1, 49, 22, "my word(s) for this season:"),
    ]
    for r1,c1,r2,c2,hint in boxes:
        bg(ws, r1,c1,r2,c2, WARM_LIGHT)
        bdr_range(ws,r1,c1,r2,c2, bdr(left=sd("thin",WARM_MED),right=sd("thin",WARM_MED),
                                       top=sd("thin",WARM_MED),bottom=sd("thin",WARM_MED)))
        cell = mc(ws, r1,c1,r2,c2)
        cell.value = hint; cell.font = mf(8, italic=True, color=TEXT_LIGHT)
        cell.alignment = al("center")
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# BUCKET LIST
# ══════════════════════════════════════════════════════════════════════════════
def make_bucket_list(wb):
    ws = wb.create_sheet("Bucket List")
    landscape_all(ws)
    for c in range(1, 22): ws.column_dimensions[get_column_letter(c)].width = 6.3
    for r in range(1, 50):  ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 50, 21, CREAM)
    bg(ws, 1, 1, 4,  21, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, 21)
    c.value = "Bucket List"; c.font = tf(22, bold=True); c.alignment = C
    c = mc(ws, 4, 1, 4, 21)
    c.value = "Places to go  ·  things to try  ·  memories to make"
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    cats = [("Travel & Adventures",6),("Family Experiences",18),
            ("Personal Growth",30),("Just for Fun",42)]
    for cat_name, sr in cats:
        c = mc(ws, sr, 2, sr, 21)
        c.value = cat_name; c.font = tf(12, bold=True, color=ACCENT); c.alignment = L
        bg(ws, sr, 2, sr, 21, ACCENT_LIGHT)
        for lr in range(sr+2, sr+10):
            ws.cell(lr, 2).value = "○"
            ws.cell(lr, 2).font  = mf(9, color=ACCENT)
            ws.cell(lr, 2).alignment = al("center")
            for lc in range(3, 22):
                ws.cell(lr, lc).border = bdr(bottom=sd("hair", WARM_MED))
            ws.row_dimensions[lr].height = 18
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# SELF-CARE TRACKER
# ══════════════════════════════════════════════════════════════════════════════
def make_self_care_tracker(wb):
    ws = wb.create_sheet("Self-Care Tracker")
    landscape_all(ws, scale=80)
    ws.column_dimensions["A"].width = 18
    for c in range(2, 33): ws.column_dimensions[get_column_letter(c)].width = 3.2
    ws.column_dimensions[get_column_letter(33)].width = 10

    for r in range(1, 32): ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 32, 33, CREAM)
    bg(ws, 1, 1, 3,  33, WARM_LIGHT)

    c = mc(ws, 2, 1, 2, 33)
    c.value = "Self-Care Tracker"; c.font = tf(20, bold=True); c.alignment = C
    c = mc(ws, 3, 1, 3, 33)
    c.value = "You can't pour from an empty cup. Track the habits that fill yours."
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    ws.row_dimensions[5].height = 14
    ws.cell(5, 1).value = "Habit"
    ws.cell(5, 1).font  = mf(8, bold=True, color=TEXT_MED)
    for d in range(1, 32):
        cell = ws.cell(5, d+1)
        cell.value = d; cell.font = mf(7, bold=True, color=TEXT_MED)
        cell.alignment = al("center")
    ws.cell(5, 33).value = "Notes"
    ws.cell(5, 33).font  = mf(8, bold=True, color=TEXT_MED)
    bg(ws, 5, 1, 5, 33, WARM_LIGHT)

    habits = [
        "8 hrs sleep","Drank water","Moved my body","Time outside",
        "Read / learned","Journaled","Connected with a friend",
        "Ate nourishing food","Screen-free time","Quiet time / prayer",
        "Bath / skincare","Something just for me",
    ]
    for i, habit in enumerate(habits):
        r = 7 + i*2
        ws.row_dimensions[r].height = 16
        ws.cell(r, 1).value = habit
        ws.cell(r, 1).font  = mf(8, color=TEXT_DARK)
        bg(ws, r, 1, r, 1, WARM_LIGHT)
        for d in range(1, 32):
            cell = ws.cell(r, d+1)
            cell.value = "○"
            cell.font  = Font(name="Montserrat", size=7, color=WARM_DARK)
            cell.alignment = al("center")
            cell.border = bdr(left=sd("hair",WARM_MED), bottom=sd("hair",WARM_MED))
        ws.cell(r, 33).border = bdr(bottom=sd("hair", WARM_MED))
        ws.row_dimensions[r+1].height = 3
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def build_planner():
    wb = Workbook()
    wb.remove(wb.active)

    print("Cover & intro...")
    make_cover(wb)
    make_how_to_use(wb)
    make_important_contacts(wb)

    print("Year at a Glance (Jul 2026 – Feb 2028, 5×4)...")
    make_year_view(wb)

    print("Planning pages...")
    make_brain_map(wb)
    make_brainstorm(wb)
    make_goals_page(wb, "Family")
    make_goals_page(wb, "Personal")
    make_goals_page(wb, "Professional")
    make_vision_board(wb)
    make_bucket_list(wb)
    make_self_care_tracker(wb)

    print("Monthly + weekly layouts...")
    for (year, month) in MONTHS_PLANNER:
        print(f"  {MN[month]} {year}")
        make_monthly_layout(wb, year, month)
        for week in weeks_of_month(year, month):
            make_weekly_layout(wb, week, MN[month], year)

    out = "/home/user/momlifeplanner/MomLifePlanner.xlsx"
    wb.save(out)
    print(f"\nSaved → {out}  ({len(wb.sheetnames)} sheets)")

if __name__ == "__main__":
    build_planner()
