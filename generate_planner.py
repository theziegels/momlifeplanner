"""
Ultimate Mom Life Planner – v2
Aug 2026 – Dec 2027  (17 planner months)
Year-at-a-Glance: Jul 2026 – Feb 2028 = 20 months, 5 × 4 grid
All layout pages: landscape letter, fit to one page
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
WHITE        = "FFFFFF"

# ── Font helpers ──────────────────────────────────────────────────────────────
def tf(size=11, bold=False, color=TEXT_DARK, italic=False):
    return Font(name="Cormorant Garamond", size=size, bold=bold,
                color=color, italic=italic)

def mf(size=9, bold=False, color=TEXT_DARK, italic=False):
    return Font(name="Montserrat", size=size, bold=bold,
                color=color, italic=italic)

# ── Fill / border helpers ─────────────────────────────────────────────────────
def fill(c):
    return PatternFill("solid", fgColor=c)

def sd(style="thin", color=WARM_DARK):
    return Side(style=style, color=color)

def bdr(left=None, right=None, top=None, bottom=None):
    return Border(left=left   or Side(style=None),
                  right=right or Side(style=None),
                  top=top     or Side(style=None),
                  bottom=bottom or Side(style=None))

def al(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

C = al("center"); L = al("left"); R = al("right")

# ── Range helpers ─────────────────────────────────────────────────────────────
def bg(ws, r1, c1, r2, c2, color):
    f = fill(color)
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            cell.fill = f

def bdr_range(ws, r1, c1, r2, c2, b):
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            cell.border = b

def mc(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    return ws.cell(r1, c1)

def ruled(ws, r1, r2, c1, c2, color=WARM_MED, style="hair"):
    b = bdr(bottom=sd(style, color))
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = b

def landscape_letter(ws, scale=75):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 1          # Letter
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale     = scale
    ws.page_margins.left   = 0.4
    ws.page_margins.right  = 0.4
    ws.page_margins.top    = 0.4
    ws.page_margins.bottom = 0.4

def portrait_letter(ws, scale=90):
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize   = 1
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale     = scale
    ws.page_margins.left   = 0.4
    ws.page_margins.right  = 0.4
    ws.page_margins.top    = 0.4
    ws.page_margins.bottom = 0.4

# ── Calendar data ─────────────────────────────────────────────────────────────
MONTHS_PLANNER = (
    [(2026, m) for m in range(8, 13)] +
    [(2027, m) for m in range(1, 13)]
)  # 17 months, Aug 2026 – Dec 2027

MN  = ["","January","February","March","April","May","June",
       "July","August","September","October","November","December"]
MA  = ["","Jan","Feb","Mar","Apr","May","Jun",
       "Jul","Aug","Sep","Oct","Nov","Dec"]
DF  = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

def weeks_of_month(year, month):
    first = date(year, month, 1)
    last  = date(year, month, calendar.monthrange(year, month)[1])
    start = first - timedelta(days=first.weekday())   # Monday
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
    portrait_letter(ws)
    for c in range(1, 11):
        ws.column_dimensions[get_column_letter(c)].width = 8.5
    for r in range(1, 62):
        ws.row_dimensions[r].height = 14

    bg(ws, 1, 1, 62, 10, CREAM)
    bg(ws, 1, 1, 4,  10, ACCENT_LIGHT)
    bg(ws, 1, 1, 62, 1,  WARM_MED)

    c = mc(ws, 11, 2, 11, 10)
    c.value = "THE ULTIMATE"; c.font = tf(13, italic=True, color=TEXT_MED); c.alignment = C

    c = mc(ws, 13, 2, 14, 10)
    c.value = "Mom Life"; c.font = tf(42, bold=True, color=TEXT_DARK); c.alignment = C

    c = mc(ws, 15, 2, 15, 10)
    c.value = "P  L  A  N  N  E  R"; c.font = mf(11, color=ACCENT); c.alignment = C

    for col in range(3, 10):
        ws.cell(17, col).border = bdr(bottom=sd("medium", ACCENT))

    c = mc(ws, 19, 2, 19, 10)
    c.value = "August 2026 – December 2027"
    c.font = mf(10, italic=True, color=TEXT_MED); c.alignment = C

    for box_cols in [(3,4),(6,7),(9,9)]:
        bg(ws, 25, box_cols[0], 27, box_cols[1], WARM_LIGHT)
        bdr_range(ws, 25, box_cols[0], 27, box_cols[1],
                  bdr(left=sd(), right=sd(), top=sd(), bottom=sd()))

    c = mc(ws, 38, 2, 39, 10)
    c.value = '"You are doing better than you think."'
    c.font = tf(13, italic=True, color=TEXT_MED); c.alignment = C

    c = mc(ws, 52, 2, 52, 10)
    c.value = "plan  ·  dream  ·  thrive"
    c.font = mf(9, italic=True, color=TEXT_LIGHT); c.alignment = C

    bg(ws, 59, 1, 62, 10, WARM_LIGHT)
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# HOW TO USE
# ══════════════════════════════════════════════════════════════════════════════
def make_how_to_use(wb):
    ws = wb.create_sheet("How To Use")
    portrait_letter(ws)
    for c in range(1, 13): ws.column_dimensions[get_column_letter(c)].width = 8
    for r in range(1, 80):  ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 80, 12, CREAM)
    bg(ws, 1, 1, 4,  12, WARM_LIGHT)

    c = mc(ws, 3, 1, 3, 12)
    c.value = "How to Use This Planner"; c.font = tf(22, bold=True); c.alignment = C

    items = [
        (8,  "Year at a Glance",
             "20 months on one landscape page (July 2026 – February 2028). "
             "Mark holidays, vacations, and big life moments at a glance."),
        (14, "Monthly Layout",
             "One landscape page per month. A full calendar grid with generous "
             "writing space per day, plus a notes column on the right."),
        (20, "Weekly Layout",
             "One landscape page per week. Left side: Mon–Sun in a 2-column "
             "grid with ruled lines. Right side: Priorities, Habits tracker "
             "(M T W T F S S), To-Do list, and Notes."),
        (26, "Goals Pages",
             "Separate pages for Family, Personal, and Professional goals — "
             "big goals, quarterly breakdown, and action steps."),
        (32, "Brain Map & Brainstorm",
             "Brain Map: central idea branching outward in all directions. "
             "Brainstorm: spine-and-branch layout for free-flow thinking."),
        (38, "Vision Board",
             "Paste images, write words, capture the feeling of your ideal life. "
             "Return to it often."),
        (44, "Extras",
             "Bucket List, Self-Care Tracker, and Important Contacts — "
             "because a well-planned life is a well-lived life."),
    ]
    for row, title, body in items:
        bg(ws, row, 2, row, 2, ACCENT)
        ws.row_dimensions[row].height = 18
        c = mc(ws, row, 3, row, 12)
        c.value = title; c.font = tf(13, bold=True, color=ACCENT); c.alignment = L
        c = mc(ws, row+1, 3, row+3, 12)
        c.value = body; c.font = mf(9, color=TEXT_MED); c.alignment = al("left","top")

    c = mc(ws, 60, 2, 61, 11)
    c.value = "This planner is yours — write in it, doodle in it, make it beautiful and real."
    c.font = tf(11, italic=True, color=TEXT_MED); c.alignment = C
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# IMPORTANT CONTACTS
# ══════════════════════════════════════════════════════════════════════════════
def make_important_contacts(wb):
    ws = wb.create_sheet("Important Contacts")
    portrait_letter(ws)
    for c in range(1, 15): ws.column_dimensions[get_column_letter(c)].width = 7.5
    for r in range(1, 80):  ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 80, 14, CREAM)
    bg(ws, 1, 1, 4,  14, WARM_LIGHT)

    c = mc(ws, 3, 1, 3, 14)
    c.value = "Important Contacts"; c.font = tf(22, bold=True); c.alignment = C

    cats = ["Emergency","Pediatrician","Dentist","School / Teacher",
            "Babysitter / Childcare","Husband / Partner","Mom / In-Laws",
            "Plumber / Electrician","Pharmacy","Other"]
    hdrs = ["Name","Phone","Email","Notes"]
    hcols = [2, 5, 8, 11]

    r = 8
    ws.row_dimensions[r].height = 16
    for lbl, hc in zip(hdrs, hcols):
        c = mc(ws, r, hc, r, hc+2)
        c.value = lbl; c.font = mf(8, bold=True, color=ACCENT); c.alignment = L
        bg(ws, r, hc, r, hc+2, WARM_LIGHT)

    r = 10
    for cat in cats:
        ws.row_dimensions[r].height = 16
        ws.cell(r, 1).value = cat
        ws.cell(r, 1).font  = mf(8, bold=True, color=TEXT_MED)
        for hc in hcols:
            cell = mc(ws, r, hc, r, hc+2)
            cell.border = bdr(bottom=sd("thin", WARM_DARK))
        r += 2
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# YEAR AT A GLANCE  –  5 cols × 4 rows = 20 months  (landscape)
# Jul 2026 → Feb 2028
# ══════════════════════════════════════════════════════════════════════════════
def make_year_view(wb):
    ws = wb.create_sheet("Year at a Glance")
    landscape_letter(ws, scale=80)

    # 20 months: Jul 2026 → Feb 2028
    ym_all = (
        [(2026, m) for m in range(7, 13)] +   # Jul–Dec 2026
        [(2027, m) for m in range(1, 13)] +   # Jan–Dec 2027
        [(2028, m) for m in range(1, 3)]       # Jan–Feb 2028
    )  # 6+12+2 = 20

    planner_months = set(MONTHS_PLANNER)
    faded_months   = {(2026,7), (2028,1), (2028,2)}

    # ── Column layout ─────────────────────────────────────────────────────────
    # 5 mini-cals per row, each 7 day-cols wide; 4 gap cols between them
    # col pattern: [7][1][7][1][7][1][7][1][7]  = 39 cols total
    DAY_W = 3.8
    GAP_W = 1.8

    col_starts = []   # starting col of each of the 5 calendar blocks
    c = 1
    for i in range(5):
        col_starts.append(c)
        c += 7
        if i < 4:
            c += 1   # gap

    for ci in range(1, 40):
        letter = get_column_letter(ci)
        # is this a gap col?
        is_gap = any(ci == col_starts[k] + 7 for k in range(4))
        ws.column_dimensions[letter].width = GAP_W if is_gap else DAY_W

    # ── Row layout ────────────────────────────────────────────────────────────
    # Header: rows 1-3
    # 4 calendar row-groups, each: [month-name, day-hdrs, 6 week-rows, gap-row] = 9 rows
    # Total: 3 + 4*9 = 39 rows
    TITLE_ROWS = 3
    ROWS_PER_CAL = 9   # 1 name + 1 day-hdr + 6 week-rows + 1 gap

    for r in range(1, 42):
        ws.row_dimensions[r].height = 14

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 8   # thin accent line row

    # gap rows between calendar row-groups
    for g in range(4):
        gap_r = TITLE_ROWS + (g+1)*ROWS_PER_CAL
        ws.row_dimensions[gap_r].height = 6

    # ── Background ────────────────────────────────────────────────────────────
    bg(ws, 1, 1, 41, 39, CREAM)
    bg(ws, 1, 1, 2, 39, WARM_LIGHT)

    # ── Page title ────────────────────────────────────────────────────────────
    c = mc(ws, 1, 1, 1, 39)
    c.value = "Year at a Glance  ·  2026 – 2027"
    c.font  = tf(18, bold=True, color=TEXT_DARK, italic=True)
    c.alignment = al("center")

    c = mc(ws, 2, 1, 2, 39)
    c.value = "August 2026 – December 2027"
    c.font  = mf(8, italic=True, color=TEXT_MED)
    c.alignment = al("center")

    # accent underline
    for col in range(1, 40):
        ws.cell(3, col).border = bdr(bottom=sd("medium", ACCENT))

    # ── Draw each mini-calendar ───────────────────────────────────────────────
    DAY_LETTERS = ["S","M","T","W","T","F","S"]

    for idx, (yr, mo) in enumerate(ym_all):
        grid_row = idx // 5       # 0-3
        grid_col = idx  % 5       # 0-4

        cal_c = col_starts[grid_col]
        cal_r = TITLE_ROWS + 1 + grid_row * ROWS_PER_CAL

        is_planner = (yr, mo) in planner_months
        is_faded   = (yr, mo) in faded_months

        name_color  = ACCENT    if is_planner else TEXT_LIGHT
        num_color   = TEXT_DARK if is_planner else TEXT_LIGHT
        hdr_bg      = ACCENT_LIGHT if is_planner else WARM_LIGHT

        # Month name
        ws.row_dimensions[cal_r].height = 14
        c = mc(ws, cal_r, cal_c, cal_r, cal_c+6)
        c.value     = f"{MN[mo]} {yr}" if yr != 2027 else MN[mo]
        c.font      = tf(9, bold=True, color=name_color)
        c.alignment = al("center")
        bg(ws, cal_r, cal_c, cal_r, cal_c+6, hdr_bg)

        # Day letters row
        ws.row_dimensions[cal_r+1].height = 11
        for d, ltr in enumerate(DAY_LETTERS):
            cell = ws.cell(cal_r+1, cal_c+d)
            cell.value     = ltr
            cell.font      = mf(6, bold=True,
                                color=ACCENT if d in (0,6) and is_planner else TEXT_LIGHT)
            cell.alignment = al("center")

        # Calendar grid (Sun-first weeks)
        # monthcalendar gives Mon-first; rotate to Sun-first
        cal_grid = calendar.monthcalendar(yr, mo)
        for wi, week in enumerate(cal_grid):
            sun = week[6]
            row_days = [sun] + week[:6]
            ws.row_dimensions[cal_r+2+wi].height = 13
            for d, day_num in enumerate(row_days):
                cell = ws.cell(cal_r+2+wi, cal_c+d)
                if day_num:
                    cell.value = day_num
                    cell.font  = mf(7, color=ACCENT if d==0 and is_planner else num_color)
                cell.alignment = al("center")

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# MONTHLY LAYOUT  –  one landscape page per month
# ══════════════════════════════════════════════════════════════════════════════
def make_monthly_layout(wb, year, month):
    name = f"{MA[month]} {year}"
    ws   = wb.create_sheet(name)
    landscape_letter(ws, scale=80)

    # ── Column layout ─────────────────────────────────────────────────────────
    # 7 day columns (Sun–Sat), each wide; then a thin divider; then notes column
    # Total usable landscape width ≈ 10.2" (letter minus margins)
    # 7 day cols: each 10.2*0.72/7 ≈ 10.5 "char units" → use 11.5
    # Notes col: ~2" → 16 char units, split across 2 cols
    for d in range(1, 8):
        ws.column_dimensions[get_column_letter(d)].width = 11.5
    ws.column_dimensions["H"].width = 0.8   # divider
    ws.column_dimensions["I"].width = 14    # notes col 1
    ws.column_dimensions["J"].width = 14    # notes col 2

    # ── Row layout ────────────────────────────────────────────────────────────
    # Row 1-3: header  |  rows 4+: calendar
    # Each week block: date-row (h=14) + writing lines (h=12 each, 5 lines)
    # 6 possible weeks × 6 rows = 36 rows + 3 header = 39 rows → fits fine
    WEEK_LINES = 5
    WEEK_H     = 14   # date-number row height
    LINE_H     = 11   # writing line height
    BLOCK      = 1 + WEEK_LINES   # rows per week

    for r in range(1, 50):
        ws.row_dimensions[r].height = LINE_H
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 10   # accent bar

    # ── Background ────────────────────────────────────────────────────────────
    bg(ws, 1, 1, 50, 10, CREAM)
    bg(ws, 1, 1, 3,  10, WARM_LIGHT)
    bg(ws, 1, 8, 50, 8,  WARM_MED)    # divider strip

    # ── Header ────────────────────────────────────────────────────────────────
    c = mc(ws, 1, 1, 1, 7)
    c.value = MN[month]; c.font = tf(24, bold=True, italic=True); c.alignment = L

    c = mc(ws, 2, 1, 2, 7)
    c.value = str(year); c.font = mf(10, color=TEXT_MED); c.alignment = L

    for col in range(1, 8):
        ws.cell(3, col).border = bdr(bottom=sd("medium", ACCENT))

    # ── Day-of-week headers (row 4) ───────────────────────────────────────────
    ws.row_dimensions[4].height = 14
    day_names = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
    for d, dn in enumerate(day_names):
        cell = ws.cell(4, d+1)
        cell.value     = dn
        cell.font      = mf(8, bold=True, color=ACCENT if d==0 else TEXT_MED)
        cell.alignment = al("center")
        bg(ws, 4, d+1, 4, d+1, WARM_LIGHT)

    # ── Calendar weeks ────────────────────────────────────────────────────────
    cal_grid = calendar.monthcalendar(year, month)
    row = 5
    for week in cal_grid:
        sun = week[6]; week_days = [sun] + week[:6]
        ws.row_dimensions[row].height = WEEK_H

        for d, day_num in enumerate(week_days):
            col = d + 1
            if day_num:
                num_cell = ws.cell(row, col)
                num_cell.value     = day_num
                num_cell.font      = mf(9, bold=True,
                                        color=ACCENT if d==0 else TEXT_DARK)
                num_cell.alignment = al("left", "top")
                num_cell.border    = bdr(top=sd("thin", WARM_MED),
                                         left=sd("hair", WARM_MED))
                # writing lines
                for ln in range(1, WEEK_LINES+1):
                    ws.row_dimensions[row+ln].height = LINE_H
                    ws.cell(row+ln, col).border = bdr(bottom=sd("hair", WARM_MED),
                                                       left=sd("hair", WARM_MED))
            else:
                for lr in range(row, row+BLOCK):
                    bg(ws, lr, col, lr, col, WARM_LIGHT)
        row += BLOCK

    # ── Notes sidebar ─────────────────────────────────────────────────────────
    c = mc(ws, 1, 9, 1, 10)
    c.value = "Monthly Notes"; c.font = tf(14, bold=True, color=TEXT_DARK); c.alignment = L
    bg(ws, 1, 9, 3, 10, WARM_LIGHT)

    note_sections = [
        (5,  "Appointments"),
        (14, "Birthdays & Occasions"),
        (22, "Meal Ideas"),
        (30, "Don't Forget"),
    ]
    for sr, slbl in note_sections:
        c = mc(ws, sr, 9, sr, 10)
        c.value = slbl; c.font = mf(8, bold=True, color=ACCENT); c.alignment = L
        bg(ws, sr, 9, sr, 10, ACCENT_LIGHT)
        end_r = sr + 7
        for lr in range(sr+1, end_r):
            ws.row_dimensions[lr].height = 13
            for lc in range(9, 11):
                ws.cell(lr, lc).border = bdr(bottom=sd("hair", WARM_MED))

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# WEEKLY LAYOUT  –  one landscape 8.5×11 page per week
# Left 55%:  header + 7 days in 2-col grid (Mon/Tue, Wed/Thu, Fri/Sat, Sun full)
# Right 45%: Priorities box | Habits tracker, then To-Do, then Notes
# ══════════════════════════════════════════════════════════════════════════════
def make_weekly_layout(wb, week_dates, month_name, year):
    start_d = week_dates[0]
    sheet_name = f"Wk {start_d.strftime('%b %-d')}"
    ws = wb.create_sheet(sheet_name)
    landscape_letter(ws, scale=80)

    # ── Column layout ─────────────────────────────────────────────────────────
    # Left section  cols 1–13: days
    #   Left day-block  cols 1–6  (Mon/Wed/Fri/Sun-left-half)
    #   Mid gap         col  7
    #   Right day-block cols 8–13 (Tue/Thu/Sat/Sun-right-half)
    # Divider          col 14
    # Right section    cols 15–26:
    #   Priorities      cols 15–19  (5 cols)
    #   Habits          cols 20–26  (7 cols – one per day)

    DAY_COL_W = 10.5
    GAP_W     = 1.5
    DIV_W     = 1.0
    PRI_W     = 9.8
    HAB_W     = 4.5

    for c in list(range(1,7)) + list(range(8,14)):
        ws.column_dimensions[get_column_letter(c)].width = DAY_COL_W
    ws.column_dimensions["G"].width  = GAP_W
    ws.column_dimensions["N"].width  = DIV_W
    for c in range(15, 20):
        ws.column_dimensions[get_column_letter(c)].width = PRI_W
    for c in range(20, 27):
        ws.column_dimensions[get_column_letter(c)].width = HAB_W

    # ── Row heights ────────────────────────────────────────────────────────────
    for r in range(1, 58): ws.row_dimensions[r].height = 13.5
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 8    # accent bar
    ws.row_dimensions[4].height = 6    # small gap before days

    # ── Background ────────────────────────────────────────────────────────────
    bg(ws, 1, 1,  57, 13, CREAM)
    bg(ws, 1, 14, 57, 14, WARM_MED)   # spine divider
    bg(ws, 1, 15, 57, 26, CREAM)

    # ── LEFT: Header ──────────────────────────────────────────────────────────
    bg(ws, 1, 1, 2, 13, WARM_LIGHT)

    end_d = week_dates[6]
    c = mc(ws, 1, 1, 2, 9)
    c.value     = f"{month_name.lower()}  {year}"
    c.font      = tf(18, bold=True, italic=True, color=TEXT_DARK)
    c.alignment = al("left", "center")

    c = mc(ws, 1, 10, 2, 13)
    c.value     = f"{start_d.strftime('%-m/%-d')} – {end_d.strftime('%-m/%-d')}"
    c.font      = mf(8, italic=True, color=TEXT_MED)
    c.alignment = al("right", "center")

    for col in range(1, 14):
        ws.cell(3, col).border = bdr(bottom=sd("medium", ACCENT))

    # ── LEFT: Day blocks ──────────────────────────────────────────────────────
    # Mon/Tue  row 5–17   (13 rows: 1 hdr + 12 lines)
    # Wed/Thu  row 18–30
    # Fri/Sat  row 31–43
    # Sun      row 44–56  (full width cols 1–13)
    LINES  = 11   # writing lines per day
    BLOCK  = 1 + LINES + 1   # header + lines + gap-row = 13

    def draw_day(day_idx, r_start, c_start, c_end):
        d = week_dates[day_idx]
        is_wknd = d.weekday() >= 5

        ws.row_dimensions[r_start].height = 15

        num = ws.cell(r_start, c_start)
        num.value = d.day
        num.font  = mf(9, bold=True, color=ACCENT if is_wknd else TEXT_DARK)
        num.alignment = al("left", "center")

        nm = mc(ws, r_start, c_start+1, r_start, c_end)
        nm.value = DF[day_idx].upper()
        nm.font  = mf(7, color=TEXT_MED)
        nm.alignment = al("left", "center")

        for col in range(c_start, c_end+1):
            ws.cell(r_start, col).border = bdr(bottom=sd("thin", WARM_MED))

        for ln in range(1, LINES+1):
            ws.row_dimensions[r_start+ln].height = 13.5
            for col in range(c_start, c_end+1):
                ws.cell(r_start+ln, col).border = bdr(bottom=sd("hair", WARM_MED))

        ws.row_dimensions[r_start+LINES+1].height = 4   # gap row

    pairs     = [(0,1),(2,3),(4,5)]
    row_bases = [5, 5+BLOCK, 5+BLOCK*2]

    for (li, ri), rb in zip(pairs, row_bases):
        draw_day(li, rb,        1,  6)
        draw_day(ri, rb,        8, 13)

    # Sunday full width
    sun_r = 5 + BLOCK*3
    draw_day(6, sun_r, 1, 13)

    # ── RIGHT: Header band ─────────────────────────────────────────────────────
    bg(ws, 1, 15, 2, 26, WARM_LIGHT)

    # ── RIGHT: PRIORITIES (rows 1–18, cols 15–19) ─────────────────────────────
    c = mc(ws, 1, 15, 2, 19)
    c.value = "PRIORITIES"; c.font = mf(7, bold=True, color=ACCENT); c.alignment = C
    bg(ws, 1, 15, 2, 19, ACCENT_LIGHT)

    # outer box border
    bdr_range(ws, 3, 15, 18, 19,
              bdr(left=sd("thin",WARM_DARK), right=sd("thin",WARM_DARK),
                  top=sd("thin",WARM_DARK),  bottom=sd("thin",WARM_DARK)))
    # inner ruled lines
    ruled(ws, 4, 17, 15, 19, WARM_MED, "hair")

    # ── RIGHT: HABITS (rows 1–18, cols 20–26) ─────────────────────────────────
    c = mc(ws, 1, 20, 2, 26)
    c.value = "HABITS"; c.font = mf(7, bold=True, color=TEXT_MED); c.alignment = C
    bg(ws, 1, 20, 2, 26, WARM_LIGHT)

    # Day-letter headers row 3
    ws.row_dimensions[3].height = 11
    day_ltrs = ["M","T","W","T","F","S","S"]
    for di, dl in enumerate(day_ltrs):
        cell = ws.cell(3, 20+di)
        cell.value = dl
        cell.font  = mf(6, bold=True, color=ACCENT if di>=5 else TEXT_MED)
        cell.alignment = al("center")

    # 6 habit rows  (rows 4–18, every 2 rows)
    for h in range(6):
        hr = 4 + h*2
        ws.row_dimensions[hr].height = 16

        # habit-name cell (col 20, accent fill)
        bg(ws, hr, 20, hr, 20, ACCENT_LIGHT)
        ws.cell(hr, 20).border = bdr(bottom=sd("thin", WARM_DARK))

        for di in range(7):
            cell = ws.cell(hr, 20+di)
            if di > 0:
                cell.value     = "○"
                cell.font      = Font(name="Montserrat", size=7, color=WARM_DARK)
                cell.alignment = al("center")
            cell.border = bdr(bottom=sd("thin", WARM_MED),
                              left=sd("hair", WARM_MED))

        ws.row_dimensions[hr+1].height = 3   # small gap

    # ── RIGHT: TO-DO (rows 20–35) ─────────────────────────────────────────────
    TODO_START = 20
    bg(ws, TODO_START, 15, TODO_START+1, 26, WARM_LIGHT)

    c = mc(ws, TODO_START, 15, TODO_START+1, 19)
    c.value = "TO-DO"; c.font = mf(7, bold=True, color=TEXT_MED); c.alignment = al("left","center")

    c = mc(ws, TODO_START, 20, TODO_START+1, 26)
    c.value = "TO-DO  (cont'd)"; c.font = mf(7, bold=True, color=TEXT_MED); c.alignment = al("left","center")

    for ti in range(9):
        tr = TODO_START + 2 + ti
        ws.row_dimensions[tr].height = 14

        ws.cell(tr, 15).value     = "○"
        ws.cell(tr, 15).font      = Font(name="Montserrat", size=8, color=ACCENT)
        ws.cell(tr, 15).alignment = al("center")
        for lc in range(16, 20):
            ws.cell(tr, lc).border = bdr(bottom=sd("hair", WARM_MED))

        ws.cell(tr, 20).value     = "○"
        ws.cell(tr, 20).font      = Font(name="Montserrat", size=8, color=ACCENT)
        ws.cell(tr, 20).alignment = al("center")
        for lc in range(21, 27):
            ws.cell(tr, lc).border = bdr(bottom=sd("hair", WARM_MED))

    # ── RIGHT: NOTES (rows 37–57) ─────────────────────────────────────────────
    NOTES_START = 37
    bg(ws, NOTES_START, 15, NOTES_START+1, 26, ACCENT_LIGHT)

    c = mc(ws, NOTES_START, 15, NOTES_START+1, 26)
    c.value = "NOTES"; c.font = mf(7, bold=True, color=ACCENT); c.alignment = al("left","center")

    for nr in range(NOTES_START+2, 57):
        ws.row_dimensions[nr].height = 14
        ruled(ws, nr, nr, 15, 26, WARM_MED, "hair")

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# GOALS PAGE
# ══════════════════════════════════════════════════════════════════════════════
def make_goals_page(wb, goal_type):
    ws = wb.create_sheet(f"{goal_type} Goals")
    portrait_letter(ws)
    for c in range(1, 15): ws.column_dimensions[get_column_letter(c)].width = 7.5
    for r in range(1, 80):  ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 80, 14, CREAM)
    bg(ws, 1, 1, 4,  14, WARM_LIGHT)

    c = mc(ws, 3, 1, 3, 14)
    c.value = f"{goal_type} Goals"; c.font = tf(22, bold=True); c.alignment = C

    sections = [
        (8,  "My Big Goals This Year",   22, True),
        (24, "Quarter by Quarter",        44, False),
        (46, "Action Steps & Milestones", 72, True),
    ]
    for s_row, s_title, e_row, is_lined in sections:
        c = mc(ws, s_row, 2, s_row, 13)
        c.value = s_title; c.font = tf(13, bold=True, color=ACCENT); c.alignment = L
        bg(ws, s_row, 2, s_row, 13, ACCENT_LIGHT)

        if "Quarter" in s_title:
            q_labels = ["Q1  Jan–Mar","Q2  Apr–Jun","Q3  Jul–Sep","Q4  Oct–Dec"]
            q_starts = [2, 5, 9, 12]
            for q_lbl, qc in zip(q_labels, q_starts):
                c = mc(ws, s_row+2, qc, s_row+2, qc+2)
                c.value = q_lbl; c.font = mf(8, bold=True, color=TEXT_MED); c.alignment = L
                for lr in range(s_row+3, s_row+14):
                    for lc in range(qc, qc+3):
                        ws.cell(lr, lc).border = bdr(bottom=sd("hair", WARM_MED))
                    ws.row_dimensions[lr].height = 16
        else:
            for i, lr in enumerate(range(s_row+2, e_row)):
                if is_lined and "Big" in s_title:
                    ws.cell(lr, 2).value = f"{i+1}."
                    ws.cell(lr, 2).font  = mf(9, color=ACCENT)
                for lc in range(3, 14):
                    ws.cell(lr, lc).border = bdr(bottom=sd("hair", WARM_MED))
                ws.row_dimensions[lr].height = 18
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# BRAIN MAP
# ══════════════════════════════════════════════════════════════════════════════
def make_brain_map(wb):
    ws = wb.create_sheet("Brain Map")
    portrait_letter(ws)
    for c in range(1, 23): ws.column_dimensions[get_column_letter(c)].width = 5
    for r in range(1, 70):  ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 70, 22, CREAM)
    bg(ws, 1, 1, 4,  22, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, 22)
    c.value = "Brain Map"; c.font = tf(22, bold=True); c.alignment = C

    c = mc(ws, 4, 1, 4, 22)
    c.value = "Place your central idea in the center, then branch outward."
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    # Central box
    cr, cc = 35, 10
    bg(ws, cr-2, cc-1, cr+2, cc+2, ACCENT_LIGHT)
    cell = mc(ws, cr-2, cc-1, cr+2, cc+2)
    cell.value = "Central\nIdea"; cell.font = tf(12, bold=True, color=ACCENT)
    cell.alignment = al("center")
    bdr_range(ws, cr-2, cc-1, cr+2, cc+2,
              bdr(left=sd("medium",ACCENT),right=sd("medium",ACCENT),
                  top=sd("medium",ACCENT),bottom=sd("medium",ACCENT)))

    # Branches: (box_row, box_col_start, box_col_end, label)
    branch_boxes = [
        (10, 3, 5), (10, 9, 11), (10, 15, 17),
        (35, 2, 4),              (35, 16, 18),
        (57, 3, 5), (57, 9, 11),(57, 15, 17),
    ]
    for (br_r, bc1, bc2) in branch_boxes:
        bg(ws, br_r, bc1, br_r+3, bc2, WARM_LIGHT)
        cell = mc(ws, br_r, bc1, br_r+3, bc2)
        cell.value = ""; cell.alignment = al("center")
        bdr_range(ws, br_r, bc1, br_r+3, bc2,
                  bdr(left=sd(),right=sd(),top=sd(),bottom=sd()))
        # sub-branch lines
        for sub_r in range(br_r, br_r+4):
            for sub_c in range(max(1,bc1-3), bc1):
                ws.cell(sub_r, sub_c).border = bdr(bottom=sd("hair", WARM_MED))
            for sub_c in range(bc2+1, min(22, bc2+4)):
                ws.cell(sub_r, sub_c).border = bdr(bottom=sd("hair", WARM_MED))
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# BRAINSTORM
# ══════════════════════════════════════════════════════════════════════════════
def make_brainstorm(wb):
    ws = wb.create_sheet("Brainstorm")
    portrait_letter(ws)
    for c in range(1, 23): ws.column_dimensions[get_column_letter(c)].width = 5
    for r in range(1, 75):  ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 75, 22, CREAM)
    bg(ws, 1, 1, 4,  22, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, 22)
    c.value = "Brainstorm"; c.font = tf(22, bold=True); c.alignment = C

    c = mc(ws, 4, 1, 4, 22)
    c.value = "Start with your main idea on the left. Let every branch spark the next."
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    # Horizontal spine
    spine_row = 38
    for col in range(2, 22):
        ws.cell(spine_row, col).border = bdr(bottom=sd("medium", TEXT_MED))

    # Main idea box (left)
    bg(ws, spine_row-1, 1, spine_row+1, 3, ACCENT_LIGHT)
    cell = mc(ws, spine_row-1, 1, spine_row+1, 3)
    cell.value = "Main\nIdea"; cell.font = tf(10, bold=True, color=ACCENT)
    cell.alignment = al("center")
    bdr_range(ws, spine_row-1, 1, spine_row+1, 3,
              bdr(left=sd("medium",ACCENT),right=sd("medium",ACCENT),
                  top=sd("medium",ACCENT),bottom=sd("medium",ACCENT)))

    # Branch positions
    for bc in [5, 8, 11, 14, 17, 20]:
        for arm in range(1, 5):
            for r_dir, sign in [(-1, -1), (1, 1)]:
                br = spine_row + sign * arm * 4
                if 6 <= br <= 70:
                    c = mc(ws, br, bc, br, bc+1)
                    c.border = bdr(bottom=sd("thin", WARM_DARK))
                    for leaf in range(1, 3):
                        lc = bc - leaf
                        if lc >= 1:
                            ws.cell(br, lc).border = bdr(bottom=sd("hair", WARM_MED))
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# VISION BOARD
# ══════════════════════════════════════════════════════════════════════════════
def make_vision_board(wb):
    ws = wb.create_sheet("Vision Board")
    portrait_letter(ws)
    for c in range(1, 19): ws.column_dimensions[get_column_letter(c)].width = 6.5
    for r in range(1, 75):  ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 75, 18, CREAM)
    bg(ws, 1, 1, 5,  18, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, 18)
    c.value = "Vision Board"; c.font = tf(26, bold=True); c.alignment = C

    c = mc(ws, 4, 1, 4, 18)
    c.value = "Paste images  ·  write words  ·  capture feelings  ·  dream boldly"
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    words = ["FAITH","FAMILY","HEALTH","JOY","GROWTH","PEACE",
             "ADVENTURE","HOME","PURPOSE","LOVE","ABUNDANCE","FREEDOM"]
    for i, word in enumerate(words):
        wr = 7 + (i//4)*3; wc = 2 + (i%4)*4
        cell = mc(ws, wr, wc, wr+1, wc+2)
        cell.value = word; cell.font = mf(9, bold=True, color=TEXT_LIGHT)
        cell.alignment = al("center")
        bg(ws, wr, wc, wr+1, wc+2, WARM_LIGHT)
        bdr_range(ws, wr, wc, wr+1, wc+2, bdr(left=sd(),right=sd(),top=sd(),bottom=sd()))

    boxes = [
        (20, 2, 40, 9,   "paste an image or write your dream here"),
        (20, 10, 40, 17, "paste an image or write your dream here"),
        (42, 2, 58, 9,   "this year I will..."),
        (42, 10, 58, 17, "I am grateful for..."),
        (60, 2, 70, 17,  "my word(s) for this season:"),
    ]
    for r1, c1, r2, c2, hint in boxes:
        bg(ws, r1, c1, r2, c2, WARM_LIGHT)
        bdr_range(ws, r1, c1, r2, c2, bdr(left=sd("thin",WARM_MED),right=sd("thin",WARM_MED),
                                           top=sd("thin",WARM_MED),bottom=sd("thin",WARM_MED)))
        cell = mc(ws, r1, c1, r2, c2)
        cell.value = hint; cell.font = mf(8, italic=True, color=TEXT_LIGHT)
        cell.alignment = al("center")
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# BUCKET LIST
# ══════════════════════════════════════════════════════════════════════════════
def make_bucket_list(wb):
    ws = wb.create_sheet("Bucket List")
    portrait_letter(ws)
    for c in range(1, 15): ws.column_dimensions[get_column_letter(c)].width = 7.5
    for r in range(1, 80):  ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 80, 14, CREAM)
    bg(ws, 1, 1, 5,  14, WARM_LIGHT)

    c = mc(ws, 3, 1, 3, 14)
    c.value = "Bucket List"; c.font = tf(22, bold=True); c.alignment = C

    c = mc(ws, 4, 1, 4, 14)
    c.value = "Places to go  ·  things to try  ·  memories to make"
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    cats = [("Travel & Adventures",8),("Family Experiences",26),
            ("Personal Growth",44),("Just for Fun",60)]
    for cat_name, sr in cats:
        c = mc(ws, sr, 2, sr, 13)
        c.value = cat_name; c.font = tf(12, bold=True, color=ACCENT); c.alignment = L
        bg(ws, sr, 2, sr, 13, ACCENT_LIGHT)
        for lr in range(sr+2, sr+14):
            ws.cell(lr, 2).value = "○"
            ws.cell(lr, 2).font  = mf(9, color=ACCENT)
            ws.cell(lr, 2).alignment = al("center")
            for lc in range(3, 14):
                ws.cell(lr, lc).border = bdr(bottom=sd("hair", WARM_MED))
            ws.row_dimensions[lr].height = 18
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# SELF-CARE TRACKER
# ══════════════════════════════════════════════════════════════════════════════
def make_self_care_tracker(wb):
    ws = wb.create_sheet("Self-Care Tracker")
    landscape_letter(ws, scale=80)

    # 33 cols: col 1 = habit name (wide), cols 2-32 = days 1-31, col 33 = notes
    ws.column_dimensions["A"].width = 18
    for c in range(2, 33):
        ws.column_dimensions[get_column_letter(c)].width = 3.2
    ws.column_dimensions[get_column_letter(33)].width = 10

    for r in range(1, 30): ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 30, 33, CREAM)
    bg(ws, 1, 1, 3,  33, WARM_LIGHT)

    c = mc(ws, 2, 1, 2, 33)
    c.value = "Self-Care Tracker"; c.font = tf(20, bold=True); c.alignment = C

    c = mc(ws, 3, 1, 3, 33)
    c.value = "You can't pour from an empty cup. Track the habits that fill yours."
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    # Day headers
    ws.row_dimensions[5].height = 14
    ws.cell(5, 1).value = "Habit"
    ws.cell(5, 1).font  = mf(8, bold=True, color=TEXT_MED)
    for d in range(1, 32):
        cell = ws.cell(5, d+1)
        cell.value = d; cell.font = mf(7, bold=True, color=TEXT_MED); cell.alignment = al("center")
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
        cell = ws.cell(r, 1)
        cell.value = habit; cell.font = mf(8, color=TEXT_DARK)
        bg(ws, r, 1, r, 1, WARM_LIGHT)

        for d in range(1, 32):
            cell = ws.cell(r, d+1)
            cell.value     = "○"
            cell.font      = Font(name="Montserrat", size=7, color=WARM_DARK)
            cell.alignment = al("center")
            cell.border    = bdr(left=sd("hair",WARM_MED), bottom=sd("hair",WARM_MED))

        ws.cell(r, 33).border = bdr(bottom=sd("hair", WARM_MED))
        ws.row_dimensions[r+1].height = 3
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# MAIN BUILD
# ══════════════════════════════════════════════════════════════════════════════
def build_planner():
    wb = Workbook()
    wb.remove(wb.active)

    print("Building cover & intro pages...")
    make_cover(wb)
    make_how_to_use(wb)
    make_important_contacts(wb)

    print("Building Year at a Glance (Jul 2026 – Feb 2028, 5×4)...")
    make_year_view(wb)

    print("Building planning pages...")
    make_brain_map(wb)
    make_brainstorm(wb)
    make_goals_page(wb, "Family")
    make_goals_page(wb, "Personal")
    make_goals_page(wb, "Professional")
    make_vision_board(wb)
    make_bucket_list(wb)
    make_self_care_tracker(wb)

    print("Building monthly + weekly layouts...")
    for (year, month) in MONTHS_PLANNER:
        print(f"  {MN[month]} {year}")
        make_monthly_layout(wb, year, month)
        for week in weeks_of_month(year, month):
            make_weekly_layout(wb, week, MN[month], year)

    out = "/home/user/momlifeplanner/MomLifePlanner.xlsx"
    wb.save(out)
    print(f"\nSaved → {out}  ({len(wb.sheetnames)} sheets)")
    return out

if __name__ == "__main__":
    build_planner()
