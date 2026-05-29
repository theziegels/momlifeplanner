"""
Ultimate Mom Life Planner – v4
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

def dot_line(ws, r, c1, c2, color=WARM_DARK):
    """Single dotted ruling line across a row."""
    b = bdr(bottom=sd("dotted", color))
    for c in range(c1, c2+1):
        ws.cell(r, c).border = b

def dot_lines(ws, r1, r2, c1, c2, color=WARM_DARK):
    """Dotted ruling lines for a range of rows."""
    for r in range(r1, r2+1):
        dot_line(ws, r, c1, c2, color)

def landscape_all(ws, scale=80):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 1
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale     = scale
    ws.page_margins.left   = 0.35
    ws.page_margins.right  = 0.35
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
    """Only weeks whose Monday falls within the month — eliminates duplicates."""
    first = date(year, month, 1)
    last  = date(year, month, calendar.monthrange(year, month)[1])
    # First Monday on-or-after the 1st
    days_ahead = (7 - first.weekday()) % 7
    start = first if first.weekday() == 0 else first + timedelta(days=days_ahead)
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

    TOTAL_COLS = 20
    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 6.8
    for r in range(1, 50):
        ws.row_dimensions[r].height = 14

    bg(ws, 1, 1, 50, TOTAL_COLS, CREAM)

    # Border strips — all WARM_MED, all four sides
    STRIP = 1   # col/row width of border strip
    bg(ws, 1,    1,    3,    TOTAL_COLS, WARM_MED)   # top bar
    bg(ws, 48,   1,    50,   TOTAL_COLS, WARM_MED)   # bottom bar
    bg(ws, 1,    1,    50,   STRIP,      WARM_MED)   # left strip
    bg(ws, 1,    TOTAL_COLS, 50, TOTAL_COLS, WARM_MED)  # right strip

    # Thin accent rule just inside the top bar
    for col in range(2, TOTAL_COLS):
        ws.cell(4, col).border = bdr(bottom=sd("thin", ACCENT))

    # ── Title block
    ws.row_dimensions[9].height  = 18
    ws.row_dimensions[10].height = 18
    ws.row_dimensions[12].height = 55
    ws.row_dimensions[13].height = 55
    ws.row_dimensions[14].height = 18

    c = mc(ws, 9, 2, 9, TOTAL_COLS-1)
    c.value = "THE ULTIMATE"; c.font = tf(13, italic=True, color=TEXT_MED)
    c.alignment = al("center")

    c = mc(ws, 12, 2, 13, TOTAL_COLS-1)
    c.value = "Mom Life"; c.font = tf(52, bold=True, color=TEXT_DARK)
    c.alignment = al("center")

    c = mc(ws, 14, 2, 14, TOTAL_COLS-1)
    c.value = "P  L  A  N  N  E  R"; c.font = mf(13, color=ACCENT)
    c.alignment = al("center")

    # Thin accent rule under title
    for col in range(4, TOTAL_COLS-2):
        ws.cell(15, col).border = bdr(bottom=sd("thin", ACCENT))

    c = mc(ws, 17, 2, 17, TOTAL_COLS-1)
    c.value = "August 2026 – December 2027"
    c.font = mf(10, italic=True, color=TEXT_MED); c.alignment = al("center")

    # ── Three quote boxes
    ws.row_dimensions[22].height = 16
    ws.row_dimensions[23].height = 20
    ws.row_dimensions[24].height = 20
    ws.row_dimensions[25].height = 16

    quote_boxes = [
        (2, 7,   "Be present.\nBe patient.\nBe kind to yourself."),
        (9, 13,  "You are enough\nfor today."),
        (15, TOTAL_COLS-1, "Small steps\nstill move you\nforward."),
    ]
    for c1, c2, quote in quote_boxes:
        bg(ws, 22, c1, 25, c2, WARM_LIGHT)
        bdr_range(ws, 22, c1, 25, c2,
                  bdr(left=sd("thin",WARM_MED), right=sd("thin",WARM_MED),
                      top=sd("thin",WARM_MED),  bottom=sd("thin",WARM_MED)))
        cell = mc(ws, 22, c1, 25, c2)
        cell.value = quote
        cell.font  = tf(9, italic=True, color=TEXT_MED)
        cell.alignment = al("center", "center")

    # ── Inspirational quote
    ws.row_dimensions[33].height = 18
    ws.row_dimensions[34].height = 18
    c = mc(ws, 33, 3, 34, TOTAL_COLS-2)
    c.value = '"You are doing better than you think."'
    c.font  = tf(14, italic=True, color=TEXT_MED); c.alignment = al("center")

    # ── Tagline
    ws.row_dimensions[40].height = 16
    c = mc(ws, 40, 2, 40, TOTAL_COLS-1)
    c.value = "plan  ·  dream  ·  thrive"
    c.font  = mf(9, italic=True, color=TEXT_LIGHT); c.alignment = al("center")

    # ── Branding
    ws.row_dimensions[44].height = 18
    ws.row_dimensions[45].height = 16
    _add_branding(ws, 44, 2, TOTAL_COLS-1)

    return ws


def _add_branding(ws, r, c1, c2):
    """Adds two-part brand name: bold-italic Collective title + lighter subtitle."""
    mid = (c1 + c2) // 2
    cell_a = mc(ws, r, c1, r, mid)
    cell_a.value = "Beloved Kingdom Collective™"
    cell_a.font  = tf(10, bold=True, italic=True, color=TEXT_MED)
    cell_a.alignment = al("right")

    cell_b = mc(ws, r, mid+1, r, c2)
    cell_b.value = "  ·  Mom Life Planner"
    cell_b.font  = mf(9, italic=False, color=TEXT_LIGHT)
    cell_b.alignment = al("left")


# ══════════════════════════════════════════════════════════════════════════════
# HOW TO USE
# ══════════════════════════════════════════════════════════════════════════════
def make_how_to_use(wb):
    ws = wb.create_sheet("How To Use")
    landscape_all(ws)

    TOTAL_COLS = 22
    # Col 1: narrow left margin; col 2: thin accent strip; cols 3+: content
    ws.column_dimensions["A"].width = 1.5
    ws.column_dimensions["B"].width = 2.0   # narrow accent strip
    for c in range(3, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 6.2

    for r in range(1, 56): ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 56, TOTAL_COLS, CREAM)
    bg(ws, 1, 1, 4,  TOTAL_COLS, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, TOTAL_COLS)
    c.value = "How to Use This Planner"; c.font = tf(24, bold=True); c.alignment = C

    items = [
        (6,  "Year at a Glance",
             "20 months on one landscape page (July 2026 – February 2028). "
             "Mark holidays, vacations, and big life moments at a glance."),
        (12, "Monthly Layout",
             "One landscape page per month. Full calendar grid with generous "
             "writing space per day, plus a notes column on the right."),
        (18, "Weekly Layout",
             "One landscape page per week. Left: Mon–Sun in a 2-column grid "
             "with ruled writing lines. Right: Priorities, To-Do list, Notes."),
        (24, "Goals Pages",
             "Separate pages for Family, Personal, and Professional goals — "
             "big goals, quarterly breakdown, and action steps."),
        (30, "Brain Map & Brainstorm",
             "Brain Map: central idea branching outward in all directions. "
             "Brainstorm: spine-and-branch layout for free-flow thinking."),
        (36, "Vision Board + Extras",
             "Vision Board, Bucket List, Self-Care Tracker, and Important Contacts — "
             "because a well-planned life is a well-lived life."),
    ]

    for row, title, body in items:
        # Accent strip spans title row AND all body rows (row to row+4)
        bg(ws, row, 2, row+4, 2, ACCENT)
        ws.row_dimensions[row].height   = 20
        ws.row_dimensions[row+1].height = 14
        ws.row_dimensions[row+2].height = 14
        ws.row_dimensions[row+3].height = 14
        ws.row_dimensions[row+4].height = 8   # gap after section

        c = mc(ws, row, 3, row, TOTAL_COLS)
        c.value = title; c.font = tf(14, bold=True, color=ACCENT); c.alignment = L

        c = mc(ws, row+1, 3, row+3, TOTAL_COLS)
        c.value = body; c.font = mf(9, color=TEXT_MED)
        c.alignment = al("left", "top")

    c = mc(ws, 52, 3, 53, TOTAL_COLS-1)
    c.value = "This planner is yours — write in it, doodle in it, make it beautiful and real."
    c.font = tf(11, italic=True, color=TEXT_MED); c.alignment = C
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# IMPORTANT CONTACTS
# ══════════════════════════════════════════════════════════════════════════════
def make_important_contacts(wb):
    ws = wb.create_sheet("Important Contacts")
    landscape_all(ws)

    # Col A: category label (wide); cols 2-5: Name; 6: gap; 7-10: Phone;
    # 11: gap; 12-16: Email; 17: gap; 18-22: Notes
    ws.column_dimensions["A"].width = 20    # category label — wide enough for full name
    for c in range(2, 6):   ws.column_dimensions[get_column_letter(c)].width = 8   # Name
    ws.column_dimensions["F"].width = 1.5   # gap
    for c in range(7, 11):  ws.column_dimensions[get_column_letter(c)].width = 8   # Phone
    ws.column_dimensions["K"].width = 1.5   # gap
    for c in range(12, 17): ws.column_dimensions[get_column_letter(c)].width = 7   # Email
    ws.column_dimensions["Q"].width = 2.0   # gap — clear separation before Notes
    for c in range(18, 23): ws.column_dimensions[get_column_letter(c)].width = 8   # Notes

    for r in range(1, 50): ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 50, 22, CREAM)
    bg(ws, 1, 1, 4,  22, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, 22)
    c.value = "Important Contacts"; c.font = tf(24, bold=True); c.alignment = C

    # Column headers
    r = 7; ws.row_dimensions[r].height = 18
    hdrs = [("Name",2,5),("Phone",7,10),("Email",12,16),("Notes",18,22)]
    for lbl, hc1, hc2 in hdrs:
        c = mc(ws, r, hc1, r, hc2)
        c.value = lbl; c.font = mf(9, bold=True, color=ACCENT); c.alignment = L
        bg(ws, r, hc1, r, hc2, WARM_LIGHT)

    cats = ["Emergency","Pediatrician","Dentist","School / Teacher",
            "Babysitter / Childcare","Husband / Partner","Mom / In-Laws",
            "Plumber / Electrician","Pharmacy","Other"]

    r = 9
    for cat in cats:
        ws.row_dimensions[r].height = 20
        ws.cell(r, 1).value = cat
        ws.cell(r, 1).font  = mf(9, bold=True, color=TEXT_MED)
        ws.cell(r, 1).alignment = al("left","center")
        for hc1, hc2 in [(2,5),(7,10),(12,16),(18,22)]:
            cell = mc(ws, r, hc1, r, hc2)
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

    DAY_W, GAP_W = 3.9, 1.6
    col_starts = []
    c = 1
    for i in range(5):
        col_starts.append(c); c += 7
        if i < 4: c += 1

    for ci in range(1, 40):
        is_gap = any(ci == col_starts[k]+7 for k in range(4))
        ws.column_dimensions[get_column_letter(ci)].width = GAP_W if is_gap else DAY_W

    ROWS_PER = 9
    for r in range(1, 44): ws.row_dimensions[r].height = 13
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 6
    for g in range(4):
        ws.row_dimensions[3 + (g+1)*ROWS_PER].height = 5

    bg(ws, 1, 1, 43, 39, CREAM)
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
        gcol = idx % 5; grow = idx // 5
        cal_c = col_starts[gcol]
        cal_r = 4 + grow * ROWS_PER

        is_planner = (yr, mo) in planner_months
        name_color = ACCENT if is_planner else TEXT_LIGHT
        num_color  = TEXT_DARK if is_planner else TEXT_LIGHT
        hdr_bg     = ACCENT_LIGHT if is_planner else WARM_LIGHT

        ws.row_dimensions[cal_r].height = 14
        c = mc(ws, cal_r, cal_c, cal_r, cal_c+6)
        label = MN[mo] if yr == 2027 else f"{MN[mo]} '{str(yr)[2:]}"
        c.value = label; c.font = tf(8, bold=True, color=name_color)
        c.alignment = al("center")
        bg(ws, cal_r, cal_c, cal_r, cal_c+6, hdr_bg)

        ws.row_dimensions[cal_r+1].height = 10
        for d, ltr in enumerate(DAY_LTRS):
            cell = ws.cell(cal_r+1, cal_c+d)
            cell.value = ltr
            cell.font  = mf(6, bold=True,
                            color=ACCENT if d in (0,6) and is_planner else TEXT_LIGHT)
            cell.alignment = al("center")

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

    for d in range(1, 8):  ws.column_dimensions[get_column_letter(d)].width = 12.5
    ws.column_dimensions["H"].width = 0.8
    ws.column_dimensions["I"].width = 16
    ws.column_dimensions["J"].width = 16

    LINES = 5; BLOCK = 1 + LINES

    for r in range(1, 52): ws.row_dimensions[r].height = 12
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 14
    ws.row_dimensions[4].height = 8
    ws.row_dimensions[5].height = 14

    bg(ws, 1, 1, 52, 10, CREAM)
    bg(ws, 1, 1, 4,  10, WARM_LIGHT)
    bg(ws, 1, 8, 52,  8, WARM_MED)

    c = mc(ws, 1, 1, 2, 7)
    c.value = MN[month]; c.font = tf(26, bold=True, italic=True); c.alignment = L
    c = mc(ws, 3, 1, 3, 7)
    c.value = str(year); c.font = mf(10, color=TEXT_MED); c.alignment = L
    for col in range(1, 8):
        ws.cell(4, col).border = bdr(bottom=sd("medium", ACCENT))

    day_names = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
    for d, dn in enumerate(day_names):
        cell = ws.cell(5, d+1)
        cell.value = dn
        cell.font  = mf(8, bold=True, color=ACCENT if d==0 else TEXT_MED)
        cell.alignment = al("center")
        bg(ws, 5, d+1, 5, d+1, WARM_LIGHT)

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
                    ws.cell(row+ln, col).border = bdr(bottom=sd("dotted",WARM_DARK),
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
            ws.row_dimensions[lr].height = 14
            dot_line(ws, lr, 9, 10)

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# WEEKLY LAYOUT  –  landscape, one page per week
# ══════════════════════════════════════════════════════════════════════════════
def make_weekly_layout(wb, week_dates, month_name, year):
    start_d = week_dates[0]
    end_d   = week_dates[6]
    ws = wb.create_sheet(f"Wk {start_d.strftime('%b %-d')}")
    landscape_all(ws, scale=80)

    # ── Columns ───────────────────────────────────────────────────────────────
    for c in list(range(1,7)) + list(range(8,14)):
        ws.column_dimensions[get_column_letter(c)].width = 10.8
    ws.column_dimensions["G"].width  = 1.2
    ws.column_dimensions["N"].width  = 1.2
    for c in range(15, 27):
        ws.column_dimensions[get_column_letter(c)].width = 8.2

    # ── Row heights ───────────────────────────────────────────────────────────
    LINES     = 10
    SUN_LINES = 12
    LINE_H    = 16.0
    DAY_HDR_H = 16
    GAP_H     = 5

    PAIR_BLOCK = 1 + LINES + 1   # hdr + lines + gap = 12
    SUN_BLOCK  = 1 + SUN_LINES   # hdr + lines = 13

    for r in range(1, 58): ws.row_dimensions[r].height = LINE_H
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 7

    pair_starts = [4, 4 + PAIR_BLOCK, 4 + PAIR_BLOCK*2]
    sun_start   = 4 + PAIR_BLOCK*3

    for rs in pair_starts:
        ws.row_dimensions[rs].height = DAY_HDR_H
        for ln in range(1, LINES+1):
            ws.row_dimensions[rs+ln].height = LINE_H
        ws.row_dimensions[rs+LINES+1].height = GAP_H

    ws.row_dimensions[sun_start].height = DAY_HDR_H
    for ln in range(1, SUN_LINES+1):
        ws.row_dimensions[sun_start+ln].height = LINE_H

    # ── Backgrounds ───────────────────────────────────────────────────────────
    bg(ws, 1, 1,  57, 13, CREAM)
    bg(ws, 1, 14, 57, 14, WARM_MED)
    bg(ws, 1, 15, 57, 26, CREAM)

    # ── LEFT: Header ──────────────────────────────────────────────────────────
    bg(ws, 1, 1, 2, 13, WARM_LIGHT)

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

    # ── LEFT: Day blocks ──────────────────────────────────────────────────────
    def draw_day(day_idx, rs, c_start, c_end):
        d      = week_dates[day_idx]
        is_wkd = d.weekday() >= 5

        ws.cell(rs, c_start).value = d.day
        ws.cell(rs, c_start).font  = mf(9, bold=True,
                                         color=ACCENT if is_wkd else TEXT_DARK)
        ws.cell(rs, c_start).alignment = al("left", "center")

        cell = mc(ws, rs, c_start+1, rs, c_end)
        cell.value = DF[day_idx].upper()
        cell.font  = mf(7, color=TEXT_MED)
        cell.alignment = al("left", "center")

        for col in range(c_start, c_end+1):
            ws.cell(rs, col).border = bdr(bottom=sd("thin", WARM_MED))

        n = SUN_LINES if day_idx == 6 else LINES
        for ln in range(1, n+1):
            dot_line(ws, rs+ln, c_start, c_end)

    pairs     = [(0,1),(2,3),(4,5)]
    for (li, ri), rs in zip(pairs, pair_starts):
        draw_day(li, rs, 1,  6)
        draw_day(ri, rs, 8, 13)

    draw_day(6, sun_start, 1, 13)

    # ── RIGHT: Header ─────────────────────────────────────────────────────────
    bg(ws, 1, 15, 2, 26, WARM_LIGHT)

    # ── RIGHT: PRIORITIES ─────────────────────────────────────────────────────
    PRI_START = 4
    ws.row_dimensions[PRI_START].height = 15
    c = mc(ws, PRI_START, 15, PRI_START, 26)
    c.value = "PRIORITIES"; c.font = mf(7, bold=True, color=ACCENT)
    c.alignment = al("left", "center")
    bg(ws, PRI_START, 15, PRI_START, 26, ACCENT_LIGHT)

    PRI_LINES = 10
    for pi in range(PRI_LINES):
        pr = PRI_START + 1 + pi
        ws.row_dimensions[pr].height = LINE_H
        ws.cell(pr, 15).value     = "○"
        ws.cell(pr, 15).font      = Font(name="Montserrat", size=8, color=ACCENT)
        ws.cell(pr, 15).alignment = al("center")
        dot_line(ws, pr, 16, 26)

    # ── RIGHT: TO-DO ──────────────────────────────────────────────────────────
    TODO_START = PRI_START + PRI_LINES + 2
    ws.row_dimensions[TODO_START].height = 15
    c = mc(ws, TODO_START, 15, TODO_START, 26)
    c.value = "TO-DO"; c.font = mf(7, bold=True, color=TEXT_MED)
    c.alignment = al("left", "center")
    bg(ws, TODO_START, 15, TODO_START, 26, WARM_LIGHT)

    TODO_LINES = 16
    for ti in range(TODO_LINES):
        tr = TODO_START + 1 + ti
        ws.row_dimensions[tr].height = LINE_H
        ws.cell(tr, 15).value     = "○"
        ws.cell(tr, 15).font      = Font(name="Montserrat", size=8, color=ACCENT)
        ws.cell(tr, 15).alignment = al("center")
        dot_line(ws, tr, 16, 26)

    # ── RIGHT: NOTES ──────────────────────────────────────────────────────────
    NOTES_START = TODO_START + TODO_LINES + 2
    ws.row_dimensions[NOTES_START].height = 15
    c = mc(ws, NOTES_START, 15, NOTES_START, 26)
    c.value = "NOTES"; c.font = mf(7, bold=True, color=ACCENT)
    c.alignment = al("left", "center")
    bg(ws, NOTES_START, 15, NOTES_START, 26, ACCENT_LIGHT)

    for nr in range(NOTES_START+1, 58):
        ws.row_dimensions[nr].height = LINE_H
        dot_line(ws, nr, 15, 26)

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
    TOTAL_COLS = 22
    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 6.0
    for r in range(1, 55): ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 55, TOTAL_COLS, CREAM)
    bg(ws, 1, 1, 4,  TOTAL_COLS, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, TOTAL_COLS)
    c.value = f"{goal_type} Goals"; c.font = tf(26, bold=True); c.alignment = C

    # ── Section 1: Big Goals
    S1 = 6
    ws.row_dimensions[S1].height = 22   # taller to fit full section title text
    c = mc(ws, S1, 2, S1, TOTAL_COLS)
    c.value = GOAL_TITLES[goal_type]
    c.font  = tf(13, bold=True, color=ACCENT); c.alignment = al("left","center")
    bg(ws, S1, 2, S1, TOTAL_COLS, ACCENT_LIGHT)

    for i, lr in enumerate(range(S1+2, S1+14)):
        ws.row_dimensions[lr].height = 18
        ws.cell(lr, 2).value = f"{i+1}."
        ws.cell(lr, 2).font  = mf(9, color=ACCENT)
        dot_line(ws, lr, 3, TOTAL_COLS)

    # ── Section 2: Quarter by Quarter
    S2 = S1 + 17
    ws.row_dimensions[S2].height = 22
    c = mc(ws, S2, 2, S2, TOTAL_COLS)
    c.value = "Quarter by Quarter"
    c.font  = tf(13, bold=True, color=ACCENT); c.alignment = al("left","center")
    bg(ws, S2, 2, S2, TOTAL_COLS, ACCENT_LIGHT)

    q_labels = ["Q1  Jan–Mar","Q2  Apr–Jun","Q3  Jul–Sep","Q4  Oct–Dec"]
    q_cols   = [2, 7, 12, 17]
    for q_lbl, qc in zip(q_labels, q_cols):
        ws.row_dimensions[S2+2].height = 16
        c = mc(ws, S2+2, qc, S2+2, qc+3)
        c.value = q_lbl; c.font = mf(8, bold=True, color=TEXT_MED); c.alignment = L
        for lr in range(S2+3, S2+11):
            ws.row_dimensions[lr].height = 16
            dot_line(ws, lr, qc, qc+3)

    # ── Section 3: Action Steps
    S3 = S2 + 14
    ws.row_dimensions[S3].height = 22
    c = mc(ws, S3, 2, S3, TOTAL_COLS)
    c.value = "Action Steps & Milestones"
    c.font  = tf(13, bold=True, color=ACCENT); c.alignment = al("left","center")
    bg(ws, S3, 2, S3, TOTAL_COLS, ACCENT_LIGHT)

    for lr in range(S3+2, min(S3+12, 55)):
        ws.row_dimensions[lr].height = 18
        dot_line(ws, lr, 2, TOTAL_COLS)

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# BRAIN MAP  (flagged for redesign)
# ══════════════════════════════════════════════════════════════════════════════
def make_brain_map(wb):
    ws = wb.create_sheet("Brain Map")
    landscape_all(ws)
    TOTAL_COLS = 28
    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 4.5
    for r in range(1, 50): ws.row_dimensions[r].height = 13
    bg(ws, 1, 1, 50, TOTAL_COLS, CREAM)
    bg(ws, 1, 1, 3,  TOTAL_COLS, WARM_LIGHT)

    c = mc(ws, 2, 1, 2, TOTAL_COLS)
    c.value = "Brain Map"; c.font = tf(22, bold=True); c.alignment = C
    c = mc(ws, 3, 1, 3, TOTAL_COLS)
    c.value = "Place your central idea in the center, then branch outward."
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    cr, cc = 26, 13
    bg(ws, cr-2, cc-1, cr+2, cc+3, ACCENT_LIGHT)
    cell = mc(ws, cr-2, cc-1, cr+2, cc+3)
    cell.value = "Central\nIdea"; cell.font = tf(12, bold=True, color=ACCENT)
    cell.alignment = al("center")
    bdr_range(ws, cr-2, cc-1, cr+2, cc+3,
              bdr(left=sd("medium",ACCENT),right=sd("medium",ACCENT),
                  top=sd("medium",ACCENT),bottom=sd("medium",ACCENT)))

    branch_boxes = [(8,3,6),(8,12,15),(8,21,24),(26,2,5),(26,18,21),
                    (40,3,6),(40,12,15),(40,21,24)]
    for (br_r, bc1, bc2) in branch_boxes:
        bg(ws, br_r, bc1, br_r+3, bc2, WARM_LIGHT)
        cell = mc(ws, br_r, bc1, br_r+3, bc2)
        bdr_range(ws, br_r, bc1, br_r+3, bc2,
                  bdr(left=sd(),right=sd(),top=sd(),bottom=sd()))
        for sub_r in range(br_r, br_r+4):
            for sub_c in range(max(1,bc1-3), bc1):
                ws.cell(sub_r, sub_c).border = bdr(bottom=sd("hair",WARM_MED))
            for sub_c in range(bc2+1, min(TOTAL_COLS,bc2+4)):
                ws.cell(sub_r, sub_c).border = bdr(bottom=sd("hair",WARM_MED))
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# BRAINSTORM  (flagged for redesign)
# ══════════════════════════════════════════════════════════════════════════════
def make_brainstorm(wb):
    ws = wb.create_sheet("Brainstorm")
    landscape_all(ws)
    TOTAL_COLS = 28
    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 4.5
    for r in range(1, 50): ws.row_dimensions[r].height = 13
    bg(ws, 1, 1, 50, TOTAL_COLS, CREAM)
    bg(ws, 1, 1, 3,  TOTAL_COLS, WARM_LIGHT)

    c = mc(ws, 2, 1, 2, TOTAL_COLS)
    c.value = "Brainstorm"; c.font = tf(22, bold=True); c.alignment = C
    c = mc(ws, 3, 1, 3, TOTAL_COLS)
    c.value = "Start with your main idea on the left. Let every branch spark the next."
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    spine_row = 26
    for col in range(2, TOTAL_COLS):
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
    TOTAL_COLS = 22
    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 6.2
    for r in range(1, 50): ws.row_dimensions[r].height = 13
    bg(ws, 1, 1, 50, TOTAL_COLS, CREAM)
    bg(ws, 1, 1, 4,  TOTAL_COLS, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, TOTAL_COLS)
    c.value = "Vision Board"; c.font = tf(26, bold=True); c.alignment = C
    c = mc(ws, 4, 1, 4, TOTAL_COLS)
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
        (14,12, 30, TOTAL_COLS, "paste an image or write your dream here"),
        (32, 1, 42, 11, "this year I will..."),
        (32,12, 42, TOTAL_COLS, "I am grateful for..."),
        (44, 1, 49, TOTAL_COLS, "my word(s) for this season:"),
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
    TOTAL_COLS = 22
    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 6.3
    bg(ws, 1, 1, 55, TOTAL_COLS, CREAM)
    bg(ws, 1, 1, 4,  TOTAL_COLS, WARM_LIGHT)

    for r in range(1, 55): ws.row_dimensions[r].height = 14

    c = mc(ws, 2, 1, 3, TOTAL_COLS)
    c.value = "Bucket List"; c.font = tf(24, bold=True); c.alignment = C
    c = mc(ws, 4, 1, 4, TOTAL_COLS)
    c.value = "Places to go  ·  things to try  ·  memories to make"
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    # 4 categories, each with header + 9 item rows, spaced evenly
    cats = [
        ("Travel & Adventures",  6),
        ("Family Experiences",   18),
        ("Personal Growth",      30),
        ("Just for Fun",         42),
    ]
    for cat_name, sr in cats:
        ws.row_dimensions[sr].height = 22   # tall enough for header text
        c = mc(ws, sr, 2, sr, TOTAL_COLS)
        c.value = cat_name
        c.font  = tf(12, bold=True, color=ACCENT)
        c.alignment = al("left","center")
        bg(ws, sr, 2, sr, TOTAL_COLS, ACCENT_LIGHT)

        for lr in range(sr+2, sr+10):
            ws.row_dimensions[lr].height = 18
            ws.cell(lr, 2).value = "○"
            ws.cell(lr, 2).font  = mf(9, color=ACCENT)
            ws.cell(lr, 2).alignment = al("center")
            dot_line(ws, lr, 3, TOTAL_COLS)

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# SELF-CARE TRACKER
# ══════════════════════════════════════════════════════════════════════════════
def make_self_care_tracker(wb):
    ws = wb.create_sheet("Self-Care Tracker")
    landscape_all(ws, scale=80)

    # Wide habit name col + 31 day cols + notes col
    ws.column_dimensions["A"].width = 22
    for c in range(2, 33):
        ws.column_dimensions[get_column_letter(c)].width = 3.4
    ws.column_dimensions[get_column_letter(33)].width = 12

    for r in range(1, 36): ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 36, 33, CREAM)

    # Title header
    bg(ws, 1, 1, 4, 33, WARM_LIGHT)
    ws.row_dimensions[1].height = 8   # top padding
    ws.row_dimensions[2].height = 24  # title row — tall enough
    ws.row_dimensions[3].height = 16  # subtitle
    ws.row_dimensions[4].height = 8   # bottom padding

    c = mc(ws, 2, 1, 2, 33)
    c.value = "Self-Care Tracker"; c.font = tf(22, bold=True); c.alignment = C
    c = mc(ws, 3, 1, 3, 33)
    c.value = "You can't pour from an empty cup. Track the habits that fill yours."
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    # Column headers row
    ws.row_dimensions[6].height = 16
    ws.cell(6, 1).value = "Habit"
    ws.cell(6, 1).font  = mf(8, bold=True, color=TEXT_MED)
    for d in range(1, 32):
        cell = ws.cell(6, d+1)
        cell.value = d; cell.font = mf(7, bold=True, color=TEXT_MED)
        cell.alignment = al("center")
    ws.cell(6, 33).value = "Notes"
    ws.cell(6, 33).font  = mf(8, bold=True, color=TEXT_MED)
    bg(ws, 6, 1, 6, 33, WARM_LIGHT)

    # Pre-filled habits (universal mom habits) + blank rows
    pre_filled = [
        "Quiet time / prayer",
        "Time in the Bible",
        "Drink enough water",
        "Healthy food choices",
        "Intentional time with each child",
        "Move my body",
        "Good night's sleep",
    ]
    blank_rows = 5  # user fills these in

    all_habits = pre_filled + [""] * blank_rows  # 12 total

    for i, habit in enumerate(all_habits):
        r = 8 + i*2
        ws.row_dimensions[r].height = 18
        ws.cell(r, 1).value = habit
        ws.cell(r, 1).font  = mf(8, color=TEXT_DARK if habit else TEXT_LIGHT)
        if not habit:
            ws.cell(r, 1).value = "________________"
            ws.cell(r, 1).font  = mf(8, color=WARM_DARK)
        bg(ws, r, 1, r, 1, WARM_LIGHT)

        for d in range(1, 32):
            cell = ws.cell(r, d+1)
            cell.value = "○"
            cell.font  = Font(name="Montserrat", size=7, color=WARM_DARK)
            cell.alignment = al("center")
            cell.border = bdr(left=sd("hair",WARM_MED), bottom=sd("dotted",WARM_DARK))
        ws.cell(r, 33).border = bdr(bottom=sd("dotted", WARM_DARK))
        ws.row_dimensions[r+1].height = 3

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# NEXT YEAR GOALS & CONSIDERATIONS  (last tab)
# ══════════════════════════════════════════════════════════════════════════════
def make_next_year_page(wb):
    ws = wb.create_sheet("Looking Ahead · 2028")
    landscape_all(ws, scale=80)

    TOTAL_COLS = 30

    # Left section (cols 1-18): 2028 mini calendar preview
    # Right section (cols 20-30): goals + notes
    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 4.5
    ws.column_dimensions[get_column_letter(19)].width = 1.5  # divider

    for r in range(1, 55): ws.row_dimensions[r].height = 13
    bg(ws, 1, 1, 55, TOTAL_COLS, CREAM)
    bg(ws, 1, 1, 4,  TOTAL_COLS, WARM_LIGHT)

    # ── Title
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 14
    c = mc(ws, 2, 1, 2, TOTAL_COLS)
    c.value = "Looking Ahead · 2028"; c.font = tf(26, bold=True); c.alignment = C

    c = mc(ws, 3, 1, 3, TOTAL_COLS)
    c.value = "Dream it. Plan it. Live it."
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    for col in range(1, TOTAL_COLS+1):
        ws.cell(4, col).border = bdr(bottom=sd("medium", ACCENT))

    # ── Left: 2028 year calendar (4 cols × 3 rows = 12 months)
    # Mini-cal layout: 4 months per row, 3 rows
    # Each mini-cal: 7 day-cols + 1 gap = 8 col units; 4 minis × but we have 18 cols
    # Use 4 cols wide per mini-cal: month name (merged 4) + 4 day cols (S-S abbreviated 2-per-col)
    # Simpler: use 3 cals per row × 2 rows = 6 cols done differently
    # Let's do 3 per row × 4 rows = 12 months using cols 1-18, 6 cols per mini-cal

    CAL_W  = 4   # cols per mini-calendar
    GAP_C  = 1   # gap col between cals per row
    CALS_PER_ROW = 4
    # col starts for 4 minis: 1, 6, 11, 16
    cal_col_starts = [1, 1+(CAL_W+GAP_C), 1+2*(CAL_W+GAP_C), 1+3*(CAL_W+GAP_C)]
    # = 1, 6, 11, 16

    ROWS_PER_CAL = 9  # name + day-hdr + 6 weeks + gap
    cal_row_starts = [6, 6+ROWS_PER_CAL, 6+2*ROWS_PER_CAL]  # 3 rows × 4 cols = 12 months

    for mi, mo in enumerate(range(1, 13)):
        grow = mi // CALS_PER_ROW
        gcol = mi  % CALS_PER_ROW
        cr   = cal_row_starts[grow]
        cc   = cal_col_starts[gcol]

        # Month name
        ws.row_dimensions[cr].height = 14
        c = mc(ws, cr, cc, cr, cc+CAL_W-1)
        c.value = MA[mo]; c.font = tf(8, bold=True, color=TEXT_MED)
        c.alignment = al("center")
        bg(ws, cr, cc, cr, cc+CAL_W-1, WARM_LIGHT)

        # Day letters (S M T W T F S) compressed into 4 cols — use first letter only
        ws.row_dimensions[cr+1].height = 10
        for d, ltr in enumerate(["S","M","T","W","T","F","S"]):
            col_offset = d % CAL_W
            cell = ws.cell(cr+1, cc+col_offset)
            # We'll just show dates — skip full day header in 4-col layout

        # Dates (Sun-first, compressed)
        for wi, week in enumerate(calendar.monthcalendar(2028, mo)):
            sun = week[6]; row_days = [sun] + week[:6]
            ws.row_dimensions[cr+2+wi].height = 11
            for d, day_num in enumerate(row_days):
                if day_num:
                    col_pos = cc + (d % CAL_W)
                    cell = ws.cell(cr+2+wi, col_pos)
                    if not cell.value:
                        cell.value = day_num
                    cell.font  = mf(6, color=TEXT_MED)
                    cell.alignment = al("center")

    # ── Divider strip
    bg(ws, 5, 19, 55, 19, WARM_MED)

    # ── Right: Goals + Notes (cols 20-30)
    RS = 20  # right start col
    RE = TOTAL_COLS  # right end col

    sections = [
        (6,  "Goals for 2028",           16),
        (18, "Intentions & Hopes",       28),
        (30, "Word(s) for the Year",     36),
        (38, "Things to Start / Stop",   48),
    ]
    for sr, slbl, er in sections:
        ws.row_dimensions[sr].height = 20
        c = mc(ws, sr, RS, sr, RE)
        c.value = slbl; c.font = tf(11, bold=True, color=ACCENT); c.alignment = L
        bg(ws, sr, RS, sr, RE, ACCENT_LIGHT)
        for lr in range(sr+2, er):
            ws.row_dimensions[lr].height = 16
            dot_line(ws, lr, RS, RE)

    # ── Branding at bottom
    ws.row_dimensions[51].height = 18
    ws.row_dimensions[52].height = 16
    _add_branding(ws, 51, 2, TOTAL_COLS-1)

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

    print("Year at a Glance...")
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

    print("Next year page...")
    make_next_year_page(wb)

    out = "/home/user/momlifeplanner/MomLifePlanner.xlsx"
    wb.save(out)
    print(f"\nSaved → {out}  ({len(wb.sheetnames)} sheets)")

if __name__ == "__main__":
    build_planner()
