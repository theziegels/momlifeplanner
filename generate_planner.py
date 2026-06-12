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
from openpyxl.drawing.image import Image as XLImage

# ── Brand Colors ──────────────────────────────────────────────────────────────
CREAM        = "FDFCFB"
WARM_LIGHT   = "F6F0EB"
WARM_MED     = "EAE0D9"
WARM_DARK    = "CBBFB8"
ACCENT       = "CE8282"
ACCENT_LIGHT = "F4E4E4"
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

def landscape_all(ws, scale=80, rows=None, cols=None):
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
    if rows and cols:
        ws.print_area = f"A1:{get_column_letter(cols)}{rows}"

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

HOLIDAYS = {
    (2026,  1,  1): "New Year's Day",
    (2026,  4,  3): "Good Friday",
    (2026,  4,  5): "Easter",
    (2026,  5, 10): "Mother's Day",
    (2026,  5, 25): "Memorial Day",
    (2026,  6, 21): "Father's Day",
    (2026,  7,  4): "Independence Day",
    (2026,  9,  7): "Labor Day",
    (2026, 11, 26): "Thanksgiving",
    (2026, 12, 24): "Christmas Eve",
    (2026, 12, 25): "Christmas Day",
    (2026, 12, 31): "New Year's Eve",
    (2027,  1,  1): "New Year's Day",
    (2027,  3, 26): "Good Friday",
    (2027,  3, 28): "Easter",
    (2027,  5,  9): "Mother's Day",
    (2027,  5, 31): "Memorial Day",
    (2027,  6, 20): "Father's Day",
    (2027,  7,  4): "Independence Day",
    (2027,  9,  6): "Labor Day",
    (2027, 11, 25): "Thanksgiving",
    (2027, 12, 24): "Christmas Eve",
    (2027, 12, 25): "Christmas Day",
    (2027, 12, 31): "New Year's Eve",
    (2028,  1,  1): "New Year's Day",
}

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
    TOTAL_COLS = 20
    landscape_all(ws, rows=56, cols=TOTAL_COLS)
    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 11.4
    for r in range(1, 57):
        ws.row_dimensions[r].height = 14

    bg(ws, 1, 1, 56, TOTAL_COLS, CREAM)

    # Border strips — all WARM_MED, all four sides
    STRIP = 1
    bg(ws, 1,  1,    3,    TOTAL_COLS, WARM_MED)   # top bar
    bg(ws, 53, 1,    56,   TOTAL_COLS, WARM_MED)   # bottom bar
    bg(ws, 1,  1,    56,   STRIP,      WARM_MED)   # left strip
    bg(ws, 1,  TOTAL_COLS, 56, TOTAL_COLS, WARM_MED)  # right strip

    # Thin accent rule just inside the top bar
    for col in range(2, TOTAL_COLS):
        ws.cell(4, col).border = bdr(bottom=sd("thin", ACCENT))

    # ── Title block — centered at ~42% down the page
    ws.row_dimensions[20].height = 65   # Mom Life (top)
    ws.row_dimensions[21].height = 65   # Mom Life (bottom)
    ws.row_dimensions[22].height = 3    # gap
    ws.row_dimensions[23].height = 22   # PLANNER
    ws.row_dimensions[24].height = 24   # accent rule under title

    c = mc(ws, 20, 2, 21, TOTAL_COLS-1)
    c.value = "Mom Life"; c.font = tf(66, bold=True, color=TEXT_DARK)
    c.alignment = al("center")

    c = mc(ws, 23, 2, 23, TOTAL_COLS-1)
    c.value = "P  L  A  N  N  E  R"; c.font = mf(18, color=ACCENT)
    c.alignment = al("center")

    # Thin accent rule under title
    for col in range(4, TOTAL_COLS-2):
        ws.cell(24, col).border = bdr(bottom=sd("thin", ACCENT))

    c = mc(ws, 27, 2, 27, TOTAL_COLS-1)
    c.value = "August 2026 – December 2027"
    c.font = mf(10, italic=True, color=TEXT_MED); c.alignment = al("center")

    # ── Scripture quote
    ws.row_dimensions[33].height = 18
    ws.row_dimensions[34].height = 18
    c = mc(ws, 33, 3, 34, TOTAL_COLS-2)
    c.value = '"You are worried and distracted by many things —\nthere is need of only one thing."  — Luke 10:41–42'
    c.font  = tf(14, italic=True, color=TEXT_MED); c.alignment = al("center")

    # ── Tagline
    ws.row_dimensions[42].height = 16
    c = mc(ws, 42, 2, 42, TOTAL_COLS-1)
    c.value = "dream  ·  plan  ·  do  ·  thrive"
    c.font  = mf(9, italic=True, color=TEXT_LIGHT); c.alignment = al("center")

    # ── Branding
    ws.row_dimensions[48].height = 18
    _add_branding(ws, 48, 2, TOTAL_COLS-1)

    # ── Thin accent rule just inside the bottom bar (mirrors top)
    for col in range(2, TOTAL_COLS):
        ws.cell(52, col).border = bdr(bottom=sd("thin", ACCENT))

    return ws


def _add_branding(ws, r, c1, c2):
    """Single merged cell — guaranteed one line."""
    cell = mc(ws, r, c1, r, c2)
    cell.value = "Beloved Kingdom Collective™  ·  Mom Life Planner"
    cell.font  = tf(10, bold=True, italic=True, color=TEXT_MED)
    cell.alignment = al("center")


# ══════════════════════════════════════════════════════════════════════════════
# HOW TO USE
# ══════════════════════════════════════════════════════════════════════════════
def make_how_to_use(wb):
    ws = wb.create_sheet("How To Use")
    TOTAL_COLS = 22
    landscape_all(ws, rows=56, cols=TOTAL_COLS)
    # Col 1: narrow left margin; col 2: thin accent strip; cols 3+: content
    ws.column_dimensions["A"].width = 1.5
    ws.column_dimensions["B"].width = 1.0   # narrow accent strip
    for c in range(3, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 10.0

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
        (30, "Brain Map",
             "Central idea branching outward in all directions — "
             "great for free-flow thinking and planning."),
        (36, "Vision Board + Extras",
             "Vision Board, Bucket List, and Important Contacts — "
             "because a well-planned life is a well-lived life."),
    ]

    for row, title, body in items:
        # Accent strip spans title row and body rows only (row to row+2)
        bg(ws, row, 2, row+2, 2, ACCENT)
        ws.row_dimensions[row].height   = 20
        ws.row_dimensions[row+1].height = 14
        ws.row_dimensions[row+2].height = 14
        ws.row_dimensions[row+3].height = 8   # gap after section

        c = mc(ws, row, 3, row, TOTAL_COLS)
        c.value = title; c.font = tf(14, bold=True, color=ACCENT); c.alignment = L

        c = mc(ws, row+1, 3, row+2, TOTAL_COLS)
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
    landscape_all(ws, rows=54, cols=22)

    # Col A: category label (wide); cols 2-5: Name; 6: gap; 7-10: Phone;
    # 11: gap; 12-16: Email; 17: gap; 18-22: Notes
    ws.column_dimensions["A"].width = 24
    for c in range(2, 6):   ws.column_dimensions[get_column_letter(c)].width = 11
    ws.column_dimensions["F"].width = 2.0
    for c in range(7, 11):  ws.column_dimensions[get_column_letter(c)].width = 11
    ws.column_dimensions["K"].width = 2.0
    for c in range(12, 17): ws.column_dimensions[get_column_letter(c)].width = 9.0
    ws.column_dimensions["Q"].width = 2.5
    for c in range(18, 23): ws.column_dimensions[get_column_letter(c)].width = 11

    for r in range(1, 55): ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 54, 22, CREAM)
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
            cell.border = bdr(bottom=sd("dotted", WARM_DARK))
        dot_line(ws, r+1, 2, 22)
        r += 2

    # Additional contacts header
    ws.row_dimensions[r].height = 16
    hcell = mc(ws, r, 1, r, 22)
    hcell.value = "Additional Contacts"
    hcell.font = mf(9, bold=True, color=ACCENT); hcell.alignment = al("left","center")
    bg(ws, r, 1, r, 22, WARM_LIGHT)
    r += 1

    # blank write-in rows — fill remaining page for any other contacts
    for wr in range(r, 54):
        ws.row_dimensions[wr].height = 20
        dot_line(ws, wr, 1, 22)

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# YEAR AT A GLANCE  –  5 × 4 = 20 months, landscape
# ══════════════════════════════════════════════════════════════════════════════
def make_year_view(wb):
    ws = wb.create_sheet("Year at a Glance")
    landscape_all(ws, scale=80, rows=41, cols=39)

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
    ws.row_dimensions[3].height = 14
    for g in range(4):
        ws.row_dimensions[3 + (g+1)*ROWS_PER].height = 10

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
    # 7 day cols + divider + 3 notes cols — widened to fill landscape page
    landscape_all(ws, scale=80, rows=52, cols=10)

    for d in range(1, 8):  ws.column_dimensions[get_column_letter(d)].width = 17.0
    ws.column_dimensions["H"].width = 0.8
    ws.column_dimensions["I"].width = 26
    ws.column_dimensions["J"].width = 26

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
                hkey = (year, month, day_num)
                if hkey in HOLIDAYS and LINES >= 1:
                    hcell = ws.cell(row+1, col)
                    hcell.value = HOLIDAYS[hkey]
                    hcell.font  = mf(6, italic=True, color=TEXT_LIGHT)
                    hcell.alignment = al("center", "center")
            else:
                for lr in range(row, row+BLOCK):
                    bg(ws, lr, col, lr, col, WARM_LIGHT)
        row += BLOCK

    # ── Bottom Notes section (below calendar, cols 1-7)
    if row <= 48:
        ws.row_dimensions[row].height = 16
        c = mc(ws, row, 1, row, 7)
        c.value = "Notes"; c.font = mf(8, bold=True, color=ACCENT)
        c.alignment = L
        bg(ws, row, 1, row, 7, ACCENT_LIGHT)
        for nr in range(row+1, 51):
            ws.row_dimensions[nr].height = 14
            dot_line(ws, nr, 1, 7)

    # Notes sidebar — "Don't Forget" and "Meal Ideas" swapped
    c = mc(ws, 1, 9, 2, 10)
    c.value = "Monthly Notes"; c.font = tf(14, bold=True); c.alignment = L
    bg(ws, 1, 9, 4, 10, WARM_LIGHT)

    note_secs = [(6,"Appointments"),(14,"Birthdays & Occasions"),
                 (22,"Don't Forget"),(30,"Meal Ideas")]
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
    landscape_all(ws, scale=80, rows=57, cols=26)

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
    ws.row_dimensions[3].height = 16

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

    c = mc(ws, 1, 9, 2, 11)
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
        hkey = (week_dates[day_idx].year, week_dates[day_idx].month, week_dates[day_idx].day)
        if hkey in HOLIDAYS:
            cell.value = DF[day_idx].upper() + "\n" + HOLIDAYS[hkey]
            cell.font  = mf(7, color=TEXT_MED)
            ws.row_dimensions[rs].height = 22
        else:
            cell.value = DF[day_idx].upper()
            cell.font  = mf(7, color=TEXT_MED)
        cell.alignment = al("left", "center", wrap=True)

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
    TOTAL_COLS = 22
    landscape_all(ws, rows=55, cols=TOTAL_COLS)
    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 11.5
    for r in range(1, 55): ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 55, TOTAL_COLS, CREAM)
    bg(ws, 1, 1, 4,  TOTAL_COLS, WARM_LIGHT)

    c = mc(ws, 2, 1, 3, TOTAL_COLS)
    c.value = f"{goal_type} Goals"; c.font = tf(26, bold=True); c.alignment = C

    # ── Section 1: Big Goals
    S1 = 6
    ws.row_dimensions[S1].height = 22   # taller to fit full section title text
    c = mc(ws, S1, 1, S1, TOTAL_COLS)
    c.value = GOAL_TITLES[goal_type]
    c.font  = tf(13, bold=True, color=ACCENT); c.alignment = al("left","center")
    bg(ws, S1, 1, S1, TOTAL_COLS, ACCENT_LIGHT)

    for i, lr in enumerate(range(S1+2, S1+14)):
        ws.row_dimensions[lr].height = 18
        ws.cell(lr, 1).value = f"{i+1}."
        ws.cell(lr, 1).font  = mf(9, color=ACCENT)
        dot_line(ws, lr, 2, TOTAL_COLS)

    # ── Section 2: Quarter by Quarter
    S2 = S1 + 17
    ws.row_dimensions[S2].height = 22
    c = mc(ws, S2, 1, S2, TOTAL_COLS)
    c.value = "Quarter by Quarter"
    c.font  = tf(13, bold=True, color=ACCENT); c.alignment = al("left","center")
    bg(ws, S2, 1, S2, TOTAL_COLS, ACCENT_LIGHT)

    QW = TOTAL_COLS // 4  # = 5
    q_labels = ["Q1  Jan–Mar","Q2  Apr–Jun","Q3  Jul–Sep","Q4  Oct–Dec"]
    q_starts = [1 + i*QW for i in range(4)]
    for q_lbl, qc in zip(q_labels, q_starts):
        ws.row_dimensions[S2+2].height = 16
        c = mc(ws, S2+2, qc, S2+2, qc+QW-1)
        c.value = q_lbl; c.font = mf(8, bold=True, color=TEXT_MED); c.alignment = al("center")
        for lr in range(S2+3, S2+11):
            ws.row_dimensions[lr].height = 16
            dot_line(ws, lr, qc, qc+QW-1)

    # ── Section 3: Action Steps
    S3 = S2 + 14
    ws.row_dimensions[S3].height = 22
    c = mc(ws, S3, 1, S3, TOTAL_COLS)
    c.value = "Action Steps & Milestones"
    c.font  = tf(13, bold=True, color=ACCENT); c.alignment = al("left","center")
    bg(ws, S3, 1, S3, TOTAL_COLS, ACCENT_LIGHT)

    for lr in range(S3+2, min(S3+12, 55)):
        ws.row_dimensions[lr].height = 18
        ws.cell(lr, 1).value = "○"
        ws.cell(lr, 1).font  = mf(9, color=ACCENT)
        ws.cell(lr, 1).alignment = al("center")
        dot_line(ws, lr, 2, TOTAL_COLS)

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# BRAIN MAP  –  matplotlib image embedded in worksheet
# ══════════════════════════════════════════════════════════════════════════════
def _generate_brain_map_image(path):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm
    from matplotlib.patches import FancyBboxPatch

    font_path = '/tmp/CormorantGaramond-Bold.ttf'
    try:
        fm.fontManager.addfont(font_path)
        cg_prop = fm.FontProperties(fname=font_path)
    except Exception:
        cg_prop = fm.FontProperties(family='serif')

    WARM_LIGHT_M = "#F6F0EB"; WARM_DARK_M = "#CBBFB8"
    BRANCH_COLORS = ["#D4B5B5","#C8C2BA","#B8C5C8","#C4C8B5","#C8B8C4"]

    fig, ax = plt.subplots(figsize=(16, 11))
    ax.set_xlim(0,16); ax.set_ylim(0,11); ax.set_aspect('equal'); ax.axis('off')
    fig.patch.set_facecolor("#"+CREAM); ax.set_facecolor("#"+CREAM)

    cx,cy=8.0,5.3; branch_r=0.95; nw,nh=2.2,0.40; spacing=0.62
    center_hw,center_hh=1.35,0.6

    branch_data=[
        (90,  2.5, 10.3,  9.2,  BRANCH_COLORS[0]),
        (18,  3.5, 13.76, 6.63, BRANCH_COLORS[1]),
        (306, 2.7, 12.6,  2.7,  BRANCH_COLORS[2]),
        (234, 2.7,  3.4,  2.7,  BRANCH_COLORS[3]),
        (162, 3.5,  2.04, 6.63, BRANCH_COLORS[4]),
    ]

    import numpy as np
    for angle,dist,col_x,col_y,color in branch_data:
        rad=np.radians(angle); bx=cx+np.cos(rad)*dist; by=cy+np.sin(rad)*dist
        ux,uy=np.cos(rad),np.sin(rad)
        ax.plot([cx+ux*(center_hw+0.05), bx-ux*(branch_r+0.05)],
                [cy+uy*(center_hh+0.05), by-uy*(branch_r+0.05)],
                color=WARM_DARK_M,lw=1.4,zorder=1,solid_capstyle='round')
        ax.add_patch(plt.Circle((bx,by),branch_r,color=color,zorder=3,alpha=0.85))
        ax.add_patch(plt.Circle((bx,by),branch_r,fill=False,
                                 edgecolor=WARM_DARK_M,linewidth=0.9,zorder=4))
        dx=col_x-bx; dy=col_y-by; dc=np.hypot(dx,dy); sux,suy=dx/dc,dy/dc
        use_left=sux>0.25; use_right=sux<-0.25
        use_bot=suy>0.25 and abs(sux)<=0.25
        for i in range(3):
            ny=col_y+(i-1)*spacing; nx=col_x
            ax.add_patch(FancyBboxPatch((nx-nw/2,ny-nh/2),nw,nh,
                boxstyle="round,pad=0.08",linewidth=1.2,
                edgecolor=WARM_DARK_M,facecolor=WARM_LIGHT_M,zorder=5))
            sx=bx+sux*branch_r*1.03; sy=by+suy*branch_r*1.03
            if use_left:    ex,ey=nx-nw/2,ny
            elif use_right: ex,ey=nx+nw/2,ny
            elif use_bot:   ex,ey=nx,ny-nh/2
            else:           ex,ey=nx,ny+nh/2
            ax.plot([sx,ex],[sy,ey],color=WARM_DARK_M,lw=0.75,zorder=2)

    ax.add_patch(FancyBboxPatch((cx-1.35,cy-0.6),2.7,1.2,
        boxstyle="round,pad=0.18",linewidth=1.8,
        edgecolor=WARM_DARK_M,facecolor=WARM_LIGHT_M,zorder=6))

    instr=dict(fontsize=8,color="#"+TEXT_LIGHT,fontstyle='italic',
               fontfamily='sans-serif',va='top')
    ax.text(0.3,10.75,"① Write your central topic in the center box.",ha='left',**instr)
    ax.text(0.3,10.43,"② Fill in a branch circle for each key theme.", ha='left',**instr)
    ax.text(0.3,10.11,"③ Add details or actions in the side boxes.",   ha='left',**instr)
    ax.text(13.0,9.4,"Brain Map",ha='left',va='top',
            fontsize=38,color="#"+TEXT_DARK,fontproperties=cg_prop)

    plt.tight_layout(pad=0.2)
    plt.savefig(path,dpi=220,bbox_inches='tight',facecolor="#"+CREAM)
    plt.close(fig)

def make_brain_map(wb):
    import tempfile, os
    ws = wb.create_sheet("Brain Map")
    landscape_all(ws, scale=80, rows=54, cols=42)
    bg(ws, 1, 1, 54, 42, CREAM)
    for c in range(1, 43):
        ws.column_dimensions[get_column_letter(c)].width = 5.8
    for r in range(1, 55):
        ws.row_dimensions[r].height = 13

    img_path = os.path.join(tempfile.gettempdir(), "brain_map_embed.png")
    _generate_brain_map_image(img_path)

    img = XLImage(img_path)
    # Scale to fill the print area (42 cols × 5.8 ≈ 243 units × 0.079 ≈ 19.2" → 1382px at 72dpi)
    img.width  = 1260
    img.height = 866
    img.anchor = "A1"
    ws.add_image(img)
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# VISION BOARD
# ══════════════════════════════════════════════════════════════════════════════
def make_vision_board(wb):
    ws = wb.create_sheet("Vision Board")
    TOTAL_COLS = 22
    landscape_all(ws, rows=50, cols=TOTAL_COLS)
    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 7.5
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
    TOTAL_COLS = 22
    landscape_all(ws, rows=52, cols=TOTAL_COLS)
    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 10.5
    bg(ws, 1, 1, 52, TOTAL_COLS, CREAM)
    bg(ws, 1, 1, 4,  TOTAL_COLS, WARM_LIGHT)

    for r in range(1, 53): ws.row_dimensions[r].height = 14
    ws.row_dimensions[1].height = 6
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[3].height = 14
    ws.row_dimensions[4].height = 8

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
        ws.row_dimensions[sr].height = 22
        c = mc(ws, sr, 1, sr, TOTAL_COLS)
        c.value = cat_name
        c.font  = tf(12, bold=True, color=ACCENT)
        c.alignment = al("left","center")
        bg(ws, sr, 1, sr, TOTAL_COLS, ACCENT_LIGHT)

        for lr in range(sr+1, sr+10):
            ws.row_dimensions[lr].height = 18
            ws.cell(lr, 1).value = "○"
            ws.cell(lr, 1).font  = mf(9, color=ACCENT)
            ws.cell(lr, 1).alignment = al("center")
            dot_line(ws, lr, 2, TOTAL_COLS)

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# SELF-CARE TRACKER
# ══════════════════════════════════════════════════════════════════════════════
def make_self_care_tracker(wb):
    ws = wb.create_sheet("Self-Care Tracker")
    landscape_all(ws, scale=80, rows=36, cols=33)

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
# MONTHLY HABIT TRACKER  –  one per month, inserted after monthly layout
# ══════════════════════════════════════════════════════════════════════════════
def make_monthly_habit_tracker(wb, year, month):
    days_in_month = calendar.monthrange(year, month)[1]
    ws = wb.create_sheet(f"Habits {MA[month]} {year}")

    # cols: A=habit name, B...(days+1)=day circles, last=notes
    TOTAL_COLS = 1 + days_in_month + 1
    landscape_all(ws, scale=80, rows=52, cols=TOTAL_COLS)

    # Dynamic day-col width so all months print at identical scale
    A_W, NOTES_W = 22, 12
    day_w = round((203 - A_W - NOTES_W) / days_in_month, 2)
    ws.column_dimensions["A"].width = A_W
    for c in range(2, days_in_month+2):
        ws.column_dimensions[get_column_letter(c)].width = day_w
    ws.column_dimensions[get_column_letter(TOTAL_COLS)].width = NOTES_W

    for r in range(1, 53): ws.row_dimensions[r].height = 14
    bg(ws, 1, 1, 52, TOTAL_COLS, CREAM)

    # ── Header
    bg(ws, 1, 1, 4, TOTAL_COLS, WARM_LIGHT)
    ws.row_dimensions[1].height = 8
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 8

    c = mc(ws, 2, 1, 2, TOTAL_COLS)
    c.value = f"{MN[month]} {year}"
    c.font  = tf(20, bold=True); c.alignment = C

    c = mc(ws, 3, 1, 3, TOTAL_COLS)
    c.value = "Habit Tracker"
    c.font  = mf(9, italic=True, color=TEXT_MED); c.alignment = C

    # ── Column headers: day numbers
    ws.row_dimensions[6].height = 14
    ws.cell(6, 1).value = "Habit"
    ws.cell(6, 1).font  = mf(8, bold=True, color=TEXT_MED)
    for d in range(1, days_in_month+1):
        cell = ws.cell(6, d+1)
        cell.value = d
        cell.font  = mf(7, bold=True, color=TEXT_MED)
        cell.alignment = al("center")
    ws.cell(6, TOTAL_COLS).value = "Notes"
    ws.cell(6, TOTAL_COLS).font  = mf(8, bold=True, color=TEXT_MED)
    bg(ws, 6, 1, 6, TOTAL_COLS, WARM_LIGHT)

    # ── Habit rows
    pre_filled = [
        "Quiet time / prayer",
        "Time in the Bible",
        "Drink enough water",
        "Healthy food choices",
        "Intentional time with each child",
        "Move my body",
        "Good night's sleep",
    ]
    all_habits = pre_filled + [""] * 5  # 12 total

    LAST_HABIT_ROW = 0
    for i, habit in enumerate(all_habits):
        r = 8 + i * 2
        ws.row_dimensions[r].height = 18
        LAST_HABIT_ROW = r

        if habit:
            # Filled habit: name + circles + dotted borders
            ws.cell(r, 1).value = habit
            ws.cell(r, 1).font  = mf(8, color=TEXT_DARK)
            bg(ws, r, 1, r, 1, WARM_LIGHT)
            for d in range(1, days_in_month+1):
                cell = ws.cell(r, d+1)
                cell.value = "○"
                cell.font  = Font(name="Montserrat", size=7, color=WARM_DARK)
                cell.alignment = al("center")
                cell.border = bdr(left=sd("hair", WARM_MED),
                                  bottom=sd("dotted", WARM_DARK))
            ws.cell(r, TOTAL_COLS).border = bdr(bottom=sd("dotted", WARM_DARK))
        else:
            # Blank row: solid colored block, no circles or lines — visual cue to add habits
            bg(ws, r, 1, r, TOTAL_COLS, WARM_LIGHT)

        ws.row_dimensions[r+1].height = 3

    # ── Notes (left) + Gratitude (right) sections below habits
    SEC_START = LAST_HABIT_ROW + 3
    SEC_END   = 50
    MID_COL   = TOTAL_COLS // 2

    # Notes header
    ws.row_dimensions[SEC_START].height = 16
    c = mc(ws, SEC_START, 1, SEC_START, MID_COL)
    c.value = "Notes"; c.font = mf(8, bold=True, color=TEXT_MED)
    c.alignment = L
    bg(ws, SEC_START, 1, SEC_START, MID_COL, WARM_LIGHT)

    # Gratitude header
    c = mc(ws, SEC_START, MID_COL+1, SEC_START, TOTAL_COLS)
    c.value = "Gratitude"; c.font = mf(8, bold=True, color=ACCENT)
    c.alignment = L
    bg(ws, SEC_START, MID_COL+1, SEC_START, TOTAL_COLS, WARM_LIGHT)

    # Ruled lines for both sections
    for lr in range(SEC_START+1, SEC_END+1):
        ws.row_dimensions[lr].height = 14
        dot_line(ws, lr, 1,          MID_COL)
        dot_line(ws, lr, MID_COL+1,  TOTAL_COLS)

    # Box borders around each section
    b_top    = sd("thin", WARM_MED)
    b_side   = sd("thin", WARM_MED)
    b_bottom = sd("thin", WARM_MED)
    for row in range(SEC_START, SEC_END+1):
        ws.cell(row, 1).border          = bdr(left=b_side,
            top=b_top if row==SEC_START else None,
            bottom=b_bottom if row==SEC_END else None)
        ws.cell(row, MID_COL).border    = bdr(right=b_side,
            top=b_top if row==SEC_START else None,
            bottom=b_bottom if row==SEC_END else None)
        ws.cell(row, MID_COL+1).border  = bdr(left=b_side,
            top=b_top if row==SEC_START else None,
            bottom=b_bottom if row==SEC_END else None)
        ws.cell(row, TOTAL_COLS).border = bdr(right=b_side,
            top=b_top if row==SEC_START else None,
            bottom=b_bottom if row==SEC_END else None)
    # Top and bottom full-width lines
    for c2 in range(1, MID_COL+1):
        ws.cell(SEC_START, c2).border   = bdr(top=b_top,
            left=b_side if c2==1 else None, right=b_side if c2==MID_COL else None)
        ws.cell(SEC_END,   c2).border   = bdr(bottom=b_bottom,
            left=b_side if c2==1 else None, right=b_side if c2==MID_COL else None)
    for c2 in range(MID_COL+1, TOTAL_COLS+1):
        ws.cell(SEC_START, c2).border   = bdr(top=b_top,
            left=b_side if c2==MID_COL+1 else None, right=b_side if c2==TOTAL_COLS else None)
        ws.cell(SEC_END,   c2).border   = bdr(bottom=b_bottom,
            left=b_side if c2==MID_COL+1 else None, right=b_side if c2==TOTAL_COLS else None)

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# NEXT YEAR GOALS & CONSIDERATIONS  (last tab)
# ══════════════════════════════════════════════════════════════════════════════
def make_next_year_page(wb):
    ws = wb.create_sheet("Looking Ahead · 2028")

    # Layout: 3 mini-cals per row × 4 rows = 12 months
    # Each mini-cal: 7 day cols (S M T W T F S) + 1 gap col = 8 units
    # 3 cals = 3×7 + 2 gaps = 23 cal cols (1-23)
    # col 24: divider strip
    # cols 25-42: goals/notes section (18 cols)
    TOTAL_COLS = 42
    landscape_all(ws, scale=80, rows=54, cols=TOTAL_COLS)

    # Column widths
    for c in range(1, 24):       ws.column_dimensions[get_column_letter(c)].width = 4.8
    ws.column_dimensions[get_column_letter(8)].width  = 1.2   # gap between cal 1 & 2
    ws.column_dimensions[get_column_letter(16)].width = 1.2   # gap between cal 2 & 3
    ws.column_dimensions[get_column_letter(24)].width = 1.2   # divider
    for c in range(25, TOTAL_COLS+1): ws.column_dimensions[get_column_letter(c)].width = 5.5

    for r in range(1, 55): ws.row_dimensions[r].height = 13
    bg(ws, 1, 1, 55, TOTAL_COLS, CREAM)
    bg(ws, 1, 1, 4,  TOTAL_COLS, WARM_LIGHT)

    # ── Title
    ws.row_dimensions[1].height = 6
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[3].height = 14
    ws.row_dimensions[4].height = 8
    c = mc(ws, 2, 1, 2, TOTAL_COLS)
    c.value = "Looking Ahead · 2028"; c.font = tf(26, bold=True); c.alignment = C

    c = mc(ws, 3, 1, 3, TOTAL_COLS)
    c.value = "Dream it. Plan it. Live it."
    c.font = mf(8, italic=True, color=TEXT_MED); c.alignment = C

    for col in range(1, TOTAL_COLS+1):
        ws.cell(4, col).border = bdr(bottom=sd("medium", ACCENT))

    # ── 3×4 mini-calendar grid (3 months across, 4 rows)
    # cal starts: col 1, 9, 17 (each 7 wide, gaps at 8, 16)
    CAL_W = 7
    cal_col_starts = [1, 9, 17]
    ROWS_PER_CAL = 10  # month name + day hdr + up to 6 weeks + 1 gap row
    cal_row_starts = [5, 5+ROWS_PER_CAL, 5+2*ROWS_PER_CAL, 5+3*ROWS_PER_CAL]

    for mi, mo in enumerate(range(1, 13)):
        grow = mi // 3
        gcol = mi  % 3
        cr   = cal_row_starts[grow]
        cc   = cal_col_starts[gcol]

        # Month name header
        ws.row_dimensions[cr].height = 15
        c = mc(ws, cr, cc, cr, cc+CAL_W-1)
        c.value = MN[mo]; c.font = tf(9, bold=True, color=TEXT_DARK)
        c.alignment = al("center")
        bg(ws, cr, cc, cr, cc+CAL_W-1, WARM_LIGHT)

        # Day-of-week header (S M T W T F S)
        ws.row_dimensions[cr+1].height = 10
        for d, ltr in enumerate(["S","M","T","W","T","F","S"]):
            cell = ws.cell(cr+1, cc+d)
            cell.value = ltr
            cell.font  = mf(6, bold=True, color=TEXT_MED)
            cell.alignment = al("center")

        # Date numbers (Sunday-first)
        for wi, week in enumerate(calendar.monthcalendar(2028, mo)):
            sun = week[6]; row_days = [sun] + week[:6]
            ws.row_dimensions[cr+2+wi].height = 11
            for d, day_num in enumerate(row_days):
                if day_num:
                    cell = ws.cell(cr+2+wi, cc+d)
                    cell.value = day_num
                    cell.font  = mf(7, color=TEXT_MED)
                    cell.alignment = al("center")

    # ── Divider strip
    bg(ws, 5, 24, 50, 24, WARM_MED)

    # ── Right: Goals + Notes (cols 25-42)
    RS = 25
    RE = TOTAL_COLS

    sections = [
        (5,  "Goals for 2028",          13),
        (14, "Intentions & Hopes",      24),
        (25, "Word(s) for the Year Ahead",    33),
        (34, "Things to Change",  44),
    ]
    for sr, slbl, er in sections:
        ws.row_dimensions[sr].height = 20
        c = mc(ws, sr, RS, sr, RE)
        c.value = slbl; c.font = tf(11, bold=True, color=ACCENT); c.alignment = L
        bg(ws, sr, RS, sr, RE, ACCENT_LIGHT)
        for lr in range(sr+2, min(er, 51)):
            ws.row_dimensions[lr].height = 16
            dot_line(ws, lr, RS, RE)

    # ── Branding at bottom
    ws.row_dimensions[51].height = 6
    ws.row_dimensions[52].height = 18
    ws.row_dimensions[53].height = 8
    _add_branding(ws, 52, 2, TOTAL_COLS-1)

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# SEASONAL PLANNING TAB
# ══════════════════════════════════════════════════════════════════════════════
SEASON_ICONS = {"Spring":"🌸","Summer":"☀️","Fall":"🍂","Winter":"❄️"}

SEASON_TASKS = {
    "cleaning": {
        "all": [
            "Declutter all common areas (living room, kitchen, entryways)",
            "Wipe down baseboards throughout the house",
            "Clean light fixtures and ceiling fans",
            "Wipe tops of door frames and door knobs",
            "Wash windows (inside)",
            "Deep clean refrigerator",
        ],
        "Spring": [
            "Garage cleanout + organization",
            "Kids' spring/summer clothing switchover",
            "Wash and store winter bedding",
            "Deep clean oven and stovetop",
        ],
        "Summer": [
            "Declutter kids' toys and gear",
            "Refresh outdoor entertaining area",
            "Restock sunscreen and outdoor supplies",
        ],
        "Fall": [
            "Kids' fall/winter clothing switchover",
            "Bedroom deep clean + organize closets",
            "Store and cover summer outdoor gear",
        ],
        "Winter": [
            "Deep clean guest rooms before holiday gatherings",
            "Holiday décor: organize setup + storage plan",
            "Year-end filing / paperwork organization",
        ],
    },
    "maintenance": {
        "all": [
            "Check & replace furnace/HVAC air filter",
            "Check & test smoke detectors and CO detectors",
            "Clean dryer vent (lint buildup = fire hazard)",
        ],
        "Spring": [
            "Gutter cleaning (winter debris)",
            "Check outdoor hoses and spigots",
            "Service lawn mower / outdoor equipment",
            "Check deck/porch for winter damage",
            "Outdoor furniture: uncover, wash, set up",
        ],
        "Summer": [
            "Check roof for storm damage",
            "Inspect window screens for tears/holes",
            "Check caulking around windows, tubs, sinks",
        ],
        "Fall": [
            "Schedule furnace/HVAC service before winter",
            "Check weatherstripping on doors and windows",
            "Reverse ceiling fans (clockwise for winter heat)",
            "Stock up: ice melt, shovels ready",
            "Flush water heater (remove sediment)",
            "Check pipe insulation in unheated spaces",
        ],
        "Winter": [
            "Monitor gutters for ice dams",
            "Check pipes in unheated areas (prevent freezing)",
            "Check flashlights / power outage emergency kit",
            "Schedule furnace tune-up if not done in fall",
            "Keep walkways salted and clear",
        ],
    },
}

def make_season_tab(wb, season, year):
    ws = wb.create_sheet(f"{season} {year}")
    TOTAL_COLS = 42
    landscape_all(ws, scale=80, rows=54, cols=TOTAL_COLS)

    for c in range(1, TOTAL_COLS+1):
        ws.column_dimensions[get_column_letter(c)].width = 5.5
    for r in range(1, 55): ws.row_dimensions[r].height = 14

    bg(ws, 1, 1, 54, TOTAL_COLS, CREAM)

    # ── Header (rows 1-4)
    bg(ws, 1, 1, 4, TOTAL_COLS, WARM_LIGHT)
    ws.row_dimensions[1].height = 6
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[3].height = 16
    ws.row_dimensions[4].height = 8

    c = mc(ws, 2, 1, 2, TOTAL_COLS)
    c.value = f"{season} {year}"; c.font = tf(26, bold=True); c.alignment = al("center")

    c = mc(ws, 3, 1, 3, TOTAL_COLS)
    c.value = f"A well-laid plan makes every season of life a little more beautiful."
    c.font = mf(9, italic=True, color=TEXT_MED); c.alignment = al("center")

    for col in range(1, TOTAL_COLS+1):
        ws.cell(4, col).border = bdr(bottom=sd("medium", ACCENT))

    # ── Two-column layout: rows 5-32
    LEFT_END = 20
    DIV_COL  = 21
    RIGHT_START = 22

    # Left header — bg covers through divider col to close the gap
    ws.row_dimensions[5].height = 18
    bg(ws, 5, 1, 5, TOTAL_COLS, WARM_LIGHT)   # full-width base first
    c = mc(ws, 5, 1, 5, LEFT_END)
    c.value = "Cleaning & Declutter"
    c.font = tf(11, bold=True, color=TEXT_MED); c.alignment = al("center")

    # Right header
    c = mc(ws, 5, RIGHT_START, 5, TOTAL_COLS)
    c.value = "Home Maintenance"
    c.font = tf(11, bold=True, color=TEXT_MED); c.alignment = al("center")

    # Thin divider line (no filled column)
    for dr in range(5, 33):
        ws.cell(dr, DIV_COL).border = bdr(right=sd("thin", WARM_MED))

    # Left: cleaning tasks
    clean_tasks = SEASON_TASKS["cleaning"]["all"] + SEASON_TASKS["cleaning"].get(season, [])
    # add 2 blank rows
    clean_tasks += ["", ""]
    lr = 6
    for task in clean_tasks[:14]:
        ws.row_dimensions[lr].height = 16
        ws.cell(lr, 1).value = "○"
        ws.cell(lr, 1).font  = mf(8, color=ACCENT)
        ws.cell(lr, 1).alignment = al("center")
        c = mc(ws, lr, 2, lr, LEFT_END)
        c.value = task
        c.font = mf(8, color=TEXT_DARK if task else TEXT_LIGHT)
        c.alignment = al("left", "center")
        if not task:
            dot_line(ws, lr, 2, LEFT_END)
        lr += 1

    # Right: maintenance tasks
    maint_tasks = SEASON_TASKS["maintenance"]["all"] + SEASON_TASKS["maintenance"].get(season, [])
    maint_tasks += ["", "", ""]
    mr = 6
    for task in maint_tasks[:14]:
        ws.row_dimensions[mr].height = 16
        ws.cell(mr, RIGHT_START).value = "○"
        ws.cell(mr, RIGHT_START).font  = mf(8, color=TEXT_MED)
        ws.cell(mr, RIGHT_START).alignment = al("center")
        c = mc(ws, mr, RIGHT_START+1, mr, TOTAL_COLS)
        c.value = task
        c.font = mf(8, color=TEXT_DARK if task else TEXT_LIGHT)
        c.alignment = al("left", "center")
        if not task:
            dot_line(ws, mr, RIGHT_START+1, TOTAL_COLS)
        mr += 1

    # ── Three-column bottom band: rows 33-53
    THIRD = TOTAL_COLS // 3   # 14
    SEC_TITLES = ["Scripture to Lean Into This Season", "New Recipes to Try", "Notes"]
    for i, title in enumerate(SEC_TITLES):
        sc = 1 + i * THIRD
        ec = sc + THIRD - 1 if i < 2 else TOTAL_COLS
        ws.row_dimensions[33].height = 18
        c = mc(ws, 33, sc, 33, ec)
        c.value = title; c.font = mf(8, bold=True, color=ACCENT if i == 0 else TEXT_MED)
        c.alignment = al("center")
        bg(ws, 33, sc, 33, ec, WARM_LIGHT)
        for wr in range(34, 54):
            ws.row_dimensions[wr].height = 14
            dot_line(ws, wr, sc, ec)

    # ── Branding
    ws.row_dimensions[54].height = 16
    _add_branding(ws, 54, 2, TOTAL_COLS-1)

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
    make_goals_page(wb, "Family")
    make_goals_page(wb, "Personal")
    make_goals_page(wb, "Professional")
    make_vision_board(wb)
    make_bucket_list(wb)

    SEASON_MONTHS = {
        (2026, 9):  ("Fall",   2026),
        (2026, 12): ("Winter", 2026),
        (2027, 3):  ("Spring", 2027),
        (2027, 6):  ("Summer", 2027),
        (2027, 9):  ("Fall",   2027),
        (2027, 12): ("Winter", 2027),
    }

    print("Monthly + weekly layouts...")
    for (year, month) in MONTHS_PLANNER:
        print(f"  {MN[month]} {year}")
        if (year, month) in SEASON_MONTHS:
            s, sy = SEASON_MONTHS[(year, month)]
            make_season_tab(wb, s, sy)
        make_monthly_layout(wb, year, month)
        make_monthly_habit_tracker(wb, year, month)
        for week in weeks_of_month(year, month):
            make_weekly_layout(wb, week, MN[month], year)

    print("Next year page...")
    make_next_year_page(wb)

    out = "/home/user/momlifeplanner/MomLifePlanner.xlsx"
    wb.save(out)
    print(f"\nSaved → {out}  ({len(wb.sheetnames)} sheets)")

if __name__ == "__main__":
    build_planner()
