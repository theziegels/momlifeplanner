"""
Ultimate Mom Life Planner Generator
Aug 2026 – Dec 2027 (15 months + July 2026 / Jan 2028 in year views)
Output: MomLifePlanner.xlsx
"""

import calendar
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_TEXT

# ── Brand Colors ─────────────────────────────────────────────────────────────
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
SPINE        = "D4C9BE"

# ── Fonts ─────────────────────────────────────────────────────────────────────
def tf(size=11, bold=False, color=TEXT_DARK, italic=False, name="Cormorant Garamond"):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)

def mf(size=9, bold=False, color=TEXT_DARK, italic=False):
    return Font(name="Montserrat", size=size, bold=bold, color=color, italic=italic)

# ── Fills ─────────────────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

# ── Borders ───────────────────────────────────────────────────────────────────
def side(style="thin", color=WARM_DARK):
    return Side(style=style, color=color)

def border(left=None, right=None, top=None, bottom=None):
    return Border(left=left or Side(style=None),
                  right=right or Side(style=None),
                  top=top or Side(style=None),
                  bottom=bottom or Side(style=None))

thin_bottom  = border(bottom=side())
thin_all     = border(left=side(), right=side(), top=side(), bottom=side())
accent_bottom = border(bottom=side("medium", ACCENT))
accent_all    = border(left=side("medium", ACCENT), right=side("medium", ACCENT),
                       top=side("medium", ACCENT), bottom=side("medium", ACCENT))

# ── Alignment helpers ─────────────────────────────────────────────────────────
def align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

CENTER = align("center")
LEFT   = align("left")
RIGHT  = align("right")

# ── Utility: set column widths ────────────────────────────────────────────────
def set_col_widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

def set_row_heights(ws, heights: dict):
    for row, h in heights.items():
        ws.row_dimensions[row].height = h

def mc(ws, r1, c1, r2, c2):
    """Merge cells by row/col indices."""
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    return ws.cell(r1, c1)

def write(ws, row, col, value, font=None, fill_=None, align_=None, border_=None):
    cell = ws.cell(row=row, column=col, value=value)
    if font:   cell.font = font
    if fill_:  cell.fill = fill_
    if align_: cell.alignment = align_
    if border_: cell.border = border_
    return cell

def bg_range(ws, r1, c1, r2, c2, hex_color):
    f = fill(hex_color)
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            cell.fill = f

def border_range(ws, r1, c1, r2, c2, bdr):
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            cell.border = bdr

def lined_rows(ws, r_start, r_end, c_start, c_end, color=WARM_MED):
    """Add a thin bottom border (ruled line) to each row in a range."""
    b = border(bottom=side("hair", color))
    for r in range(r_start, r_end + 1):
        for c in range(c_start, c_end + 1):
            ws.cell(r, c).border = b

# ── Month list ────────────────────────────────────────────────────────────────
MONTHS_PLANNER = []  # Aug 2026 – Dec 2027
for y, m in [(2026, mo) for mo in range(8, 13)] + [(2027, mo) for mo in range(1, 13)]:
    MONTHS_PLANNER.append((y, m))

MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]
MONTH_ABBR  = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
DAY_ABBR    = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
DAY_FULL    = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

def weeks_of_month(year, month):
    """Return list of (week_start_date, [date|None ×7]) for calendar weeks touching month."""
    cal = calendar.monthcalendar(year, month)
    weeks = []
    first_day = date(year, month, 1)
    # find Monday of first week
    start = first_day - timedelta(days=first_day.weekday())
    last_day = date(year, month, calendar.monthrange(year, month)[1])
    d = start
    while d <= last_day:
        week = [d + timedelta(days=i) for i in range(7)]
        weeks.append(week)
        d += timedelta(weeks=1)
    return weeks

# ═══════════════════════════════════════════════════════════════════════════════
# SHEET BUILDERS
# ═══════════════════════════════════════════════════════════════════════════════

def make_cover(wb):
    ws = wb.create_sheet("Cover")
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9  # A4

    # column widths – 10 cols
    for c in range(1, 11):
        ws.column_dimensions[get_column_letter(c)].width = 8.5
    for r in range(1, 60):
        ws.row_dimensions[r].height = 15

    # Full background
    bg_range(ws, 1, 1, 60, 10, CREAM)

    # Top accent bar
    bg_range(ws, 1, 1, 3, 10, ACCENT_LIGHT)
    mc(ws, 1, 1, 3, 10)

    # Decorative side strip
    bg_range(ws, 1, 1, 60, 1, WARM_MED)

    # ── Main title block
    bg_range(ws, 8, 2, 14, 10, CREAM)

    c = mc(ws, 10, 2, 10, 10)
    c.value = "THE ULTIMATE"
    c.font = tf(13, bold=False, color=TEXT_MED, italic=True)
    c.alignment = CENTER

    c = mc(ws, 12, 2, 12, 10)
    c.value = "Mom Life"
    c.font = tf(38, bold=True, color=TEXT_DARK)
    c.alignment = CENTER

    c = mc(ws, 13, 2, 13, 10)
    c.value = "P  L  A  N  N  E  R"
    c.font = mf(11, bold=False, color=ACCENT)
    c.alignment = CENTER

    # Thin accent line
    for col in range(3, 10):
        ws.cell(15, col).border = border(bottom=side("medium", ACCENT))

    # Date range
    c = mc(ws, 17, 2, 17, 10)
    c.value = "August 2026 – December 2027"
    c.font = mf(10, color=TEXT_MED, italic=True)
    c.alignment = CENTER

    # Decorative boxes
    bg_range(ws, 22, 3, 24, 4, WARM_LIGHT)
    bg_range(ws, 22, 6, 24, 7, WARM_LIGHT)
    bg_range(ws, 22, 9, 24, 9, WARM_LIGHT)
    for box in [(22,3,24,4),(22,6,24,7),(22,9,24,9)]:
        border_range(ws, *box, border(left=side("thin", WARM_DARK), right=side("thin", WARM_DARK),
                                      top=side("thin", WARM_DARK), bottom=side("thin", WARM_DARK)))

    # Quote
    c = mc(ws, 35, 2, 36, 10)
    c.value = '"You are doing better than you think."'
    c.font = tf(12, italic=True, color=TEXT_MED)
    c.alignment = CENTER

    # Bottom tagline
    c = mc(ws, 50, 2, 50, 10)
    c.value = "plan · dream · thrive"
    c.font = mf(9, color=TEXT_LIGHT, italic=True)
    c.alignment = CENTER

    # Bottom accent bar
    bg_range(ws, 58, 1, 60, 10, WARM_LIGHT)

    return ws


def make_how_to_use(wb):
    ws = wb.create_sheet("How To Use")
    ws.sheet_view.showGridLines = False
    bg_range(ws, 1, 1, 80, 12, CREAM)

    for c in range(1, 13):
        ws.column_dimensions[get_column_letter(c)].width = 8
    for r in range(1, 80):
        ws.row_dimensions[r].height = 14

    bg_range(ws, 1, 1, 5, 12, WARM_LIGHT)
    c = mc(ws, 3, 1, 3, 12)
    c.value = "How to Use This Planner"
    c.font = tf(22, bold=True, color=TEXT_DARK)
    c.alignment = CENTER

    sections = [
        (8,  "Year at a Glance",
              "Get a bird's-eye view of 2026 & 2027. Mark holidays, big events, and "
              "anything that anchors your year."),
        (14, "Monthly Layout",
              "Each month opens with a full calendar grid — perfect for mapping "
              "appointments, school events, and family milestones at a glance."),
        (20, "Weekly Layout",
              "Your detailed planning home. Left side holds Mon–Sun with "
              "lined space for each day. Right side holds Priorities, Habit "
              "Trackers, a To-Do list, and a Notes/quote space."),
        (26, "Goals Pages",
              "Separate pages for Family, Personal, and Professional goals. "
              "Break big dreams into quarterly milestones and action steps."),
        (32, "Brain Map & Brainstorm",
              "Use the Brain Map when you need to organize a central idea outward. "
              "Use the Brainstorm page for free-flowing idea generation — "
              "great for projects, meal plans, or any big life season."),
        (38, "Vision Board",
              "Capture images, words, and feelings that represent your ideal life. "
              "Return to it often to stay inspired and aligned."),
    ]

    for row, title, body in sections:
        bg_range(ws, row, 2, row, 2, ACCENT)
        ws.row_dimensions[row].height = 18
        c = mc(ws, row, 3, row, 12)
        c.value = title
        c.font = tf(13, bold=True, color=ACCENT)
        c.alignment = align("left", "center")

        c = mc(ws, row+1, 3, row+3, 12)
        c.value = body
        c.font = mf(9, color=TEXT_MED)
        c.alignment = align("left", "top")

    c = mc(ws, 55, 2, 56, 11)
    c.value = "This planner is yours — write in it, doodle in it, make it beautiful and real."
    c.font = tf(11, italic=True, color=TEXT_MED)
    c.alignment = CENTER

    return ws


def make_important_contacts(wb):
    ws = wb.create_sheet("Important Contacts")
    ws.sheet_view.showGridLines = False
    bg_range(ws, 1, 1, 80, 14, CREAM)

    for c in range(1, 15):
        ws.column_dimensions[get_column_letter(c)].width = 7.5
    for r in range(1, 80):
        ws.row_dimensions[r].height = 14

    bg_range(ws, 1, 1, 5, 14, WARM_LIGHT)
    c = mc(ws, 3, 1, 3, 14)
    c.value = "Important Contacts"
    c.font = tf(22, bold=True, color=TEXT_DARK)
    c.alignment = CENTER

    categories = [
        "Emergency", "Pediatrician", "Dentist", "School / Teacher",
        "Babysitter / Childcare", "Husband / Partner", "Mom / In-Laws",
        "Plumber / Electrician", "Pharmacy", "Other"
    ]

    col_labels = ["Name", "Phone", "Email", "Notes"]
    col_positions = [2, 5, 8, 11]
    col_widths = [3, 3, 3, 3]

    # Header row
    r = 8
    ws.row_dimensions[r].height = 16
    for i, lbl in enumerate(col_labels):
        c = mc(ws, r, col_positions[i], r, col_positions[i]+2)
        c.value = lbl
        c.font = mf(8, bold=True, color=ACCENT)
        c.alignment = align("left", "center")
        bg_range(ws, r, col_positions[i], r, col_positions[i]+2, WARM_LIGHT)

    r = 9
    for cat in categories:
        ws.row_dimensions[r].height = 16
        # Category label
        c = mc(ws, r, 1, r, 1)
        c.value = cat
        c.font = mf(8, bold=True, color=TEXT_MED)
        c.alignment = align("left", "center")

        for col_p in col_positions:
            cell = mc(ws, r, col_p, r, col_p+2)
            cell.border = border(bottom=side("thin", WARM_DARK))

        r += 2

    return ws


def make_year_view(wb, year, show_before=None, show_after=None, months_highlight=None):
    """Year at a glance. show_before/after = (year,month) for faded months."""
    name = f"Year {year}"
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    bg_range(ws, 1, 1, 85, 30, CREAM)

    for c in range(1, 31):
        ws.column_dimensions[get_column_letter(c)].width = 3.8
    for r in range(1, 85):
        ws.row_dimensions[r].height = 13

    # Title
    bg_range(ws, 1, 1, 4, 30, WARM_LIGHT)
    c = mc(ws, 2, 1, 3, 30)
    c.value = str(year)
    c.font = tf(30, bold=True, color=TEXT_DARK)
    c.alignment = CENTER

    # Layout: 4 rows × 3 cols = 12 months + optional before/after
    all_months = []
    if show_before:
        all_months.append(show_before)
    for m in range(1, 13):
        all_months.append((year, m))
    if show_after:
        all_months.append(show_after)

    # grid: show 14 months in 2 rows × 7 cols if we have before+after, else 12 in 3×4
    # Use 4 rows × 3 cols for main months, extras faded on sides
    # Simple approach: 2 rows of 7
    grid_months = all_months
    n = len(grid_months)
    cols_per_row = 5 if n > 12 else 4
    if n <= 7:
        cols_per_row = 7

    # grid positions (col_start per calendar)
    CAL_COLS = 7  # 7 day columns
    SPACING   = 1  # gap col

    # Draw months in grid: 3 wide, multiple rows
    per_row = 4
    start_row = 7
    for idx, (yr, mo) in enumerate(all_months):
        grid_row = idx // per_row
        grid_col = idx % per_row

        r_off = start_row + grid_row * 18
        c_off = 1 + grid_col * (CAL_COLS + 2)

        is_highlight = months_highlight and (yr, mo) in months_highlight
        is_faded     = not is_highlight and months_highlight is not None

        # Month name header
        c = mc(ws, r_off, c_off, r_off, c_off + CAL_COLS - 1)
        c.value = f"{MONTH_NAMES[mo]} {yr}" if yr != year else MONTH_NAMES[mo]
        c.font = tf(10, bold=True,
                    color=ACCENT if is_highlight else (TEXT_LIGHT if is_faded else TEXT_MED))
        c.alignment = CENTER
        if is_highlight:
            bg_range(ws, r_off, c_off, r_off, c_off + CAL_COLS - 1, ACCENT_LIGHT)

        # Day headers S M T W T F S (Sun first)
        day_letters = ["S","M","T","W","T","F","S"]
        for d, ltr in enumerate(day_letters):
            cell = ws.cell(r_off + 1, c_off + d, ltr)
            cell.font = mf(7, bold=True,
                           color=ACCENT if d in (0,6) else TEXT_MED)
            cell.alignment = CENTER

        # Calendar grid
        cal = calendar.monthcalendar(yr, mo)
        # monthcalendar returns Mon-first; we want Sun-first
        # Re-arrange: sun = last col → first
        for week_idx, week in enumerate(cal):
            sun = week[6]
            row_days = [sun] + week[:6]
            for d_idx, day_num in enumerate(row_days):
                cell = ws.cell(r_off + 2 + week_idx, c_off + d_idx)
                if day_num != 0:
                    cell.value = day_num
                cell.font = mf(7, color=TEXT_LIGHT if is_faded else
                               (ACCENT if d_idx in (0,6) else TEXT_DARK))
                cell.alignment = CENTER

    return ws


def make_notes_page(wb, title, subtitle=""):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    bg_range(ws, 1, 1, 70, 14, CREAM)

    for c in range(1, 15):
        ws.column_dimensions[get_column_letter(c)].width = 7.5
    for r in range(1, 70):
        ws.row_dimensions[r].height = 14

    bg_range(ws, 1, 1, 5, 14, WARM_LIGHT)
    cell = mc(ws, 3, 1, 3, 14)
    cell.value = title
    cell.font = tf(22, bold=True, color=TEXT_DARK)
    cell.alignment = CENTER

    if subtitle:
        c = mc(ws, 4, 1, 4, 14)
        c.value = subtitle
        c.font = mf(9, italic=True, color=TEXT_MED)
        c.alignment = CENTER

    for r in range(7, 68):
        for col in range(2, 14):
            ws.cell(r, col).border = border(bottom=side("hair", WARM_MED))
        ws.row_dimensions[r].height = 18

    return ws


def make_goals_page(wb, goal_type):
    ws = wb.create_sheet(f"{goal_type} Goals")
    ws.sheet_view.showGridLines = False
    bg_range(ws, 1, 1, 80, 14, CREAM)

    for c in range(1, 15):
        ws.column_dimensions[get_column_letter(c)].width = 7.5
    for r in range(1, 80):
        ws.row_dimensions[r].height = 14

    bg_range(ws, 1, 1, 5, 14, WARM_LIGHT)
    cell = mc(ws, 3, 1, 3, 14)
    cell.value = f"{goal_type} Goals"
    cell.font = tf(22, bold=True, color=TEXT_DARK)
    cell.alignment = CENTER

    # Sections: This Year's Big Goals / Quarterly Breakdown / Action Steps
    sections = [
        (8,  "My Big Goals This Year",   14, 4),
        (24, "Quarter by Quarter",        30, 4),
        (46, "Action Steps & Milestones", 60, 4),
    ]

    for s_row, s_title, e_row, indent in sections:
        c = mc(ws, s_row, 2, s_row, 13)
        c.value = s_title
        c.font = tf(13, bold=True, color=ACCENT)
        c.alignment = align("left", "center")
        bg_range(ws, s_row, 2, s_row, 13, ACCENT_LIGHT)

        if "Quarter" in s_title:
            q_labels = ["Q1  Jan–Mar", "Q2  Apr–Jun", "Q3  Jul–Sep", "Q4  Oct–Dec"]
            q_starts = [2, 5, 8, 11]
            for q_i, (q_lbl, q_col) in enumerate(zip(q_labels, q_starts)):
                c = mc(ws, s_row+2, q_col, s_row+2, q_col+2)
                c.value = q_lbl
                c.font = mf(8, bold=True, color=TEXT_MED)
                c.alignment = align("left")
                for lr in range(s_row+3, s_row+12):
                    for lc in range(q_col, q_col+3):
                        ws.cell(lr, lc).border = border(bottom=side("hair", WARM_MED))
                    ws.row_dimensions[lr].height = 16
        else:
            for lr in range(s_row+2, e_row):
                for lc in range(2, 14):
                    if "Big" in s_title:
                        # Numbered goals
                        if lc == 2:
                            num = lr - (s_row+1)
                            if 1 <= num <= 10:
                                ws.cell(lr, lc).value = f"{num}."
                                ws.cell(lr, lc).font = mf(9, color=ACCENT)
                    ws.cell(lr, lc).border = border(bottom=side("hair", WARM_MED))
                ws.row_dimensions[lr].height = 18

    return ws


def make_brain_map(wb):
    ws = wb.create_sheet("Brain Map")
    ws.sheet_view.showGridLines = False
    bg_range(ws, 1, 1, 70, 22, CREAM)

    for c in range(1, 23):
        ws.column_dimensions[get_column_letter(c)].width = 5
    for r in range(1, 70):
        ws.row_dimensions[r].height = 14

    bg_range(ws, 1, 1, 4, 22, WARM_LIGHT)
    cell = mc(ws, 2, 1, 3, 22)
    cell.value = "Brain Map"
    cell.font = tf(22, bold=True, color=TEXT_DARK)
    cell.alignment = CENTER

    c = mc(ws, 4, 1, 4, 22)
    c.value = "Place your central idea in the circle, then branch outward."
    c.font = mf(8, italic=True, color=TEXT_MED)
    c.alignment = CENTER

    # Central circle area (simulated with merged cells + border + fill)
    center_r, center_c = 35, 10
    # "Circle" = oval box 6 rows × 4 cols in center
    bg_range(ws, center_r-2, center_c-1, center_r+2, center_c+2, ACCENT_LIGHT)
    cell = mc(ws, center_r-2, center_c-1, center_r+2, center_c+2)
    cell.value = "Central\nIdea"
    cell.font = tf(12, bold=True, color=ACCENT)
    cell.alignment = CENTER
    border_range(ws, center_r-2, center_c-1, center_r+2, center_c+2,
                 border(left=side("medium", ACCENT), right=side("medium", ACCENT),
                        top=side("medium", ACCENT), bottom=side("medium", ACCENT)))

    # Branch lines + sub-topic boxes in 8 directions
    # Each branch: label box at end of "line" (represented by filled cells with bottom border)
    branches = [
        # (direction_label, row, col_start, col_end, row2, col2)  "line" then "box"
        ("Top-Left",     10, 5,  8,  8, 5,  7),
        ("Top",          10, 11, 11, 8, 10, 12),
        ("Top-Right",    10, 16, 14, 8, 14, 16),
        ("Left",         35, 3,  6,  35, 2, 5),
        ("Right",        35, 15, 18, 35, 16, 19),
        ("Bottom-Left",  55, 5,  8,  57, 5, 7),
        ("Bottom",       55, 11, 11, 57, 10, 12),
        ("Bottom-Right", 55, 15, 14, 57, 14, 16),
    ]

    for i, (lbl, lr, lc, lc2, br, bc, bc2) in enumerate(branches):
        # "line" row of cells with bottom border
        for col in range(lc, lc2+1):
            ws.cell(lr, col).border = border(bottom=side("thin", WARM_DARK))

        # Sub-topic box
        bg_range(ws, br, bc, br+2, bc2, WARM_LIGHT)
        c = mc(ws, br, bc, br+2, bc2)
        c.value = ""
        border_range(ws, br, bc, br+2, bc2,
                     border(left=side("thin", WARM_DARK), right=side("thin", WARM_DARK),
                            top=side("thin", WARM_DARK), bottom=side("thin", WARM_DARK)))

        # Sub-branch lines from each box (3 lines)
        for offset in [-1, 0, 1]:
            sub_r = br + 1 + offset
            for sub_c in range(max(1, bc-2), max(1, bc)):
                if sub_r >= 1 and sub_c >= 1:
                    ws.cell(sub_r, sub_c).border = border(bottom=side("hair", WARM_MED))

    return ws


def make_brainstorm(wb):
    ws = wb.create_sheet("Brainstorm")
    ws.sheet_view.showGridLines = False
    bg_range(ws, 1, 1, 75, 22, CREAM)

    for c in range(1, 23):
        ws.column_dimensions[get_column_letter(c)].width = 5
    for r in range(1, 75):
        ws.row_dimensions[r].height = 14

    bg_range(ws, 1, 1, 4, 22, WARM_LIGHT)
    cell = mc(ws, 2, 1, 3, 22)
    cell.value = "Brainstorm"
    cell.font = tf(22, bold=True, color=TEXT_DARK)
    cell.alignment = CENTER

    c = mc(ws, 4, 1, 4, 22)
    c.value = "Start with your main idea. Let every branch spark the next."
    c.font = mf(8, italic=True, color=TEXT_MED)
    c.alignment = CENTER

    # Central horizontal spine (fishbone / tree style)
    spine_row = 38
    for col in range(2, 22):
        ws.cell(spine_row, col).border = border(bottom=side("medium", TEXT_MED))

    # Main topic box on left end of spine
    bg_range(ws, spine_row-2, 1, spine_row+2, 3, CREAM)
    cell = mc(ws, spine_row-1, 1, spine_row+1, 3)
    cell.value = "Main\nIdea"
    cell.font = tf(11, bold=True, color=ACCENT)
    cell.alignment = CENTER
    bg_range(ws, spine_row-1, 1, spine_row+1, 3, ACCENT_LIGHT)
    border_range(ws, spine_row-1, 1, spine_row+1, 3,
                 border(left=side("medium", ACCENT), right=side("medium", ACCENT),
                        top=side("medium", ACCENT), bottom=side("medium", ACCENT)))

    # Branch positions along spine (col numbers)
    branch_cols = [5, 8, 11, 14, 17, 20]
    for i, bc in enumerate(branch_cols):
        alt = i % 2  # alternate up/down

        for arm in range(1, 5):  # 4 sub-branches each direction
            # Upper branches
            r_up = spine_row - arm * 3
            if r_up >= 6:
                c = mc(ws, r_up, bc, r_up, bc+1)
                c.border = border(bottom=side("thin", WARM_DARK))
                # Leaf line
                for leaf in range(1, 3):
                    ws.cell(r_up, bc - leaf).border = border(bottom=side("hair", WARM_MED))

            # Lower branches
            r_dn = spine_row + arm * 3
            if r_dn <= 72:
                c = mc(ws, r_dn, bc, r_dn, bc+1)
                c.border = border(bottom=side("thin", WARM_DARK))
                for leaf in range(1, 3):
                    ws.cell(r_dn, bc - leaf).border = border(bottom=side("hair", WARM_MED))

    return ws


def make_vision_board(wb):
    ws = wb.create_sheet("Vision Board")
    ws.sheet_view.showGridLines = False
    bg_range(ws, 1, 1, 75, 18, CREAM)

    for c in range(1, 19):
        ws.column_dimensions[get_column_letter(c)].width = 6.5
    for r in range(1, 75):
        ws.row_dimensions[r].height = 14

    bg_range(ws, 1, 1, 5, 18, WARM_LIGHT)
    cell = mc(ws, 2, 1, 3, 18)
    cell.value = "Vision Board"
    cell.font = tf(26, bold=True, color=TEXT_DARK)
    cell.alignment = CENTER

    c = mc(ws, 4, 1, 4, 18)
    c.value = "Paste images · write words · capture feelings · dream boldly"
    c.font = mf(8, italic=True, color=TEXT_MED)
    c.alignment = CENTER

    # Word/phrase prompts
    words = ["FAITH", "FAMILY", "HEALTH", "JOY", "GROWTH", "PEACE",
             "ADVENTURE", "HOME", "PURPOSE", "LOVE", "ABUNDANCE", "FREEDOM"]
    for i, word in enumerate(words):
        r_ = 7 + (i // 4) * 3
        c_ = 2 + (i % 4) * 4
        cell = mc(ws, r_, c_, r_+1, c_+2)
        cell.value = word
        cell.font = mf(9, bold=True, color=TEXT_LIGHT)
        cell.alignment = CENTER
        bg_range(ws, r_, c_, r_+1, c_+2, WARM_LIGHT)
        border_range(ws, r_, c_, r_+1, c_+2,
                     border(left=side("thin", WARM_DARK), right=side("thin", WARM_DARK),
                            top=side("thin", WARM_DARK), bottom=side("thin", WARM_DARK)))

    # Large image box sections
    boxes = [
        (20, 2, 42, 9,   "paste an image or write your dream here"),
        (20, 10, 42, 17, "paste an image or write your dream here"),
        (44, 2, 62, 9,   "this year I will..."),
        (44, 10, 62, 17, "I am grateful for..."),
        (64, 2, 72, 17, "my word(s) for this season:"),
    ]
    for r1, c1, r2, c2, hint in boxes:
        bg_range(ws, r1, c1, r2, c2, WARM_LIGHT)
        border_range(ws, r1, c1, r2, c2,
                     border(left=side("thin", WARM_MED), right=side("thin", WARM_MED),
                            top=side("thin", WARM_MED), bottom=side("thin", WARM_MED)))
        cell = mc(ws, r1, c1, r2, c2)
        cell.value = hint
        cell.font = mf(8, italic=True, color=TEXT_LIGHT)
        cell.alignment = CENTER

    return ws


def make_bucket_list(wb):
    ws = wb.create_sheet("Bucket List")
    ws.sheet_view.showGridLines = False
    bg_range(ws, 1, 1, 80, 14, CREAM)

    for c in range(1, 15):
        ws.column_dimensions[get_column_letter(c)].width = 7.5
    for r in range(1, 80):
        ws.row_dimensions[r].height = 14

    bg_range(ws, 1, 1, 5, 14, WARM_LIGHT)
    cell = mc(ws, 3, 1, 3, 14)
    cell.value = "Bucket List"
    cell.font = tf(22, bold=True, color=TEXT_DARK)
    cell.alignment = CENTER

    c = mc(ws, 4, 1, 4, 14)
    c.value = "Places to go · things to try · memories to make"
    c.font = mf(8, italic=True, color=TEXT_MED)
    c.alignment = CENTER

    cats = [("Travel & Adventures", 8), ("Family Experiences", 28),
            ("Personal Growth", 48), ("Just for Fun", 62)]
    for cat_name, start_r in cats:
        c = mc(ws, start_r, 2, start_r, 13)
        c.value = cat_name
        c.font = tf(12, bold=True, color=ACCENT)
        c.alignment = align("left", "center")
        bg_range(ws, start_r, 2, start_r, 13, ACCENT_LIGHT)

        for lr in range(start_r+2, start_r+16):
            ws.cell(lr, 2).value = "○"
            ws.cell(lr, 2).font = mf(9, color=ACCENT)
            ws.cell(lr, 2).alignment = CENTER
            for lc in range(3, 14):
                ws.cell(lr, lc).border = border(bottom=side("hair", WARM_MED))
            ws.row_dimensions[lr].height = 18

    return ws


def make_self_care_tracker(wb):
    ws = wb.create_sheet("Self-Care Tracker")
    ws.sheet_view.showGridLines = False
    bg_range(ws, 1, 1, 80, 20, CREAM)

    for c in range(1, 21):
        ws.column_dimensions[get_column_letter(c)].width = 5.5
    for r in range(1, 80):
        ws.row_dimensions[r].height = 14

    bg_range(ws, 1, 1, 5, 20, WARM_LIGHT)
    cell = mc(ws, 3, 1, 3, 20)
    cell.value = "Self-Care Tracker"
    cell.font = tf(22, bold=True, color=TEXT_DARK)
    cell.alignment = CENTER

    c = mc(ws, 4, 1, 4, 20)
    c.value = "You can't pour from an empty cup. Track the habits that fill yours."
    c.font = mf(8, italic=True, color=TEXT_MED)
    c.alignment = CENTER

    habits = [
        "8 hrs sleep", "Drank water", "Moved my body", "Time outside",
        "Read / learned", "Journaled", "Connected with a friend",
        "Ate nourishing food", "Screen-free time", "Quiet time / prayer",
        "Bath / skincare", "Something just for me",
    ]

    # Day number headers 1–31
    start_r = 9
    ws.cell(start_r, 2).value = "Habit"
    ws.cell(start_r, 2).font = mf(8, bold=True, color=TEXT_MED)

    for d in range(1, 32):
        c = ws.cell(start_r, d+2)
        c.value = d
        c.font = mf(7, bold=True, color=TEXT_MED)
        c.alignment = CENTER

    for i, habit in enumerate(habits):
        r = start_r + 2 + i * 2
        ws.row_dimensions[r].height = 18
        c = mc(ws, r, 2, r, 2)
        c.value = habit
        c.font = mf(8, color=TEXT_DARK)
        bg_range(ws, r, 2, r, 2, WARM_LIGHT)

        for d in range(1, 32):
            cell = ws.cell(r, d+2)
            cell.value = "○"
            cell.font = mf(8, color=ACCENT_LIGHT[:-2]+"AA" if True else WARM_DARK)
            cell.font = Font(name="Montserrat", size=8, color=WARM_DARK)
            cell.alignment = CENTER
            cell.border = border(left=side("hair", WARM_MED), bottom=side("hair", WARM_MED))

    return ws


def make_monthly_layout(wb, year, month):
    """Full calendar grid for a month."""
    name = f"{MONTH_ABBR[month]} {year}"
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    bg_range(ws, 1, 1, 55, 20, CREAM)

    # Column setup: 7 day columns + notes column
    day_col_w = 11.5
    for c in range(1, 9):
        ws.column_dimensions[get_column_letter(c)].width = day_col_w
    ws.column_dimensions["I"].width = 1  # spacer
    for c in range(10, 21):
        ws.column_dimensions[get_column_letter(c)].width = 8

    for r in range(1, 55):
        ws.row_dimensions[r].height = 13

    # Header
    bg_range(ws, 1, 1, 5, 20, WARM_LIGHT)
    c = mc(ws, 2, 1, 2, 8)
    c.value = MONTH_NAMES[month]
    c.font = tf(28, bold=True, color=TEXT_DARK, italic=True)
    c.alignment = align("left", "center")

    c = mc(ws, 3, 1, 3, 8)
    c.value = str(year)
    c.font = mf(11, color=TEXT_MED)
    c.alignment = align("left", "center")

    # Accent underline
    for col in range(1, 9):
        ws.cell(5, col).border = border(bottom=side("medium", ACCENT))

    # Day headers (Sun–Sat)
    day_headers = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    for d, dh in enumerate(day_headers):
        cell = ws.cell(7, d+1)
        cell.value = dh
        cell.font = mf(8, bold=True, color=TEXT_MED)
        cell.alignment = CENTER
        bg_range(ws, 7, d+1, 7, d+1, WARM_LIGHT)

    # Calendar weeks
    cal = calendar.monthcalendar(year, month)
    row = 9
    for week in cal:
        ws.row_dimensions[row].height = 14
        # Re-order: Sun first (monthcalendar gives Mon-first)
        sun = week[6]
        week_days = [sun] + week[:6]
        for d_idx, day_num in enumerate(week_days):
            col = d_idx + 1
            if day_num != 0:
                # Day number
                cell = ws.cell(row, col)
                cell.value = day_num
                cell.font = mf(9, bold=True,
                               color=ACCENT if d_idx == 0 else TEXT_DARK)
                cell.alignment = align("left", "top")
                cell.border = border(top=side("thin", WARM_MED),
                                     left=side("hair", WARM_MED))

                # Writing space rows
                for lr in range(row+1, row+5):
                    ws.cell(lr, col).border = border(bottom=side("hair", WARM_MED),
                                                     left=side("hair", WARM_MED))
                    ws.row_dimensions[lr].height = 14
            else:
                for lr in range(row, row+5):
                    ws.cell(lr, col).fill = fill(WARM_LIGHT)
        row += 5

    # Notes sidebar
    c = mc(ws, 7, 10, 7, 20)
    c.value = "Monthly Notes & Reminders"
    c.font = tf(11, bold=True, color=ACCENT)
    c.alignment = align("left", "center")
    bg_range(ws, 7, 10, 7, 20, ACCENT_LIGHT)

    for lr in range(9, 52):
        for lc in range(10, 21):
            ws.cell(lr, lc).border = border(bottom=side("hair", WARM_MED))
        ws.row_dimensions[lr].height = 16

    # Mini sections in notes area
    sections_r = [9, 22, 35, 44]
    section_labels = ["Appointments", "Birthdays & Occasions",
                      "Meal Ideas", "Don't Forget"]
    for sr, sl in zip(sections_r, section_labels):
        c = mc(ws, sr, 10, sr, 20)
        c.value = sl
        c.font = mf(8, bold=True, color=TEXT_MED)
        c.alignment = align("left", "center")
        bg_range(ws, sr, 10, sr, 20, WARM_LIGHT)

    return ws


def make_weekly_layout(wb, week_dates, month_name, year):
    """
    Single 8.5×11 LANDSCAPE page.
    Left ~55% : month header + 7 day blocks in 2-col grid (Mon/Tue, Wed/Thu, Fri/Sat, Sun full)
    Right ~45%: Priorities box | Habits tracker | To-Do | Notes
    Matches the reference photo layout.
    """
    start_d = week_dates[0]
    end_d   = week_dates[6]
    sheet_name = f"Wk {start_d.strftime('%b %-d')}"

    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    # ── Print setup: landscape letter ─────────────────────────────────────────
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 1   # Letter
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_view.zoomScale   = 75

    # ── Column layout (landscape 11" wide) ────────────────────────────────────
    # Left section (days): cols 1–14
    #   Day-left cols  1–6  (Mon/Wed/Fri/Sun-left)
    #   Gap col        7    (narrow divider between day columns)
    #   Day-right cols 8–13 (Tue/Thu/Sat/Sun-right)
    #   Gap col        14
    # Right section (sidebar): cols 15–26
    #   Priorities     15–19
    #   Habits         20–26

    # Column widths tuned to land in 11" landscape print area (~100 Excel units ≈ 11")
    left_day_w  = 9.5   # cols 1-6 and 8-13
    gap_w       = 1.2
    pri_w       = 9.5   # priorities cols 15-19
    hab_w       = 4.8   # habits cols 20-26 (7 cols)

    for c in list(range(1, 7)) + list(range(8, 14)):
        ws.column_dimensions[get_column_letter(c)].width = left_day_w
    ws.column_dimensions["G"].width = gap_w   # mid-day gap
    ws.column_dimensions["N"].width = gap_w   # left/right section gap
    for c in range(15, 20):
        ws.column_dimensions[get_column_letter(c)].width = pri_w
    for c in range(20, 27):
        ws.column_dimensions[get_column_letter(c)].width = hab_w

    # Row heights: aim for ~55 rows to fill 8.5" landscape height
    for r in range(1, 58):
        ws.row_dimensions[r].height = 13.2

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 10  # accent underline row

    # ── Backgrounds ────────────────────────────────────────────────────────────
    bg_range(ws, 1, 1,  57, 13, CREAM)
    bg_range(ws, 1, 14, 57, 14, WARM_MED)   # spine/divider strip
    bg_range(ws, 1, 15, 57, 26, CREAM)

    # ── LEFT: Header ──────────────────────────────────────────────────────────
    bg_range(ws, 1, 1, 2, 13, WARM_LIGHT)

    c = mc(ws, 1, 1, 2, 8)
    c.value = f"{month_name.lower()}  {year}"
    c.font = tf(16, bold=True, italic=True, color=TEXT_DARK)
    c.alignment = align("left", "center")

    c = mc(ws, 1, 9, 2, 13)
    c.value = f"{start_d.strftime('%-m/%-d')} – {end_d.strftime('%-m/%-d')}"
    c.font = mf(8, italic=True, color=TEXT_MED)
    c.alignment = align("right", "center")

    # Accent underline
    for col in range(1, 14):
        ws.cell(3, col).border = border(bottom=side("medium", ACCENT))

    # ── LEFT: Day blocks ──────────────────────────────────────────────────────
    # Grid: Mon(col 1-6) | Tue(col 8-13)  row 4-16
    #       Wed(col 1-6) | Thu(col 8-13)  row 17-29
    #       Fri(col 1-6) | Sat(col 8-13)  row 30-42
    #       Sun(col 1-13)                 row 43-55

    LINES = 9   # writing lines per day
    BLOCK = LINES + 3  # total rows per day block (header + gap + lines)

    def draw_day(day_idx, r_start, c_start, c_end):
        d = week_dates[day_idx]
        is_weekend = d.weekday() >= 5

        ws.row_dimensions[r_start].height = 15

        # Day number
        num = ws.cell(r_start, c_start)
        num.value = d.day
        num.font = mf(9, bold=True, color=ACCENT if is_weekend else TEXT_DARK)
        num.alignment = align("left", "center")

        # Day name
        nm = mc(ws, r_start, c_start+1, r_start, c_end)
        nm.value = DAY_FULL[day_idx].upper()
        nm.font = mf(7, color=TEXT_MED)
        nm.alignment = align("left", "center")

        # Underline header
        for col in range(c_start, c_end+1):
            ws.cell(r_start, col).border = border(bottom=side("thin", WARM_MED))

        # Ruled writing lines
        for lr in range(r_start+1, r_start+1+LINES):
            ws.row_dimensions[lr].height = 13.5
            for col in range(c_start, c_end+1):
                ws.cell(lr, col).border = border(bottom=side("hair", WARM_MED))

        # Gap row
        ws.row_dimensions[r_start+LINES+1].height = 5

    pairs = [(0,1), (2,3), (4,5)]  # (Mon,Tue), (Wed,Thu), (Fri,Sat)
    row_starts = [4, 4+BLOCK, 4+BLOCK*2]

    for (li, ri), rs in zip(pairs, row_starts):
        draw_day(li, rs, 1, 6)
        draw_day(ri, rs, 8, 13)

    # Sunday full width
    sun_r = 4 + BLOCK * 3
    draw_day(6, sun_r, 1, 13)

    # ── RIGHT: Header ─────────────────────────────────────────────────────────
    bg_range(ws, 1, 15, 2, 26, WARM_LIGHT)

    # ── RIGHT: PRIORITIES (rows 1-18, cols 15-19) ────────────────────────────
    c = mc(ws, 1, 15, 1, 19)
    c.value = "PRIORITIES"
    c.font = mf(7, bold=True, color=ACCENT)
    c.alignment = CENTER
    bg_range(ws, 1, 15, 2, 19, ACCENT_LIGHT)

    # Priorities box with ruled lines
    border_range(ws, 3, 15, 18, 19,
                 border(left=side("thin", WARM_DARK), right=side("thin", WARM_DARK),
                        top=side("thin", WARM_DARK), bottom=side("thin", WARM_DARK)))
    for lr in range(4, 19):
        for lc in range(15, 20):
            ws.cell(lr, lc).border = border(bottom=side("hair", WARM_MED))

    # ── RIGHT: HABITS (rows 1-18, cols 20-26) ────────────────────────────────
    c = mc(ws, 1, 20, 1, 26)
    c.value = "HABITS"
    c.font = mf(7, bold=True, color=TEXT_MED)
    c.alignment = CENTER
    bg_range(ws, 1, 20, 2, 26, WARM_LIGHT)

    # Day letter headers  M T W T F S S
    day_letters = ["M","T","W","T","F","S","S"]
    ws.row_dimensions[3].height = 11
    for di, dl in enumerate(day_letters):
        cell = ws.cell(3, 20+di)
        cell.value = dl
        cell.font = mf(6, bold=True, color=TEXT_MED)
        cell.alignment = CENTER

    # 6 habit rows (each 2.5 rows tall)
    for h in range(6):
        hr = 4 + h * 2
        ws.row_dimensions[hr].height = 15

        # Habit name cell (spans cols 20 only, rest are checkboxes)
        name_c = ws.cell(hr, 20)
        name_c.border = border(bottom=side("thin", WARM_DARK))
        bg_range(ws, hr, 20, hr, 20, ACCENT_LIGHT)

        for di in range(7):
            box = ws.cell(hr, 20+di)
            if di > 0:
                box.value = "○"
                box.font = Font(name="Montserrat", size=7, color=WARM_DARK)
                box.alignment = CENTER
            box.border = border(bottom=side("thin", WARM_MED),
                                left=side("hair", WARM_MED))

        ws.row_dimensions[hr+1].height = 4  # gap

    # ── RIGHT: TO-DO (rows 19-36) ─────────────────────────────────────────────
    todo_r = 19
    bg_range(ws, todo_r, 15, todo_r+1, 26, WARM_LIGHT)

    c = mc(ws, todo_r, 15, todo_r, 19)
    c.value = "TO-DO"
    c.font = mf(7, bold=True, color=TEXT_MED)
    c.alignment = align("left", "center")

    c = mc(ws, todo_r, 20, todo_r, 26)
    c.value = "TO-DO  (cont'd)"
    c.font = mf(7, bold=True, color=TEXT_MED)
    c.alignment = align("left", "center")

    for ti in range(9):
        tr = todo_r + 2 + ti
        ws.row_dimensions[tr].height = 14

        # Left to-do column
        ws.cell(tr, 15).value = "○"
        ws.cell(tr, 15).font = Font(name="Montserrat", size=8, color=ACCENT)
        ws.cell(tr, 15).alignment = CENTER
        for lc in range(16, 20):
            ws.cell(tr, lc).border = border(bottom=side("hair", WARM_MED))

        # Right to-do column
        ws.cell(tr, 20).value = "○"
        ws.cell(tr, 20).font = Font(name="Montserrat", size=8, color=ACCENT)
        ws.cell(tr, 20).alignment = CENTER
        for lc in range(21, 27):
            ws.cell(tr, lc).border = border(bottom=side("hair", WARM_MED))

    # ── RIGHT: NOTES (rows 37-57) ─────────────────────────────────────────────
    notes_r = 37
    bg_range(ws, notes_r, 15, notes_r+1, 26, ACCENT_LIGHT)

    c = mc(ws, notes_r, 15, notes_r, 26)
    c.value = "NOTES"
    c.font = mf(7, bold=True, color=ACCENT)
    c.alignment = align("left", "center")

    for nr in range(notes_r+2, 56):
        ws.row_dimensions[nr].height = 14
        for nc in range(15, 27):
            ws.cell(nr, nc).border = border(bottom=side("hair", WARM_MED))

    return ws


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN BUILD
# ═══════════════════════════════════════════════════════════════════════════════

def build_planner():
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    print("Building cover & intro pages...")
    make_cover(wb)
    make_how_to_use(wb)
    make_important_contacts(wb)

    print("Building year views...")
    # Year 2026: show July as 'before', highlight Aug–Dec
    highlight_2026 = [(2026, m) for m in range(8, 13)]
    make_year_view(wb, 2026,
                   show_before=(2026, 7),
                   months_highlight=highlight_2026)

    # Year 2027: all months highlighted, show Jan 2028 after
    highlight_2027 = [(2027, m) for m in range(1, 13)]
    make_year_view(wb, 2027,
                   show_after=(2028, 1),
                   months_highlight=highlight_2027)

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
        print(f"  {MONTH_NAMES[month]} {year}")
        make_monthly_layout(wb, year, month)

        weeks = weeks_of_month(year, month)
        for week in weeks:
            make_weekly_layout(wb, week, MONTH_NAMES[month], year)

    out = "/home/user/momlifeplanner/MomLifePlanner.xlsx"
    wb.save(out)
    print(f"\nSaved → {out}")
    return out

if __name__ == "__main__":
    build_planner()
