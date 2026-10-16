"""
analysis.py
E-Commerce Product Analytics & Conversion Optimization
Reads output/events.csv, runs SQL analysis via SQLite,
and writes a fully-formatted Excel dashboard.

Dependencies: openpyxl  (pip install openpyxl)
"""

import csv
import math
import sqlite3
import os
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side,
    GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.worksheet.filters import AutoFilter

# ─────────────────────────── Colour Palette ──────────────────────────────────
C_DARK_BG   = "1A1F2E"   # deep navy background
C_HEADER_BG = "2C3E6B"   # header fill
C_ACCENT1   = "4472C4"   # blue accent
C_ACCENT2   = "70AD47"   # green
C_ACCENT3   = "FF0000"   # red (drop-off)
C_ACCENT4   = "FFC000"   # amber (warning)
C_WHITE     = "FFFFFF"
C_LIGHT_GRAY= "F2F2F2"
C_MID_GRAY  = "D9D9D9"
C_DARK_TEXT = "1F2937"
C_SUBHEAD   = "44546A"

# Border helpers
THIN_BORDER = Border(
    left=Side(style="thin",   color="B0B0B0"),
    right=Side(style="thin",  color="B0B0B0"),
    top=Side(style="thin",    color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
MED_BORDER = Border(
    left=Side(style="medium",   color=C_ACCENT1),
    right=Side(style="medium",  color=C_ACCENT1),
    top=Side(style="medium",    color=C_ACCENT1),
    bottom=Side(style="medium", color=C_ACCENT1),
)

def hdr_fill(hex_color): return PatternFill("solid", fgColor=hex_color)
def hdr_font(bold=True, color=C_WHITE, size=11):
    return Font(name="Calibri", bold=bold, color=color, size=size)
def body_font(bold=False, color=C_DARK_TEXT, size=10):
    return Font(name="Calibri", bold=bold, color=color, size=size)
def center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def left():   return Alignment(horizontal="left",   vertical="center", wrap_text=True)

def style_header_row(ws, row, cols, bg=C_HEADER_BG, fg=C_WHITE, size=11):
    for c in range(1, cols+1):
        cell = ws.cell(row=row, column=c)
        cell.fill   = hdr_fill(bg)
        cell.font   = hdr_font(color=fg, size=size)
        cell.alignment = center()
        cell.border = THIN_BORDER

def style_data_row(ws, row, cols, alt=False):
    bg = C_LIGHT_GRAY if alt else C_WHITE
    for c in range(1, cols+1):
        cell = ws.cell(row=row, column=c)
        cell.fill      = hdr_fill(bg)
        cell.font      = body_font()
        cell.alignment = center()
        cell.border    = THIN_BORDER

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def freeze(ws, cell="A2"):
    ws.freeze_panes = cell

# ─────────────────────────── Load Data ───────────────────────────────────────
def load_events(path="output/events.csv"):
    events = []
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            row["price"]   = float(row["price"])
            row["revenue"] = float(row["revenue"])
            events.append(row)
    return events

def build_db(events):
    conn = sqlite3.connect(":memory:")
    conn.row_factory = sqlite3.Row
    conn.execute("""
        CREATE TABLE events (
            event_id TEXT, user_id TEXT, session_id TEXT, event_type TEXT,
            product_id TEXT, category TEXT, brand TEXT, price REAL,
            price_band TEXT, user_type TEXT, timestamp TEXT, revenue REAL
        )""")
    conn.executemany(
        "INSERT INTO events VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
        [(e["event_id"],e["user_id"],e["session_id"],e["event_type"],
          e["product_id"],e["category"],e["brand"],e["price"],
          e["price_band"],e["user_type"],e["timestamp"],e["revenue"])
         for e in events])
    conn.commit()
    return conn

def q(conn, sql):
    return [dict(r) for r in conn.execute(sql).fetchall()]

# ─────────────────────────── Sheet 1: Overview KPIs ─────────────────────────
def sheet_overview(wb, conn):
    ws = wb.create_sheet("Overview KPIs")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C_ACCENT1

    # ── Title ──
    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value     = "E-COMMERCE PRODUCT ANALYTICS  |  Conversion Optimization Dashboard"
    t.fill      = hdr_fill(C_DARK_BG)
    t.font      = Font(name="Calibri", bold=True, color=C_WHITE, size=14)
    t.alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:H2")
    ws["A2"].value     = "Reporting Period: July 18 – October 15, 2024   |   Data Refreshed: Auto-generated"
    ws["A2"].fill      = hdr_fill(C_SUBHEAD)
    ws["A2"].font      = Font(name="Calibri", italic=True, color=C_WHITE, size=10)
    ws["A2"].alignment = center()

    # ── KPI Data ──
    row = q(conn, """
        SELECT
          SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END) views,
          SUM(CASE WHEN event_type='add_to_cart'  THEN 1 ELSE 0 END) carts,
          SUM(CASE WHEN event_type='purchase'     THEN 1 ELSE 0 END) purchases,
          COUNT(DISTINCT user_id)    unique_users,
          COUNT(DISTINCT session_id) sessions,
          ROUND(100.0*SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) v2c,
          ROUND(100.0*SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END),0),2) c2p,
          ROUND(100.0*SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) cvr,
          ROUND(SUM(revenue),2) revenue,
          ROUND(SUM(revenue)/NULLIF(SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END),0),2) aov
        FROM events""")[0]

    kpis = [
        ("Total Sessions",        f"{row['sessions']:,}",    C_ACCENT1),
        ("Unique Users",          f"{row['unique_users']:,}", "3A86FF"),
        ("Product Views",         f"{row['views']:,}",       "8338EC"),
        ("Add-to-Cart Events",    f"{row['carts']:,}",       C_ACCENT4),
        ("Purchases",             f"{row['purchases']:,}",   C_ACCENT2),
        ("View→Cart Rate",        f"{row['v2c']}%",          C_ACCENT4),
        ("Cart→Purchase Rate",    f"{row['c2p']}%",          C_ACCENT2),
        ("Overall CVR",           f"{row['cvr']}%",          "FF006E"),
        ("Total Revenue",         f"${row['revenue']:,.2f}", C_ACCENT2),
        ("Avg Order Value",       f"${row['aov']:,.2f}",     "06D6A0"),
    ]

    # Two rows of KPI cards (5 per row)
    for idx, (label, value, color) in enumerate(kpis):
        col   = (idx % 5) * 2 + 1   # columns 1,3,5,7,9
        r_top = 4 if idx < 5 else 8

        ws.merge_cells(start_row=r_top,   start_column=col, end_row=r_top,   end_column=col+1)
        ws.merge_cells(start_row=r_top+1, start_column=col, end_row=r_top+1, end_column=col+1)
        ws.merge_cells(start_row=r_top+2, start_column=col, end_row=r_top+2, end_column=col+1)

        lbl_cell = ws.cell(row=r_top,   column=col, value=label)
        val_cell = ws.cell(row=r_top+1, column=col, value=value)
        spc_cell = ws.cell(row=r_top+2, column=col, value="")

        lbl_cell.fill      = hdr_fill(color)
        lbl_cell.font      = Font(name="Calibri", bold=True, color=C_WHITE, size=10)
        lbl_cell.alignment = center()
        val_cell.fill      = hdr_fill(C_WHITE)
        val_cell.font      = Font(name="Calibri", bold=True, color=color, size=18)
        val_cell.alignment = center()
        spc_cell.fill      = hdr_fill(C_MID_GRAY)

        for r in [r_top, r_top+1, r_top+2]:
            for c in range(col, col+2):
                ws.cell(row=r, column=c).border = THIN_BORDER
        ws.row_dimensions[r_top].height   = 22
        ws.row_dimensions[r_top+1].height = 38
        ws.row_dimensions[r_top+2].height = 6

    # ── Funnel Summary Table ──
    ws.merge_cells("A11:J11")
    ws["A11"].value     = "CONVERSION FUNNEL SUMMARY"
    ws["A11"].fill      = hdr_fill(C_HEADER_BG)
    ws["A11"].font      = hdr_font(size=12)
    ws["A11"].alignment = center()
    ws.row_dimensions[11].height = 24

    headers = ["Funnel Stage","Events","% of Views","Drop-off (n)","Drop-off %","Status"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=12, column=c).value = h
    style_header_row(ws, 12, len(headers), bg="44546A")

    v, ca, p = row["views"], row["carts"], row["purchases"]
    funnel_data = [
        ("Product View",  v,  100.0,          0,                0.0,         "🟢 Baseline"),
        ("Add to Cart",   ca, round(ca/v*100,2), v-ca,          round((v-ca)/v*100,2), "🟡 Monitor"),
        ("Purchase",      p,  round(p/v*100,2),  ca-p,          round((ca-p)/ca*100,2),"🔴 Optimize"),
    ]
    for i, (stage, events_n, pct, drop_n, drop_pct, status) in enumerate(funnel_data):
        r = 13 + i
        vals = [stage, f"{events_n:,}", f"{pct}%", f"{drop_n:,}", f"{drop_pct}%", status]
        for c, v2 in enumerate(vals, 1):
            ws.cell(row=r, column=c).value = v2
        style_data_row(ws, r, len(headers), alt=(i%2==0))

    set_col_widths(ws, [24, 16, 14, 14, 14, 16, 14, 14, 14, 14])
    freeze(ws, "A3")

# ─────────────────────────── Sheet 2: Funnel by Category ────────────────────
def sheet_funnel(wb, conn):
    ws = wb.create_sheet("Funnel by Category")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "FF6B6B"

    ws.merge_cells("A1:I1")
    ws["A1"].value     = "CONVERSION FUNNEL ANALYSIS BY CATEGORY"
    ws["A1"].fill      = hdr_fill(C_DARK_BG)
    ws["A1"].font      = hdr_font(size=13)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 28

    headers = ["Category","Views","Add-to-Cart","Purchases",
               "View→Cart %","Cart→Purchase %","Overall CVR %","Revenue ($)","Avg Order Value ($)"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=2, column=c).value = h
    style_header_row(ws, 2, len(headers))

    rows = q(conn, """
        SELECT category,
          SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END) views,
          SUM(CASE WHEN event_type='add_to_cart'  THEN 1 ELSE 0 END) carts,
          SUM(CASE WHEN event_type='purchase'     THEN 1 ELSE 0 END) purchases,
          ROUND(100.0*SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) v2c,
          ROUND(100.0*SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END),0),2) c2p,
          ROUND(100.0*SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) cvr,
          ROUND(SUM(revenue),2) revenue,
          ROUND(SUM(revenue)/NULLIF(SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END),0),2) aov
        FROM events GROUP BY category ORDER BY cvr DESC""")

    for i, r in enumerate(rows):
        row_n = 3 + i
        vals  = [r["category"], r["views"], r["carts"], r["purchases"],
                 r["v2c"], r["c2p"], r["cvr"], r["revenue"], r["aov"]]
        for c, v in enumerate(vals, 1):
            ws.cell(row=row_n, column=c).value = v
        style_data_row(ws, row_n, len(headers), alt=(i%2==0))

    # Conditional formatting on CVR column (col 7)
    last = 2 + len(rows)
    ws.conditional_formatting.add(
        f"G3:G{last}",
        ColorScaleRule(start_type="min", start_color="FF6B6B",
                       mid_type="percentile", mid_value=50, mid_color="FFFF00",
                       end_type="max", end_color="70AD47"))

    # ── Bar Chart ──
    chart = BarChart()
    chart.type        = "col"
    chart.title       = "Overall Conversion Rate by Category (%)"
    chart.y_axis.title= "CVR (%)"
    chart.x_axis.title= "Category"
    chart.style       = 10
    chart.grouping    = "clustered"
    chart.width, chart.height = 18, 10

    cats = Reference(ws, min_col=1, min_row=3, max_row=2+len(rows))
    data = Reference(ws, min_col=7, min_row=2, max_row=2+len(rows))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "A" + str(last + 3))

    set_col_widths(ws, [16, 12, 14, 12, 14, 18, 14, 14, 18])
    freeze(ws, "A3")
    ws.auto_filter.ref = f"A2:I{last}"

# ─────────────────────────── Sheet 3: Segmentation ───────────────────────────
def sheet_segmentation(wb, conn):
    ws = wb.create_sheet("Segmentation")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "70AD47"

    row_offset = 1

    # ── Helper ──
    def write_block(title, sql_rows, cols, bg=C_HEADER_BG):
        nonlocal row_offset
        ws.merge_cells(start_row=row_offset, start_column=1,
                       end_row=row_offset, end_column=len(cols))
        t = ws.cell(row=row_offset, column=1, value=title)
        t.fill      = hdr_fill(bg)
        t.font      = hdr_font(size=12)
        t.alignment = center()
        ws.row_dimensions[row_offset].height = 24
        row_offset += 1

        for c, h in enumerate(cols, 1):
            ws.cell(row=row_offset, column=c).value = h
        style_header_row(ws, row_offset, len(cols), bg="44546A")
        row_offset += 1

        for i, r in enumerate(sql_rows):
            for c, k in enumerate(r.keys(), 1):
                ws.cell(row=row_offset, column=c).value = r[k]
            style_data_row(ws, row_offset, len(cols), alt=(i%2==0))
            row_offset += 1
        row_offset += 2   # spacing

    # By Brand
    brand_rows = q(conn, """
        SELECT brand, category,
          SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END) views,
          SUM(CASE WHEN event_type='add_to_cart'  THEN 1 ELSE 0 END) carts,
          SUM(CASE WHEN event_type='purchase'     THEN 1 ELSE 0 END) purchases,
          ROUND(100.0*SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) v2c,
          ROUND(100.0*SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END),0),2) c2p,
          ROUND(100.0*SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) cvr,
          ROUND(SUM(revenue),2) revenue
        FROM events GROUP BY brand, category ORDER BY cvr DESC LIMIT 20""")
    write_block("BY BRAND — Top 20 (Ranked by CVR)", brand_rows,
                ["Brand","Category","Views","Cart","Purchases","View→Cart %","Cart→Pur %","Overall CVR %","Revenue ($)"])

    # By Price Band
    price_rows = q(conn, """
        SELECT price_band,
          SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END) views,
          SUM(CASE WHEN event_type='add_to_cart'  THEN 1 ELSE 0 END) carts,
          SUM(CASE WHEN event_type='purchase'     THEN 1 ELSE 0 END) purchases,
          ROUND(100.0*SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) v2c,
          ROUND(100.0*SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END),0),2) c2p,
          ROUND(100.0*SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) cvr,
          ROUND(SUM(revenue),2) revenue
        FROM events GROUP BY price_band
        ORDER BY CASE price_band WHEN '<$25' THEN 1 WHEN '$25-$75' THEN 2
                                  WHEN '$75-$200' THEN 3 ELSE 4 END""")
    write_block("BY PRICE BAND", price_rows,
                ["Price Band","Views","Add-to-Cart","Purchases","View→Cart %","Cart→Pur %","Overall CVR %","Revenue ($)"])

    # By User Type
    user_rows = q(conn, """
        SELECT user_type,
          COUNT(DISTINCT user_id) unique_users,
          SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END) views,
          SUM(CASE WHEN event_type='add_to_cart'  THEN 1 ELSE 0 END) carts,
          SUM(CASE WHEN event_type='purchase'     THEN 1 ELSE 0 END) purchases,
          ROUND(100.0*SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) v2c,
          ROUND(100.0*SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END),0),2) c2p,
          ROUND(100.0*SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) cvr,
          ROUND(SUM(revenue),2) revenue
        FROM events GROUP BY user_type ORDER BY cvr DESC""")
    write_block("BY USER TYPE — New vs. Returning", user_rows,
                ["User Type","Unique Users","Views","Cart","Purchases","View→Cart %","Cart→Pur %","Overall CVR %","Revenue ($)"])

    set_col_widths(ws, [16,14,12,12,12,14,14,14,14])

# ─────────────────────────── Sheet 4: Drop-off Analysis ─────────────────────
def sheet_dropoff(wb, conn):
    ws = wb.create_sheet("Drop-off Analysis")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "FF0000"

    ws.merge_cells("A1:H1")
    ws["A1"].value     = "DROP-OFF ANALYSIS — Identifying Conversion Leaks"
    ws["A1"].fill      = hdr_fill("8B0000")
    ws["A1"].font      = hdr_font(size=13)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 28

    # ── View → Cart Drop-off ──
    ws.merge_cells("A2:H2")
    ws["A2"].value     = "VIEW → CART DROP-OFF (Ranked Worst First)"
    ws["A2"].fill      = hdr_fill(C_HEADER_BG)
    ws["A2"].font      = hdr_font(size=11)
    ws["A2"].alignment = center()

    h1 = ["Category","Views","Add-to-Cart","Lost Users","Drop-off %","Impact","Priority","Recommendation"]
    for c, h in enumerate(h1, 1):
        ws.cell(row=3, column=c).value = h
    style_header_row(ws, 3, len(h1), bg="8B0000", fg=C_WHITE)

    rows_v2c = q(conn, """
        SELECT category,
          SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END) views,
          SUM(CASE WHEN event_type='add_to_cart'  THEN 1 ELSE 0 END) carts,
          SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END)
           -SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END) lost,
          ROUND(100.0*(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END)
                      -SUM(CASE WHEN event_type='add_to_cart'  THEN 1 ELSE 0 END))
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) dropoff_pct
        FROM events GROUP BY category ORDER BY dropoff_pct DESC""")

    impact_map  = {0: "🔴 Critical", 1: "🟠 High", 2: "🟡 Medium", 3: "🟢 Low", 4: "🟢 Low"}
    priority_map = {0: "P1", 1: "P1", 2: "P2", 3: "P3", 4: "P3"}
    reco_map = {
        "Electronics": "Improve product imagery, add comparison tools, display trust badges",
        "Apparel":     "Add size guides, more angles, customer photo reviews",
        "Home":        "Add room-visualizer, bundle deals, detailed specs",
        "Beauty":      "Add shade finder, ingredient highlights, skin-type filters",
        "Sports":      "Add fit guide, activity filter, performance specs",
    }

    for i, r in enumerate(rows_v2c):
        row_n = 4 + i
        impact   = impact_map.get(i, "🟢 Low")
        priority = priority_map.get(i, "P3")
        reco     = reco_map.get(r["category"], "Improve product detail page")
        vals = [r["category"], r["views"], r["carts"], r["lost"],
                f"{r['dropoff_pct']}%", impact, priority, reco]
        for c, v in enumerate(vals, 1):
            ws.cell(row=row_n, column=c).value = v
        style_data_row(ws, row_n, len(h1), alt=(i%2==0))
        # Red tint for top drop-offs
        if i < 2:
            for c in range(1, len(h1)+1):
                ws.cell(row=row_n, column=c).fill = hdr_fill("FFE0E0")

    row_off = 4 + len(rows_v2c) + 2

    # ── Cart → Purchase Drop-off ──
    ws.merge_cells(f"A{row_off}:H{row_off}")
    ws.cell(row=row_off, column=1).value     = "CART → PURCHASE DROP-OFF (Ranked Worst First)"
    ws.cell(row=row_off, column=1).fill      = hdr_fill(C_HEADER_BG)
    ws.cell(row=row_off, column=1).font      = hdr_font(size=11)
    ws.cell(row=row_off, column=1).alignment = center()

    h2 = ["Category","Add-to-Cart","Purchases","Abandoned","Abandon %","Impact","Priority","Recommendation"]
    for c, h in enumerate(h2, 1):
        ws.cell(row=row_off+1, column=c).value = h
    style_header_row(ws, row_off+1, len(h2), bg="8B0000")

    rows_c2p = q(conn, """
        SELECT category,
          SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END) carts,
          SUM(CASE WHEN event_type='purchase'    THEN 1 ELSE 0 END) purchases,
          SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END)
           -SUM(CASE WHEN event_type='purchase'  THEN 1 ELSE 0 END) abandoned,
          ROUND(100.0*(SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END)
                      -SUM(CASE WHEN event_type='purchase'    THEN 1 ELSE 0 END))
                /NULLIF(SUM(CASE WHEN event_type='add_to_cart' THEN 1 ELSE 0 END),0),2) abandon_pct
        FROM events GROUP BY category ORDER BY abandon_pct DESC""")

    cart_reco = {
        "Electronics": "Add EMI/financing options, price-match guarantee, reduce checkout steps",
        "Apparel":     "Free shipping threshold, easy returns messaging, size exchange assurance",
        "Home":        "Highlight delivery date, assembly service offer, payment plans",
        "Beauty":      "Add loyalty points reminder, free samples at checkout, subscribe & save",
        "Sports":      "Flash sale timer, bundle discount, guest checkout option",
    }

    for i, r in enumerate(rows_c2p):
        row_n  = row_off + 2 + i
        impact   = impact_map.get(i, "🟢 Low")
        priority = priority_map.get(i, "P3")
        reco     = cart_reco.get(r["category"], "Streamline checkout")
        vals = [r["category"], r["carts"], r["purchases"], r["abandoned"],
                f"{r['abandon_pct']}%", impact, priority, reco]
        for c, v in enumerate(vals, 1):
            ws.cell(row=row_n, column=c).value = v
        style_data_row(ws, row_n, len(h2), alt=(i%2==0))
        if i < 2:
            for c in range(1, len(h2)+1):
                ws.cell(row=row_n, column=c).fill = hdr_fill("FFE0E0")

    set_col_widths(ws, [16, 12, 14, 12, 12, 14, 10, 52])
    freeze(ws, "A2")

# ─────────────────────────── Sheet 5: A/B Testing Framework ─────────────────
def sample_size(baseline_cvr, mde, alpha=0.05, power=0.80):
    """Two-tailed z-test sample size per variant."""
    z_alpha = 1.96   # alpha=0.05
    z_beta  = 0.842  # power=0.80
    p1      = baseline_cvr / 100
    p2      = p1 * (1 + mde / 100)
    p_avg   = (p1 + p2) / 2
    n = ((z_alpha * math.sqrt(2 * p_avg * (1 - p_avg))
          + z_beta * math.sqrt(p1*(1-p1) + p2*(1-p2))) ** 2) / ((p2 - p1) ** 2)
    return int(math.ceil(n))

def sheet_ab(wb, conn):
    ws = wb.create_sheet("AB Test Framework")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "FFC000"

    ws.merge_cells("A1:J1")
    ws["A1"].value     = "A/B TESTING FRAMEWORK — Hypotheses, Sample Sizes & Methodology"
    ws["A1"].fill      = hdr_fill(C_DARK_BG)
    ws["A1"].font      = hdr_font(size=13)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 28

    # ── Methodology Box ──
    ws.merge_cells("A2:J5")
    ws["A2"].value = (
        "METHODOLOGY: All tests use two-tailed z-tests (α=0.05, power=80%). "
        "Minimum detectable effect (MDE) = minimum business-meaningful lift. "
        "Tests run until significance OR 4 weeks (whichever first). "
        "No peeking — evaluate only at pre-defined checkpoints (Day 7, 14, 21, 28)."
    )
    ws["A2"].fill      = hdr_fill("FFF9C4")
    ws["A2"].font      = Font(name="Calibri", color="5D4037", size=10, italic=True)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 60

    # ── Sample Size Calculator Table ──
    ws.merge_cells("A6:J6")
    ws["A6"].value     = "SAMPLE SIZE CALCULATOR (per variant, 80% power, α=0.05)"
    ws["A6"].fill      = hdr_fill(C_HEADER_BG)
    ws["A6"].font      = hdr_font(size=11)
    ws["A6"].alignment = center()

    calc_hdrs = ["Baseline CVR (%)","MDE (%)","Required n / Variant","Daily Traffic","Est. Duration (days)","Confidence","Statistical Power","Notes"]
    for c, h in enumerate(calc_hdrs, 1):
        ws.cell(row=7, column=c).value = h
    style_header_row(ws, 7, len(calc_hdrs))

    calc_scenarios = [
        (5.0,  10, 3000, "Checkout flow optimisation"),
        (5.0,  15, 3000, "Low-lift urgency test"),
        (8.0,  10, 3000, "Cart recovery email"),
        (8.0,  20, 3000, "Bold MDE — quick win"),
        (12.0, 10, 3000, "High-base category test"),
        (15.0,  5, 3000, "Mature category, small gain"),
        (3.0,  20, 3000, "Low-base, high-gain scenario"),
        (10.0, 10, 3000, "Standard product page test"),
    ]

    for i, (base, mde, daily, note) in enumerate(calc_scenarios):
        n       = sample_size(base, mde)
        dur     = math.ceil(n / (daily / 2))  # split traffic 50/50
        row_n   = 8 + i
        vals    = [f"{base}%", f"{mde}%", f"{n:,}", f"{daily:,}",
                   f"{dur} days", "95%", "80%", note]
        for c, v in enumerate(vals, 1):
            ws.cell(row=row_n, column=c).value = v
        style_data_row(ws, row_n, len(calc_hdrs), alt=(i%2==0))

    # ── Pre-built Hypothesis Table ──
    hyp_start = 8 + len(calc_scenarios) + 2
    ws.merge_cells(f"A{hyp_start}:J{hyp_start}")
    ws.cell(row=hyp_start, column=1).value     = "PRE-BUILT TEST HYPOTHESES — Ranked by Estimated Impact"
    ws.cell(row=hyp_start, column=1).fill      = hdr_fill(C_HEADER_BG)
    ws.cell(row=hyp_start, column=1).font      = hdr_font(size=11)
    ws.cell(row=hyp_start, column=1).alignment = center()

    hyp_hdrs = ["#","Hypothesis","Funnel Stage","Category","Metric","Baseline","Target Lift","Est. Revenue Impact","Priority","Status"]
    for c, h in enumerate(hyp_hdrs, 1):
        ws.cell(row=hyp_start+1, column=c).value = h
    style_header_row(ws, hyp_start+1, len(hyp_hdrs))

    hypotheses = [
        (1, "Add EMI / Buy-Now-Pay-Later at checkout reduces cart abandonment for Electronics",
         "Cart→Purchase","Electronics","Cart Abandon Rate","60%","-15%","$18,000+","🔴 P1","Planned"),
        (2, "Show 'Low Stock' badge on high-demand Apparel increases view-to-cart rate",
         "View→Cart","Apparel","Add-to-Cart Rate","32%","+12%","$9,500+","🔴 P1","Planned"),
        (3, "Checkout progress bar reduces purchase drop-off across all categories",
         "Cart→Purchase","All","Checkout CVR","55%","+8%","$12,000+","🔴 P1","Planned"),
        (4, "Product comparison widget for Electronics improves add-to-cart rate",
         "View→Cart","Electronics","Add-to-Cart Rate","28%","+10%","$7,800+","🟠 P2","Backlog"),
        (5, "Free shipping threshold display at cart increases avg order value",
         "Cart→Purchase","All","AOV","$142","+5%","$6,200+","🟠 P2","Backlog"),
        (6, "Beauty shade-finder tool increases view-to-cart for Beauty category",
         "View→Cart","Beauty","Add-to-Cart Rate","35%","+18%","$5,400+","🟠 P2","Backlog"),
        (7, "Guest checkout option reduces cart-to-purchase drop-off for new users",
         "Cart→Purchase","All (New)","CVR New Users","45%","+10%","$4,900+","🟠 P2","Backlog"),
        (8, "Returning user personalised homepage increases session conversion",
         "View→Cart","All (Returning)","Session CVR","22%","+6%","$3,800+","🟡 P3","Backlog"),
        (9, "Sports activity-based filter reduces browse-to-add friction",
         "View→Cart","Sports","Add-to-Cart Rate","30%","+8%","$2,600+","🟡 P3","Backlog"),
        (10,"Home décor room-visualizer increases add-to-cart for $200+ items",
         "View→Cart","Home","High-ticket CVR","18%","+15%","$4,100+","🟡 P3","Backlog"),
    ]

    for i, hyp in enumerate(hypotheses):
        row_n = hyp_start + 2 + i
        for c, v in enumerate(hyp, 1):
            ws.cell(row=row_n, column=c).value = v
        style_data_row(ws, row_n, len(hyp_hdrs), alt=(i%2==0))

    set_col_widths(ws, [4, 52, 16, 16, 20, 12, 12, 18, 10, 10])
    freeze(ws, "A7")

# ─────────────────────────── Sheet 6: Recommendations ───────────────────────
def sheet_recommendations(wb, conn):
    ws = wb.create_sheet("Recommendations")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C_ACCENT2

    ws.merge_cells("A1:G1")
    ws["A1"].value     = "PRIORITIZED RECOMMENDATIONS — Data-Driven Action Plan"
    ws["A1"].fill      = hdr_fill(C_DARK_BG)
    ws["A1"].font      = hdr_font(size=13)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 28

    headers = ["Priority","Category","Funnel Stage","Finding","Recommendation","Expected Impact","Owner / Timeline"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=2, column=c).value = h
    style_header_row(ws, 2, len(headers))

    recs = [
        ("🔴 HIGH","Electronics","Cart → Purchase",
         "Highest cart abandonment rate (~60%) driven by high-ticket prices",
         "Introduce BNPL / EMI options + price-match guarantee at checkout",
         "Est. +15% cart CVR → +$18K/month revenue","Product Team | 4 weeks"),
        ("🔴 HIGH","All Categories","Cart → Purchase",
         "Users drop off at checkout — no progress indicator",
         "Add multi-step checkout progress bar & trust badges (SSL, returns)",
         "Est. +8% checkout CVR across all categories","Engineering | 2 weeks"),
        ("🔴 HIGH","Apparel","View → Cart",
         "High view volume but lowest add-to-cart rate due to sizing uncertainty",
         "Implement AI-powered size recommender + model size callouts",
         "Est. +12% view-to-cart → +$9.5K/month","UX Team | 6 weeks"),
        ("🟠 MED","Electronics","View → Cart",
         "Users browse but don't add — lack of product comparison capability",
         "Build side-by-side product comparison widget on listing pages",
         "Est. +10% add-to-cart rate for Electronics","Engineering | 5 weeks"),
        ("🟠 MED","Beauty","View → Cart",
         "Shade/tone ambiguity causes high drop-off on cosmetics",
         "Embed virtual shade-finder & skin-tone filter on product pages",
         "Est. +18% add-to-cart → +$5.4K/month","Product | 8 weeks"),
        ("🟠 MED","All (New Users)","Cart → Purchase",
         "New users drop off 15% more at checkout vs. returning",
         "Enable frictionless guest checkout + post-purchase account creation",
         "Est. +10% new-user CVR","Engineering | 2 weeks"),
        ("🟠 MED","All","Cart → Purchase",
         "High-value carts abandoned — no re-engagement mechanism",
         "Deploy cart recovery email sequence (1hr, 24hr, 72hr) with incentive",
         "Est. +8% recovery rate, AOV +5%","Marketing | 3 weeks"),
        ("🟡 LOW","Home","View → Cart",
         "Low add-to-cart for high-ticket Home items — buyers need assurance",
         "Add augmented reality room-visualizer for furniture & decor items",
         "Est. +15% high-ticket item CVR","Product | 12 weeks"),
        ("🟡 LOW","Sports","View → Cart",
         "Browse sessions are long but conversion low — filtering friction",
         "Add activity-type and performance-level smart filters",
         "Est. +8% add-to-cart rate for Sports","Engineering | 4 weeks"),
        ("🟡 LOW","All (Returning)","View → Cart",
         "Returning users are not being personalised on homepage",
         "Personalise returning-user homepage with last-viewed + complementary items",
         "Est. +6% session CVR for returning users","Data/ML Team | 10 weeks"),
    ]

    priority_colors = {"🔴 HIGH": "FFE0E0", "🟠 MED": "FFF3CD", "🟡 LOW": "E8F5E9"}

    for i, rec in enumerate(recs):
        row_n = 3 + i
        for c, v in enumerate(rec, 1):
            cell = ws.cell(row=row_n, column=c, value=v)
            cell.font      = body_font(size=10)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border    = THIN_BORDER
            cell.fill      = hdr_fill(priority_colors.get(rec[0], C_WHITE))

    set_col_widths(ws, [12, 14, 16, 38, 44, 32, 24])
    for r in range(3, 3 + len(recs)):
        ws.row_dimensions[r].height = 48
    freeze(ws, "A3")
    ws.auto_filter.ref = f"A2:G{2+len(recs)}"

# ─────────────────────────── Sheet 7: Raw Data ───────────────────────────────
def sheet_raw(wb, conn):
    ws = wb.create_sheet("Raw Data")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "808080"

    headers = ["event_id","user_id","session_id","event_type","product_id",
               "category","brand","price","price_band","user_type","timestamp","revenue"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c).value = h
    style_header_row(ws, 1, len(headers))

    rows = q(conn, "SELECT * FROM events ORDER BY timestamp LIMIT 5000")
    for i, r in enumerate(rows):
        row_n = 2 + i
        for c, k in enumerate(headers, 1):
            ws.cell(row=row_n, column=c).value = r[k]
        if i % 2 == 0:
            for c in range(1, len(headers)+1):
                ws.cell(row=row_n, column=c).fill = hdr_fill(C_LIGHT_GRAY)

    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    set_col_widths(ws, [12, 10, 10, 14, 10, 14, 14, 10, 12, 12, 20, 10])
    freeze(ws, "A2")

# ─────────────────────────── Sheet 8: Daily Trend ────────────────────────────
def sheet_daily(wb, conn):
    ws = wb.create_sheet("Daily Trend")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "3A86FF"

    ws.merge_cells("A1:F1")
    ws["A1"].value     = "DAILY CONVERSION TREND (July 18 – October 15, 2024)"
    ws["A1"].fill      = hdr_fill(C_DARK_BG)
    ws["A1"].font      = hdr_font(size=13)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 28

    headers = ["Date","Views","Add-to-Cart","Purchases","Daily CVR (%)","Daily Revenue ($)"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=2, column=c).value = h
    style_header_row(ws, 2, len(headers))

    rows = q(conn, """
        SELECT DATE(timestamp) date,
          SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END) views,
          SUM(CASE WHEN event_type='add_to_cart'  THEN 1 ELSE 0 END) carts,
          SUM(CASE WHEN event_type='purchase'     THEN 1 ELSE 0 END) purchases,
          ROUND(100.0*SUM(CASE WHEN event_type='purchase' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN event_type='product_view' THEN 1 ELSE 0 END),0),2) cvr,
          ROUND(SUM(revenue),2) revenue
        FROM events GROUP BY DATE(timestamp) ORDER BY date""")

    for i, r in enumerate(rows):
        row_n = 3 + i
        vals  = [r["date"], r["views"], r["carts"], r["purchases"], r["cvr"], r["revenue"]]
        for c, v in enumerate(vals, 1):
            ws.cell(row=row_n, column=c).value = v
        style_data_row(ws, row_n, len(headers), alt=(i%2==0))

    # CVR trend line chart
    last = 2 + len(rows)
    chart = BarChart()
    chart.type        = "line" if False else "bar"   # use bar for clarity
    chart.title       = "Daily Purchases & Revenue Trend"
    chart.y_axis.title= "Purchases"
    chart.x_axis.title= "Date"
    chart.style       = 10
    chart.width, chart.height = 24, 12

    purch_ref = Reference(ws, min_col=4, min_row=2, max_row=last)
    cat_ref   = Reference(ws, min_col=1, min_row=3, max_row=last)
    chart.add_data(purch_ref, titles_from_data=True)
    chart.set_categories(cat_ref)
    ws.add_chart(chart, f"A{last+3}")

    # Conditional formatting on CVR
    ws.conditional_formatting.add(
        f"E3:E{last}",
        ColorScaleRule(start_type="min", start_color="FF6B6B",
                       end_type="max", end_color="70AD47"))

    set_col_widths(ws, [14, 10, 14, 12, 14, 18])
    freeze(ws, "A3")

# ─────────────────────────── Main ────────────────────────────────────────────
def main():
    os.makedirs("output", exist_ok=True)
    print("Loading events.csv ...")
    events = load_events()
    print(f"   {len(events):,} events loaded.")

    print("Building in-memory SQLite database ...")
    conn = build_db(events)

    print("Creating Excel workbook ...")
    wb = Workbook()
    wb.remove(wb.active)   # remove default sheet

    print("   [1/8] Sheet 1: Overview KPIs")
    sheet_overview(wb, conn)
    print("   [2/8] Sheet 2: Funnel by Category")
    sheet_funnel(wb, conn)
    print("   [3/8] Sheet 3: Segmentation")
    sheet_segmentation(wb, conn)
    print("   [4/8] Sheet 4: Drop-off Analysis")
    sheet_dropoff(wb, conn)
    print("   [5/8] Sheet 5: A/B Test Framework")
    sheet_ab(wb, conn)
    print("   [6/8] Sheet 6: Recommendations")
    sheet_recommendations(wb, conn)
    print("   [7/8] Sheet 7: Raw Data (first 5,000 rows)")
    sheet_raw(wb, conn)
    print("   [8/8] Sheet 8: Daily Trend")
    sheet_daily(wb, conn)

    out = "output/ecommerce_dashboard.xlsx"
    wb.save(out)
    print(f"\n[DONE] Dashboard saved -> {out}")
    print("       Open it in Microsoft Excel for full interactivity (slicers, filters, charts).")
    conn.close()

if __name__ == "__main__":
    main()
