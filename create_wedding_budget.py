#!/usr/bin/env python3
"""Generate a wedding expenses workbook for a 2-day wedding (24–25 Feb)."""

from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    NamedStyle,
    PatternFill,
    Side,
)
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

wb = Workbook()

# --- Styles ---
thin = Border(
    left=Side(style="thin", color="D4C4B0"),
    right=Side(style="thin", color="D4C4B0"),
    top=Side(style="thin", color="D4C4B0"),
    bottom=Side(style="thin", color="D4C4B0"),
)
header_fill = PatternFill("solid", fgColor="2C3E50")
header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
title_font = Font(name="Calibri", bold=True, size=18, color="2C3E50")
subtitle_font = Font(name="Calibri", bold=True, size=13, color="5D6D7E")
section_fill = PatternFill("solid", fgColor="E8F0F5")
section_font = Font(name="Calibri", bold=True, size=11, color="1A5276")
money_font = Font(name="Calibri", size=11)
input_fill = PatternFill("solid", fgColor="FFF8E7")
actual_fill = PatternFill("solid", fgColor="E8F8F5")
suggest_fill = PatternFill("solid", fgColor="F5F5F5")
over_fill = PatternFill("solid", fgColor="FDEDEC")
under_fill = PatternFill("solid", fgColor="E8F8F5")
note_font = Font(name="Calibri", italic=True, size=10, color="7F8C8D")
label_font = Font(name="Calibri", bold=True, size=11, color="2C3E50")
total_fill = PatternFill("solid", fgColor="1A5276")
total_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
alt_row = PatternFill("solid", fgColor="F7FBFC")
white = PatternFill("solid", fgColor="FFFFFF")
amber = PatternFill("solid", fgColor="FCF3CF")
green_header = PatternFill("solid", fgColor="1E8449")
teal_header = PatternFill("solid", fgColor="148F77")

INR = '_("₹"* #,##0);_("₹"* (#,##0);_("₹"* "-"??_);_(@_)'


def style_header(ws, row, start_col, end_col, fill=header_fill):
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin


def money(cell):
    cell.number_format = INR
    cell.border = thin
    cell.alignment = Alignment(horizontal="right")


def text_cell(cell, value=None):
    if value is not None:
        cell.value = value
    cell.border = thin
    cell.alignment = Alignment(vertical="center", wrap_text=True)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ============================================================
# SHEET 1: Overview & Setup
# ============================================================
ws = wb.active
ws.title = "Overview"
ws.sheet_view.showGridLines = False

ws["A1"] = "Wedding Expense Planner"
ws["A1"].font = title_font
ws.merge_cells("A1:F1")

ws["A2"] = "2-day wedding · 24 & 25 February · Main event: 25 Feb"
ws["A2"].font = subtitle_font
ws.merge_cells("A2:F2")

ws["A4"] = "YOUR SETUP (edit yellow cells)"
ws["A4"].font = section_font
ws["A4"].fill = section_fill
ws.merge_cells("A4:C4")

setup = [
    ("Couple / Event name", "Our Wedding"),
    ("Wedding dates", "24–25 Feb"),
    ("Main event date", "25 Feb"),
    ("City / venue city", ""),
    ("Expected guest count (main day)", 120),
    ("Guests needing stay (nights)", 40),
    ("Nights of stay", 2),
    ("Currency", "INR (₹)"),
    ("Total budget target", 2500000),
]

for i, (label, val) in enumerate(setup, start=5):
    ws.cell(row=i, column=1, value=label).font = label_font
    ws.cell(row=i, column=1).border = thin
    c = ws.cell(row=i, column=2, value=val)
    c.fill = input_fill
    c.border = thin
    if isinstance(val, (int, float)) and "budget" in label.lower():
        money(c)
        c.fill = input_fill

ws["A15"] = "HOW TO USE THIS WORKBOOK"
ws["A15"].font = section_font
ws["A15"].fill = section_fill
ws.merge_cells("A15:F15")

instructions = [
    "1. Fill yellow cells on Overview (guest count, stay nights, total budget).",
    "2. Go to Budget Summary for a live roll-up of planned vs actual vs remaining.",
    "3. Enter quotes / deposits / paid amounts in Detailed Expenses (green = actuals).",
    "4. Use Food & Meals for per-meal catering lines (24th lunch/dinner, 25th breakfast/lunch/dinner).",
    "5. Use Guest Stay to list hotels/rooms and nightly costs.",
    "6. Track vendor contacts & payment schedule on Vendors and Payments sheets.",
    "7. Suggested amounts are mid-range Indian wedding estimates for ~100–150 guests — adjust to your city & style.",
    "8. Status column: Planned → Booked → Deposit paid → Fully paid → Cancelled.",
]

for i, line in enumerate(instructions, start=16):
    ws.cell(row=i, column=1, value=line).font = Font(name="Calibri", size=11)
    ws.merge_cells(f"A{i}:F{i}")

ws["A25"] = "SUGGESTED BUDGET SPLIT (typical 2-day wedding)"
ws["A25"].font = section_font
ws["A25"].fill = section_fill
ws.merge_cells("A25:D25")

headers = ["Category", "% of budget", "On ₹25L target", "Notes"]
for i, h in enumerate(headers, 1):
    ws.cell(row=26, column=i, value=h)
style_header(ws, 26, 1, 4)

splits = [
    ("Venue / mandap / décor", 0.18, "Hall + décor package; outdoor often higher"),
    ("Catering & all meals", 0.28, "Biggest line — includes 24th + 25th meals & teas"),
    ("Guest stay / accommodation", 0.15, "Rooms for outstation guests, 1–2 nights"),
    ("Outfits & jewellery (couple + family)", 0.12, "Bride/groom + parents; jewellery optional"),
    ("Photography & videography", 0.08, "Candid + traditional + reel/highlights"),
    ("Travel / cabs / logistics", 0.05, "Airport runs, guest shuttles, baraat"),
    ("Mehendi / sangeet / entertainment", 0.06, "Artist, DJ, sound, stage"),
    ("Invites, favours & misc", 0.03, "Cards, gifts for guests, stationery"),
    ("Makeup / hair / grooming", 0.03, "Bridal + family; trials"),
    ("Contingency buffer", 0.02, "Keep 5–10% if budget allows — highly recommended"),
]

for i, (cat, pct, note) in enumerate(splits, start=27):
    r = i
    text_cell(ws.cell(row=r, column=1), cat)
    c = ws.cell(row=r, column=2, value=pct)
    c.number_format = "0%"
    c.border = thin
    c = ws.cell(row=r, column=3, value=2500000 * pct)
    money(c)
    text_cell(ws.cell(row=r, column=4), note)
    if i % 2 == 0:
        for col in range(1, 5):
            if ws.cell(row=r, column=col).fill.fgColor is None or ws.cell(row=r, column=col).fill.fgColor.rgb == "00000000":
                ws.cell(row=r, column=col).fill = alt_row

ws["A38"] = "Tip: Change Total budget target above, then scale suggestions on Detailed Expenses."
ws["A38"].font = note_font

ws["A40"] = "QUICK SUGGESTIONS BY SCALE (main-day guests)"
ws["A40"].font = section_font
ws["A40"].fill = section_fill
ws.merge_cells("A40:E40")

scale_headers = ["Scale", "Guests", "Rough total", "Per guest (all-in)", "What fits"]
for i, h in enumerate(scale_headers, 1):
    ws.cell(row=41, column=i, value=h)
style_header(ws, 41, 1, 5, teal_header)

scales = [
    ("Intimate", "40–60", 800000, 15000, "Boutique venue, limited stay, focused catering"),
    ("Medium (default in this sheet)", "100–150", 2500000, 18000, "Hotel/banquet, guest rooms, full décor + photo"),
    ("Large", "200–300", 4500000, 18000, "Premium venue, more rooms, bigger entertainment"),
    ("Destination / luxury", "80–120", 6000000, 50000, "Resort buyout, designer outfits, high-end photo"),
]
for i, (s, g, tot, per, fit) in enumerate(scales, start=42):
    text_cell(ws.cell(row=i, column=1), s)
    text_cell(ws.cell(row=i, column=2), g)
    c = ws.cell(row=i, column=3, value=tot)
    money(c)
    c = ws.cell(row=i, column=4, value=per)
    money(c)
    text_cell(ws.cell(row=i, column=5), fit)

set_widths(ws, [42, 28, 18, 42, 48, 14])
ws.row_dimensions[1].height = 28
ws.row_dimensions[26].height = 22

# Named cells for Overview (for formulas)
ws["B9"].value = 120  # guests - already set
# B5 = name, B9 = guests, B10 = stay guests, B11 = nights, B13 = budget

# ============================================================
# SHEET 2: Budget Summary
# ============================================================
ws2 = wb.create_sheet("Budget Summary")
ws2.sheet_view.showGridLines = False

ws2["A1"] = "Budget Summary — live roll-up"
ws2["A1"].font = title_font
ws2.merge_cells("A1:G1")

ws2["A2"] = "Numbers pull from Detailed Expenses, Food & Meals, and Guest Stay. Update those sheets; this updates automatically."
ws2["A2"].font = note_font
ws2.merge_cells("A2:G2")

ws2["A4"] = "TARGET BUDGET"
ws2["A4"].font = label_font
ws2["B4"] = "=Overview!B13"
money(ws2["B4"])
ws2["B4"].fill = input_fill

ws2["A5"] = "TOTAL PLANNED"
ws2["A5"].font = label_font
ws2["B5"] = "=SUMIF('Detailed Expenses'!G:G,\"<>\",'Detailed Expenses'!E:E)+SUMIF('Food & Meals'!H:H,\"<>\",'Food & Meals'!F:F)+SUMIF('Guest Stay'!I:I,\"<>\",'Guest Stay'!G:G)"
# Simpler approach: sum planned columns with clean ranges
ws2["B5"] = "='Detailed Expenses'!E60+'Food & Meals'!F30+'Guest Stay'!G25"
money(ws2["B5"])

ws2["A6"] = "TOTAL ACTUAL / COMMITTED"
ws2["A6"].font = label_font
ws2["B6"] = "='Detailed Expenses'!F60+'Food & Meals'!G30+'Guest Stay'!H25"
money(ws2["B6"])

ws2["A7"] = "TOTAL PAID SO FAR"
ws2["A7"].font = label_font
ws2["B7"] = "='Detailed Expenses'!G60+'Food & Meals'!I30+'Guest Stay'!J25"
money(ws2["B7"])

ws2["A8"] = "REMAINING VS TARGET"
ws2["A8"].font = label_font
ws2["B8"] = "=B4-B5"
money(ws2["B8"])

ws2["A9"] = "STILL TO PAY (actual − paid)"
ws2["A9"].font = label_font
ws2["B9"] = "=B6-B7"
money(ws2["B9"])

ws2["A11"] = "BY CATEGORY"
ws2["A11"].font = section_font
ws2["A11"].fill = section_fill
ws2.merge_cells("A11:F11")

cat_headers = ["Category", "Planned (₹)", "Actual (₹)", "Paid (₹)", "Variance (Actual−Planned)", "% of planned"]
for i, h in enumerate(cat_headers, 1):
    ws2.cell(row=12, column=i, value=h)
style_header(ws2, 12, 1, 6)

# Categories matching Detailed Expenses section tags + Food + Stay
categories = [
    ("Venue & décor", "Venue & décor"),
    ("Outfits & jewellery", "Outfits & jewellery"),
    ("Photography & video", "Photography & video"),
    ("Travel & cabs", "Travel & cabs"),
    ("Entertainment & events", "Entertainment & events"),
    ("Beauty & grooming", "Beauty & grooming"),
    ("Invites & favours", "Invites & favours"),
    ("Misc & contingency", "Misc & contingency"),
    ("Catering & meals (Food sheet)", None),
    ("Guest stay (Stay sheet)", None),
]

# We'll use SUMIF on Detailed Expenses column A (Category group) 
# For food and stay use sheet totals

for i, (label, key) in enumerate(categories, start=13):
    text_cell(ws2.cell(row=i, column=1), label)
    if key:
        ws2.cell(row=i, column=2, value=f"=SUMIF('Detailed Expenses'!$A$5:$A$55,\"{key}\",'Detailed Expenses'!$E$5:$E$55)")
        ws2.cell(row=i, column=3, value=f"=SUMIF('Detailed Expenses'!$A$5:$A$55,\"{key}\",'Detailed Expenses'!$F$5:$F$55)")
        ws2.cell(row=i, column=4, value=f"=SUMIF('Detailed Expenses'!$A$5:$A$55,\"{key}\",'Detailed Expenses'!$G$5:$G$55)")
    elif "Catering" in label:
        ws2.cell(row=i, column=2, value="='Food & Meals'!F30")
        ws2.cell(row=i, column=3, value="='Food & Meals'!G30")
        ws2.cell(row=i, column=4, value="='Food & Meals'!I30")
    else:
        ws2.cell(row=i, column=2, value="='Guest Stay'!G25")
        ws2.cell(row=i, column=3, value="='Guest Stay'!H25")
        ws2.cell(row=i, column=4, value="='Guest Stay'!J25")
    for col in range(2, 5):
        money(ws2.cell(row=i, column=col))
    ws2.cell(row=i, column=5, value=f"=C{i}-B{i}")
    money(ws2.cell(row=i, column=5))
    ws2.cell(row=i, column=6, value=f"=IF(B{i}=0,\"\",C{i}/B{i})")
    ws2.cell(row=i, column=6).number_format = "0.0%"
    ws2.cell(row=i, column=6).border = thin

# Totals row
r = 23
ws2.cell(row=r, column=1, value="TOTAL").font = total_font
ws2.cell(row=r, column=1).fill = total_fill
for col, letter in enumerate(["B", "C", "D", "E"], 2):
    cell = ws2.cell(row=r, column=col, value=f"=SUM({letter}13:{letter}22)")
    money(cell)
    cell.fill = total_fill
    cell.font = total_font
ws2.cell(row=r, column=6).fill = total_fill
ws2.cell(row=r, column=6).border = thin

ws2["A25"] = "STATUS CHECKLIST"
ws2["A25"].font = section_font
ws2["A25"].fill = section_fill
ws2.merge_cells("A25:C25")

checks = [
    "Venue booked & deposit paid?",
    "Caterer menu locked for all meals?",
    "Guest rooms blocked (24 & 25 Feb)?",
    "Outfits ordered / fittings scheduled?",
    "Photographer & video booked?",
    "Cab/shuttle plan for both days?",
    "Makeup trial done?",
    "Contingency ≥ 5% of budget kept aside?",
]
for i, c in enumerate(checks, start=26):
    text_cell(ws2.cell(row=i, column=1), c)
    ws2.merge_cells(f"A{i}:B{i}")
    cell = ws2.cell(row=i, column=3, value="No")
    cell.fill = input_fill
    cell.border = thin

dv_yesno = DataValidation(type="list", formula1='"Yes,No,In progress"', allow_blank=True)
ws2.add_data_validation(dv_yesno)
dv_yesno.add("C26:C33")

# Pie chart data helper
ws2["H12"] = "For chart"
ws2["H12"].font = note_font
for i in range(13, 23):
    ws2.cell(row=i, column=8, value=f"=A{i}")
    ws2.cell(row=i, column=9, value=f"=B{i}")
    money(ws2.cell(row=i, column=9))

pie = PieChart()
pie.title = "Planned spend by category"
labels = Reference(ws2, min_col=8, min_row=13, max_row=22)
data = Reference(ws2, min_col=9, min_row=12, max_row=22)
pie.add_data(data, titles_from_data=True)
pie.set_categories(labels)
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showVal = False
pie.dataLabels.showCatName = False
pie.width = 18
pie.height = 12
ws2.add_chart(pie, "A35")

set_widths(ws2, [36, 16, 16, 14, 26, 14, 12, 36, 14])

# ============================================================
# SHEET 3: Detailed Expenses
# ============================================================
ws3 = wb.create_sheet("Detailed Expenses")
ws3.sheet_view.showGridLines = False

ws3["A1"] = "Detailed Expenses — venue, outfits, travel, photo & more"
ws3["A1"].font = title_font
ws3.merge_cells("A1:L1")

ws3["A2"] = "Yellow = your inputs. Suggested = mid-range for ~120 guests / medium Indian wedding. Actual = final quote or committed amount. Paid = money already transferred."
ws3["A2"].font = note_font
ws3.merge_cells("A2:L2")

det_headers = [
    "Category",
    "Item",
    "Suggestion / tip",
    "Suggested (₹)",
    "Planned (₹)",
    "Actual (₹)",
    "Paid (₹)",
    "Status",
    "Vendor",
    "Due date",
    "Day (24/25/Both)",
    "Notes",
]
for i, h in enumerate(det_headers, 1):
    ws3.cell(row=4, column=i, value=h)
style_header(ws3, 4, 1, 12)
ws3.row_dimensions[4].height = 30
ws3.freeze_panes = "A5"

# (category, item, tip, suggested)
items = [
    # Venue
    ("Venue & décor", "Wedding venue / banquet hall (25th)", "Book early for Feb; ask about outdoor + indoor backup", 250000),
    ("Venue & décor", "Mandap / stage / floral décor", "Package often cheaper than à la carte; confirm flower types", 180000),
    ("Venue & décor", "Lighting & draping", "Warm LEDs + fairy lights elevate photos", 45000),
    ("Venue & décor", "Entrance / photo booth setup", "Optional but guests love it", 25000),
    ("Venue & décor", "Generator / power backup", "Essential if outdoor or winter evening", 15000),
    ("Venue & décor", "Furniture / sofa / props rental", "For VIP seating & décor accents", 20000),
    # Outfits
    ("Outfits & jewellery", "Bride outfit(s) — ceremony + reception", "Budget 1–2 looks; reuse jewellery across days", 120000),
    ("Outfits & jewellery", "Groom outfit(s) — sherwani / suit", "Include footwear & accessories", 45000),
    ("Outfits & jewellery", "Family outfits (parents + siblings)", "Coordinate colours with décor palette", 80000),
    ("Outfits & jewellery", "Jewellery (own / rental / new)", "Rental can cut cost 40–60% for statement pieces", 100000),
    ("Outfits & jewellery", "Alterations & styling accessories", "Blouse, petticoat, safas, brooches", 15000),
    # Photo
    ("Photography & video", "Photography (candid + traditional)", "2-day coverage; check deliverables & timeline", 90000),
    ("Photography & video", "Cinematography / highlight film", "Often bundled; ask for raw footage access", 80000),
    ("Photography & video", "Drone / special coverage", "Venue permission required", 15000),
    ("Photography & video", "Photo album / prints", "Order after selecting favourites", 20000),
    # Travel
    ("Travel & cabs", "Airport / station guest pickups", "Spreadsheet of arrivals saves chaos", 25000),
    ("Travel & cabs", "Guest shuttle (hotel ↔ venue) both days", "Tempo traveller / mini-bus often best value", 35000),
    ("Travel & cabs", "Baraat / family procession vehicles", "Decorated car + support vehicles", 20000),
    ("Travel & cabs", "Couple / VIP car (24–25 Feb)", "Vintage or luxury — optional", 15000),
    ("Travel & cabs", "Outstation / intercity travel help", "Reimburse close family if needed", 30000),
    ("Travel & cabs", "Parking / valet", "Ask venue if included", 8000),
    # Entertainment
    ("Entertainment & events", "DJ / sound system (25th)", "Confirm power load with venue", 40000),
    ("Entertainment & events", "Sangeet / mehendi artist (24th)", "Dhol, singer, or playlist + host", 50000),
    ("Entertainment & events", "Emcee / host", "Keeps timeline on track", 20000),
    ("Entertainment & events", "Entertainment extras (fireworks, LED)", "Check local permits", 25000),
    # Beauty
    ("Beauty & grooming", "Bridal makeup & hair (trial + day)", "Trial 2–3 weeks before", 35000),
    ("Beauty & grooming", "Family makeup slots", "Bundle package for 4–6 people", 25000),
    ("Beauty & grooming", "Groom grooming / barber", "Day-of touch-up", 5000),
    # Invites
    ("Invites & favours", "Invitations (print + digital)", "Digital RSVP cuts waste & cost", 20000),
    ("Invites & favours", "Guest favours / return gifts", "Edible favours often most appreciated", 30000),
    ("Invites & favours", "Welcome kits for staying guests", "Water, snacks, schedule card, local treat", 15000),
    # Misc
    ("Misc & contingency", "Pooja / ceremony essentials", "Pandit, samagri, seating", 15000),
    ("Misc & contingency", "Wedding planner / coordinator day-of", "Huge stress-saver for 2-day events", 50000),
    ("Misc & contingency", "Tips & staff gratitude", "Venue, waiters, drivers, helpers", 15000),
    ("Misc & contingency", "Contingency buffer (5–10%)", "Do not spend unless needed", 75000),
    ("Misc & contingency", "Unexpected / overflow", "Last-minute guests, weather, extras", 20000),
]

dv_status = DataValidation(
    type="list",
    formula1='"Planned,Getting quotes,Booked,Deposit paid,Fully paid,Cancelled"',
    allow_blank=True,
)
ws3.add_data_validation(dv_status)

dv_day = DataValidation(type="list", formula1='"24 Feb,25 Feb,Both,Before"', allow_blank=True)
ws3.add_data_validation(dv_day)

for i, (cat, item, tip, suggested) in enumerate(items, start=5):
    text_cell(ws3.cell(row=i, column=1), cat)
    text_cell(ws3.cell(row=i, column=2), item)
    text_cell(ws3.cell(row=i, column=3), tip)
    c = ws3.cell(row=i, column=4, value=suggested)
    money(c)
    c.fill = suggest_fill
    # Planned defaults to suggested
    c = ws3.cell(row=i, column=5, value=suggested)
    money(c)
    c.fill = input_fill
    c = ws3.cell(row=i, column=6, value=0)
    money(c)
    c.fill = actual_fill
    c = ws3.cell(row=i, column=7, value=0)
    money(c)
    c.fill = actual_fill
    text_cell(ws3.cell(row=i, column=8), "Planned")
    ws3.cell(row=i, column=8).fill = input_fill
    text_cell(ws3.cell(row=i, column=9), "")
    ws3.cell(row=i, column=9).fill = input_fill
    text_cell(ws3.cell(row=i, column=10), "")
    ws3.cell(row=i, column=10).fill = input_fill
    # Day suggestion
    day = "Both"
    if "25th" in item or "ceremony" in item.lower() or "Bridal makeup" in item:
        day = "25 Feb"
    if "24th" in item or "Sangeet" in item or "mehendi" in item.lower():
        day = "24 Feb"
    if "trial" in tip.lower() or "Invitations" in item:
        day = "Before"
    text_cell(ws3.cell(row=i, column=11), day)
    ws3.cell(row=i, column=11).fill = input_fill
    text_cell(ws3.cell(row=i, column=12), "")
    ws3.cell(row=i, column=12).fill = input_fill
    if i % 2 == 0:
        ws3.cell(row=i, column=1).fill = alt_row
        ws3.cell(row=i, column=2).fill = alt_row
        ws3.cell(row=i, column=3).fill = alt_row

dv_status.add("H5:H55")
dv_day.add("K5:K55")

last_item_row = 4 + len(items)  # 40
total_row = 60
ws3.cell(row=total_row, column=1, value="TOTAL (this sheet)").font = total_font
ws3.cell(row=total_row, column=1).fill = total_fill
ws3.merge_cells(f"A{total_row}:C{total_row}")
for col in range(1, 4):
    ws3.cell(row=total_row, column=col).fill = total_fill
    ws3.cell(row=total_row, column=col).border = thin
for col, letter in [(4, "D"), (5, "E"), (6, "F"), (7, "G")]:
    cell = ws3.cell(row=total_row, column=col, value=f"=SUM({letter}5:{letter}{last_item_row})")
    money(cell)
    cell.fill = total_fill
    cell.font = total_font
for col in range(8, 13):
    ws3.cell(row=total_row, column=col).fill = total_fill
    ws3.cell(row=total_row, column=col).border = thin

ws3["A62"] = "Also budget food on 'Food & Meals' and hotels on 'Guest Stay' — they are separate so you can plan per meal / per room."
ws3["A62"].font = note_font

set_widths(ws3, [24, 42, 48, 14, 14, 14, 12, 14, 18, 12, 14, 28])

# ============================================================
# SHEET 4: Food & Meals
# ============================================================
ws4 = wb.create_sheet("Food & Meals")
ws4.sheet_view.showGridLines = False

ws4["A1"] = "Food & Catering — meal-by-meal (24 & 25 Feb)"
ws4["A1"].font = title_font
ws4.merge_cells("A1:J1")

ws4["A2"] = "Default headcounts use Overview guest count. Edit Planned headcount per meal (e.g. breakfast often fewer). Suggested rate ≈ ₹800–1,200/plate mid-range; luxury ₹1,500–2,500+."
ws4["A2"].font = note_font
ws4.merge_cells("A2:J2")

ws4["A3"] = "Main-day guest count (from Overview):"
ws4["A3"].font = label_font
ws4["B3"] = "=Overview!B9"
ws4["B3"].fill = suggest_fill
ws4["B3"].border = thin

food_headers = [
    "Day",
    "Meal",
    "Suggestion",
    "Headcount (planned)",
    "Rate / plate (₹)",
    "Planned total (₹)",
    "Actual total (₹)",
    "Vendor",
    "Paid (₹)",
    "Notes",
]
for i, h in enumerate(food_headers, 1):
    ws4.cell(row=5, column=i, value=h)
style_header(ws4, 5, 1, 10, green_header)
ws4.freeze_panes = "A6"

meals = [
    ("24 Feb", "Welcome drinks / high tea", "Light snacks on guest arrival; 60–80% of main count", 0.7, 350),
    ("24 Feb", "Lunch", "Comfort menu; many families do home-style or hotel buffet", 0.85, 900),
    ("24 Feb", "Evening snacks / mehendi bites", "Finger food during mehendi/sangeet", 0.9, 400),
    ("24 Feb", "Dinner", "Often the bigger 24th meal if sangeet follows", 1.0, 1100),
    ("24 Feb", "Late-night snacks", "Optional — chai, sandwiches after sangeet", 0.5, 250),
    ("25 Feb", "Breakfast", "For staying guests mainly — hotel often includes", 0.4, 450),
    ("25 Feb", "Mid-morning tea / snacks", "During prep / pooja; fruit + light bites", 0.5, 300),
    ("25 Feb", "Lunch", "Pre-ceremony or post-ceremony depending on timing", 0.95, 1000),
    ("25 Feb", "Wedding dinner / reception (MAIN)", "Highest plate cost — live counters, dessert, premium menu", 1.1, 1500),
    ("25 Feb", "Dessert / ice cream counter", "Sometimes separate from dinner package", 1.0, 200),
    ("Both", "Water, soft drinks & beverages", "Often underestimated — budget explicitly", 1.0, 150),
    ("Both", "Kids / special diet / Jain meals", "Confirm counts 1 week prior", 0.1, 900),
    ("Both", "Staff / vendor meals", "Photographers, makeup, band, helpers", 0.15, 400),
    ("Before", "Tasting session with caterer", "Worth every rupee — lock menu after tasting", 0.05, 0),
]

# tasting as flat suggested amount
for i, (day, meal, tip, hc_factor, rate) in enumerate(meals, start=6):
    text_cell(ws4.cell(row=i, column=1), day)
    text_cell(ws4.cell(row=i, column=2), meal)
    text_cell(ws4.cell(row=i, column=3), tip)
    if meal.startswith("Tasting"):
        hc = 8
        rate_val = 0
        planned_flat = 5000
        c = ws4.cell(row=i, column=4, value=hc)
        c.fill = input_fill
        c.border = thin
        c = ws4.cell(row=i, column=5, value=rate_val)
        money(c)
        c.fill = input_fill
        c = ws4.cell(row=i, column=6, value=planned_flat)
        money(c)
        c.fill = input_fill
    else:
        # headcount = round(Overview guests * factor)
        c = ws4.cell(row=i, column=4, value=f"=ROUND(Overview!B9*{hc_factor},0)")
        c.fill = input_fill
        c.border = thin
        c = ws4.cell(row=i, column=5, value=rate)
        money(c)
        c.fill = input_fill
        c = ws4.cell(row=i, column=6, value=f"=D{i}*E{i}")
        money(c)
        c.fill = suggest_fill
    c = ws4.cell(row=i, column=7, value=0)
    money(c)
    c.fill = actual_fill
    text_cell(ws4.cell(row=i, column=8), "")
    ws4.cell(row=i, column=8).fill = input_fill
    c = ws4.cell(row=i, column=9, value=0)
    money(c)
    c.fill = actual_fill
    text_cell(ws4.cell(row=i, column=10), "")
    ws4.cell(row=i, column=10).fill = input_fill

food_last = 5 + len(meals)  # 19

# Extra blank rows for custom meals
for i in range(food_last + 1, 29):
    for col in range(1, 11):
        ws4.cell(row=i, column=col).border = thin
        if col in (4, 5, 7, 8, 9, 10):
            ws4.cell(row=i, column=col).fill = input_fill
        if col == 6:
            ws4.cell(row=i, column=col, value=f"=IF(OR(D{i}=\"\",E{i}=\"\"),\"\",D{i}*E{i})")
            money(ws4.cell(row=i, column=col))
        if col in (5, 7, 9):
            money(ws4.cell(row=i, column=col))

# Total row 30
ws4.cell(row=30, column=1, value="TOTAL FOOD & CATERING").font = total_font
ws4.cell(row=30, column=1).fill = total_fill
ws4.merge_cells("A30:E30")
for col in range(1, 6):
    ws4.cell(row=30, column=col).fill = total_fill
    ws4.cell(row=30, column=col).border = thin
for col, letter in [(6, "F"), (7, "G"), (9, "I")]:
    cell = ws4.cell(row=30, column=col, value=f"=SUM({letter}6:{letter}28)")
    money(cell)
    cell.fill = total_fill
    cell.font = total_font
ws4.cell(row=30, column=8).fill = total_fill
ws4.cell(row=30, column=8).border = thin
ws4.cell(row=30, column=10).fill = total_fill
ws4.cell(row=30, column=10).border = thin

ws4["A32"] = "MEAL PLANNING TIPS"
ws4["A32"].font = section_font
ws4["A32"].fill = section_fill
ws4.merge_cells("A32:D32")

tips = [
    "Lock final headcounts 7–10 days before; most caterers charge on guaranteed minimum.",
    "25th dinner is usually 10–20% higher headcount than RSVPs (plus-ones & drop-ins).",
    "Breakfast is often included in hotel stay — don't double-pay; mark Notes accordingly.",
    "Ask for a clear split: veg / non-veg / Jain / kids plates and live-counter charges.",
    "Confirm whether taxes (GST), service charge, and crockery/service staff are included.",
    "For 24th lunch: if many guests arrive afternoon-only, reduce lunch and boost dinner.",
]
for i, t in enumerate(tips, start=33):
    ws4.cell(row=i, column=1, value=t).font = Font(name="Calibri", size=11)
    ws4.merge_cells(f"A{i}:J{i}")

# Day subtotals
ws4["A40"] = "DAY SUBTOTALS"
ws4["A40"].font = section_font
ws4["A40"].fill = section_fill

ws4["A41"] = "24 Feb food (planned)"
ws4["B41"] = '=SUMIF(A6:A28,"24 Feb",F6:F28)'
money(ws4["B41"])
ws4["A42"] = "25 Feb food (planned)"
ws4["B42"] = '=SUMIF(A6:A28,"25 Feb",F6:F28)'
money(ws4["B42"])
ws4["A43"] = "Both / other (planned)"
ws4["B43"] = '=SUMIF(A6:A28,"Both",F6:F28)+SUMIF(A6:A28,"Before",F6:F28)'
money(ws4["B43"])

set_widths(ws4, [12, 36, 55, 18, 16, 16, 16, 18, 12, 28])

# ============================================================
# SHEET 5: Guest Stay
# ============================================================
ws5 = wb.create_sheet("Guest Stay")
ws5.sheet_view.showGridLines = False

ws5["A1"] = "Guest Stay — rooms for 24 & 25 February"
ws5["A1"].font = title_font
ws5.merge_cells("A1:K1")

ws5["A2"] = "Block rooms early for Feb weekends. Negotiate wedding attrition (free rooms per banquet spend). Breakfast inclusion can offset Food sheet breakfast line."
ws5["A2"].font = note_font
ws5.merge_cells("A2:K2")

ws5["A3"] = "Guests needing stay:"
ws5["B3"] = "=Overview!B10"
ws5["B3"].border = thin
ws5["B3"].fill = suggest_fill
ws5["C3"] = "Nights:"
ws5["D3"] = "=Overview!B11"
ws5["D3"].border = thin
ws5["D3"].fill = suggest_fill

stay_headers = [
    "Hotel / stay",
    "Room type",
    "# Rooms",
    "Check-in",
    "Check-out",
    "Rate / room / night (₹)",
    "Planned total (₹)",
    "Actual total (₹)",
    "Guest names / block",
    "Paid (₹)",
    "Notes",
]
for i, h in enumerate(stay_headers, 1):
    ws5.cell(row=5, column=i, value=h)
style_header(ws5, 5, 1, 11, teal_header)
ws5.freeze_panes = "A6"

stay_rows = [
    ("Primary hotel (guest block)", "Deluxe double", 12, "24 Feb", "26 Feb", 7500, "Main outstation guests"),
    ("Primary hotel (guest block)", "Suite / family", 3, "24 Feb", "26 Feb", 14000, "Elderly / close family"),
    ("Secondary hotel / overflow", "Standard double", 5, "24 Feb", "26 Feb", 5500, "If primary fills up"),
    ("Couple suite / wedding night", "Suite", 1, "25 Feb", "26 Feb", 18000, "Optional upgrade"),
    ("Early arrivals", "Deluxe double", 2, "23 Feb", "26 Feb", 7500, "Decor / close family setup"),
    ("Driver / staff rooms", "Economy", 2, "24 Feb", "26 Feb", 2500, "Often forgotten — budget it"),
]

for i, (hotel, rtype, nrooms, cin, cout, rate, notes) in enumerate(stay_rows, start=6):
    text_cell(ws5.cell(row=i, column=1), hotel)
    ws5.cell(row=i, column=1).fill = input_fill
    text_cell(ws5.cell(row=i, column=2), rtype)
    ws5.cell(row=i, column=2).fill = input_fill
    c = ws5.cell(row=i, column=3, value=nrooms)
    c.fill = input_fill
    c.border = thin
    text_cell(ws5.cell(row=i, column=4), cin)
    ws5.cell(row=i, column=4).fill = input_fill
    text_cell(ws5.cell(row=i, column=5), cout)
    ws5.cell(row=i, column=5).fill = input_fill
    c = ws5.cell(row=i, column=6, value=rate)
    money(c)
    c.fill = input_fill
    # nights ≈ 2 for most; use Overview nights for formula simplicity: rooms * rate * nights
    # For early arrivals 3 nights — hardcode factor in planned
    nights = 2
    if cin == "23 Feb":
        nights = 3
    if cin == "25 Feb":
        nights = 1
    c = ws5.cell(row=i, column=7, value=f"=C{i}*F{i}*{nights}")
    money(c)
    c.fill = suggest_fill
    c = ws5.cell(row=i, column=8, value=0)
    money(c)
    c.fill = actual_fill
    text_cell(ws5.cell(row=i, column=9), "")
    ws5.cell(row=i, column=9).fill = input_fill
    c = ws5.cell(row=i, column=10, value=0)
    money(c)
    c.fill = actual_fill
    text_cell(ws5.cell(row=i, column=11), notes)
    ws5.cell(row=i, column=11).fill = input_fill

for i in range(12, 24):
    for col in range(1, 12):
        ws5.cell(row=i, column=col).border = thin
        if col in (1, 2, 3, 4, 5, 6, 9, 11):
            ws5.cell(row=i, column=col).fill = input_fill
        if col == 7:
            # assume 2 nights if user fills rooms & rate
            ws5.cell(row=i, column=col, value=f'=IF(OR(C{i}="",F{i}=""),"",C{i}*F{i}*Overview!B11)')
            money(ws5.cell(row=i, column=col))
        if col in (6, 8, 10):
            money(ws5.cell(row=i, column=col))
            if col in (8, 10):
                ws5.cell(row=i, column=col).fill = actual_fill

ws5.cell(row=25, column=1, value="TOTAL GUEST STAY").font = total_font
ws5.cell(row=25, column=1).fill = total_fill
ws5.merge_cells("A25:F25")
for col in range(1, 7):
    ws5.cell(row=25, column=col).fill = total_fill
    ws5.cell(row=25, column=col).border = thin
for col, letter in [(7, "G"), (8, "H"), (10, "J")]:
    cell = ws5.cell(row=25, column=col, value=f"=SUM({letter}6:{letter}23)")
    money(cell)
    cell.fill = total_fill
    cell.font = total_font
for col in (9, 11):
    ws5.cell(row=25, column=col).fill = total_fill
    ws5.cell(row=25, column=col).border = thin

ws5["A27"] = "STAY COST CHEAT SHEET (suggested)"
ws5["A27"].font = section_font
ws5["A27"].fill = section_fill
ws5.merge_cells("A27:D27")

cheat = [
    ("Budget hotel", "₹3,000–5,000 / room / night", "Good for overflow & staff"),
    ("Mid-range 3–4★", "₹6,000–10,000 / room / night", "Best balance for most guest blocks"),
    ("Premium 5★", "₹12,000–25,000+ / room / night", "Close family & couple"),
    ("Homestay / Airbnb cluster", "₹2,500–6,000 / night", "Works if coordinator manages keys & transport"),
]
for i, h in enumerate(["Type", "Typical rate", "When to use"], 1):
    ws5.cell(row=28, column=i, value=h)
style_header(ws5, 28, 1, 3)
for i, (a, b, c) in enumerate(cheat, start=29):
    text_cell(ws5.cell(row=i, column=1), a)
    text_cell(ws5.cell(row=i, column=2), b)
    text_cell(ws5.cell(row=i, column=3), c)

ws5["A34"] = "Rough room need ≈ guests needing stay ÷ 2 (doubles) + a few singles/suites. Example: 40 guests ≈ 18–22 rooms."
ws5["A34"].font = note_font
ws5.merge_cells("A34:K34")

set_widths(ws5, [28, 16, 10, 12, 12, 20, 16, 16, 22, 12, 36])

# ============================================================
# SHEET 6: Vendors
# ============================================================
ws6 = wb.create_sheet("Vendors")
ws6.sheet_view.showGridLines = False

ws6["A1"] = "Vendor directory"
ws6["A1"].font = title_font

vend_headers = [
    "Category",
    "Vendor name",
    "Contact person",
    "Phone",
    "Email / Instagram",
    "Quote (₹)",
    "Advance %",
    "Contract signed?",
    "Priority",
    "Notes",
]
for i, h in enumerate(vend_headers, 1):
    ws6.cell(row=3, column=i, value=h)
style_header(ws6, 3, 1, 10)

vendor_cats = [
    "Venue",
    "Caterer",
    "Decorator / florist",
    "Hotel (guest stay)",
    "Photographer",
    "Videographer",
    "Makeup artist",
    "DJ / entertainment",
    "Cab / travel company",
    "Outfit designer — bride",
    "Outfit designer — groom",
    "Planner / coordinator",
    "Pandit / ceremony",
    "Invites / favours",
    "Other",
]

dv_vcat = DataValidation(type="list", formula1='"' + ",".join(vendor_cats) + '"', allow_blank=True)
ws6.add_data_validation(dv_vcat)
dv_vcat.add("A4:A30")

dv_prio = DataValidation(type="list", formula1='"Must-book now,Soon,Flexible"', allow_blank=True)
ws6.add_data_validation(dv_prio)
dv_prio.add("I4:I30")

dv_contract = DataValidation(type="list", formula1='"Yes,No,Sent"', allow_blank=True)
ws6.add_data_validation(dv_contract)
dv_contract.add("H4:H30")

prefill = [
    ("Venue", "Must-book now"),
    ("Caterer", "Must-book now"),
    ("Hotel (guest stay)", "Must-book now"),
    ("Photographer", "Must-book now"),
    ("Decorator / florist", "Soon"),
    ("Makeup artist", "Soon"),
    ("Cab / travel company", "Soon"),
    ("DJ / entertainment", "Soon"),
    ("Outfit designer — bride", "Soon"),
    ("Outfit designer — groom", "Flexible"),
]

for i, (cat, prio) in enumerate(prefill, start=4):
    text_cell(ws6.cell(row=i, column=1), cat)
    for col in range(2, 11):
        ws6.cell(row=i, column=col).border = thin
        ws6.cell(row=i, column=col).fill = input_fill
    text_cell(ws6.cell(row=i, column=9), prio)
    ws6.cell(row=i, column=9).fill = input_fill
    if col := 6:
        money(ws6.cell(row=i, column=6))
        ws6.cell(row=i, column=6).fill = input_fill

for i in range(14, 31):
    for col in range(1, 11):
        ws6.cell(row=i, column=col).border = thin
        ws6.cell(row=i, column=col).fill = input_fill
    money(ws6.cell(row=i, column=6))
    ws6.cell(row=i, column=6).fill = input_fill

set_widths(ws6, [22, 22, 16, 14, 24, 12, 12, 14, 14, 28])

# ============================================================
# SHEET 7: Payments
# ============================================================
ws7 = wb.create_sheet("Payments")
ws7.sheet_view.showGridLines = False

ws7["A1"] = "Payment schedule & tracker"
ws7["A1"].font = title_font
ws7.merge_cells("A1:I1")

ws7["A2"] = "Typical pattern: 20–40% advance on booking → 30–40% mid (Jan/early Feb) → balance 2–3 days before or on event day."
ws7["A2"].font = note_font

pay_headers = [
    "Vendor / item",
    "Total due (₹)",
    "Advance due",
    "Advance paid?",
    "2nd instalment",
    "2nd paid?",
    "Final balance",
    "Final paid?",
    "Payment mode / ref",
]
for i, h in enumerate(pay_headers, 1):
    ws7.cell(row=4, column=i, value=h)
style_header(ws7, 4, 1, 9)

sample_pays = [
    ("Venue", 250000, 100000, 100000, 50000),
    ("Caterer (all meals)", 400000, 100000, 150000, 150000),
    ("Guest hotel block", 220000, 50000, 100000, 70000),
    ("Decorator", 180000, 60000, 60000, 60000),
    ("Photographer + video", 170000, 50000, 60000, 60000),
    ("Makeup", 60000, 20000, 0, 40000),
    ("Travel / cabs", 133000, 30000, 50000, 53000),
    ("Outfits (balance)", 200000, 80000, 70000, 50000),
]

dv_paid = DataValidation(type="list", formula1='"Yes,No,Partial"', allow_blank=True)
ws7.add_data_validation(dv_paid)
dv_paid.add("D5:D40")
dv_paid.add("F5:F40")
dv_paid.add("H5:H40")

for i, (name, total, adv, mid, final) in enumerate(sample_pays, start=5):
    text_cell(ws7.cell(row=i, column=1), name)
    ws7.cell(row=i, column=1).fill = input_fill
    c = ws7.cell(row=i, column=2, value=total)
    money(c)
    c.fill = input_fill
    c = ws7.cell(row=i, column=3, value=adv)
    money(c)
    c.fill = input_fill
    text_cell(ws7.cell(row=i, column=4), "No")
    ws7.cell(row=i, column=4).fill = input_fill
    c = ws7.cell(row=i, column=5, value=mid)
    money(c)
    c.fill = input_fill
    text_cell(ws7.cell(row=i, column=6), "No")
    ws7.cell(row=i, column=6).fill = input_fill
    c = ws7.cell(row=i, column=7, value=final)
    money(c)
    c.fill = input_fill
    text_cell(ws7.cell(row=i, column=8), "No")
    ws7.cell(row=i, column=8).fill = input_fill
    text_cell(ws7.cell(row=i, column=9), "")
    ws7.cell(row=i, column=9).fill = input_fill

for i in range(13, 25):
    for col in range(1, 10):
        ws7.cell(row=i, column=col).border = thin
        ws7.cell(row=i, column=col).fill = input_fill
    for col in (2, 3, 5, 7):
        money(ws7.cell(row=i, column=col))
        ws7.cell(row=i, column=col).fill = input_fill

ws7.cell(row=26, column=1, value="TOTALS").font = total_font
ws7.cell(row=26, column=1).fill = total_fill
for col, letter in [(2, "B"), (3, "C"), (5, "E"), (7, "G")]:
    cell = ws7.cell(row=26, column=col, value=f"=SUM({letter}5:{letter}24)")
    money(cell)
    cell.fill = total_fill
    cell.font = total_font
for col in (4, 6, 8, 9):
    ws7.cell(row=26, column=col).fill = total_fill
    ws7.cell(row=26, column=col).border = thin

ws7["A28"] = "Suggested calendar"
ws7["A28"].font = section_font
ws7["A28"].fill = section_fill
ws7.merge_cells("A28:C28")

cal = [
    ("Now → Dec", "Book venue, caterer, hotel block, photographer — pay advances"),
    ("January", "Lock menus, décor concept, outfits fittings — 2nd instalments"),
    ("1–10 Feb", "Final headcounts, rooming list, cab schedule"),
    ("15–20 Feb", "Final vendor balances where required; confirm all ETAs"),
    ("24–25 Feb", "Day-of tips & any on-site balances; keep emergency cash/UPI"),
]
for i, (when, what) in enumerate(cal, start=29):
    text_cell(ws7.cell(row=i, column=1), when)
    text_cell(ws7.cell(row=i, column=2), what)
    ws7.merge_cells(f"B{i}:E{i}")

set_widths(ws7, [28, 14, 14, 14, 14, 12, 14, 12, 20])

# ============================================================
# SHEET 8: Guest count helper
# ============================================================
ws8 = wb.create_sheet("Guest List Helper")
ws8.sheet_view.showGridLines = False

ws8["A1"] = "Guest list helper — stay & meal flags"
ws8["A1"].font = title_font
ws8.merge_cells("A1:H1")

ws8["A2"] = "Optional: list guests to estimate stay rooms and per-meal headcounts. Totals at the bottom feed your planning (copy into Overview if you want)."
ws8["A2"].font = note_font

gh = ["Name", "Side (Bride/Groom)", "City", "Needs stay?", "Nights", "24 Lunch", "24 Dinner", "25 Breakfast", "25 Lunch", "25 Dinner", "Notes"]
for i, h in enumerate(gh, 1):
    ws8.cell(row=4, column=i, value=h)
style_header(ws8, 4, 1, 11)

dv_side = DataValidation(type="list", formula1='"Bride,Groom,Both"', allow_blank=True)
ws8.add_data_validation(dv_side)
dv_side.add("B5:B104")

dv_yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
ws8.add_data_validation(dv_yn)
dv_yn.add("D5:D104")
for col in ["F", "G", "H", "I", "J"]:
    dv_yn.add(f"{col}5:{col}104")

for i in range(5, 55):
    for col in range(1, 12):
        ws8.cell(row=i, column=col).border = thin
        ws8.cell(row=i, column=col).fill = input_fill

ws8["A56"] = "COUNTS"
ws8["A56"].font = section_font
ws8["A56"].fill = section_fill

ws8["A57"] = "Guests needing stay"
ws8["B57"] = '=COUNTIF(D5:D54,"Yes")'
ws8["A58"] = "24 Lunch RSVPs"
ws8["B58"] = '=COUNTIF(F5:F54,"Yes")'
ws8["A59"] = "24 Dinner RSVPs"
ws8["B59"] = '=COUNTIF(G5:G54,"Yes")'
ws8["A60"] = "25 Breakfast RSVPs"
ws8["B60"] = '=COUNTIF(H5:H54,"Yes")'
ws8["A61"] = "25 Lunch RSVPs"
ws8["B61"] = '=COUNTIF(I5:I54,"Yes")'
ws8["A62"] = "25 Dinner RSVPs"
ws8["B62"] = '=COUNTIF(J5:J54,"Yes")'

for r in range(57, 63):
    ws8.cell(row=r, column=1).font = label_font
    ws8.cell(row=r, column=2).border = thin
    ws8.cell(row=r, column=2).fill = suggest_fill

set_widths(ws8, [22, 14, 14, 12, 10, 12, 12, 14, 12, 12, 24])

# ============================================================
# Fix Budget Summary formulas to be cleaner - already using sheet totals
# Recalculate Overview suggested based on note
# ============================================================

# Add a Assumptions sheet briefly? Skip - Overview is enough.

out = "/workspace/Wedding_Expenses_24-25_Feb.xlsx"
wb.save(out)
print(f"Saved {out}")
print(f"Sheets: {wb.sheetnames}")
