#!/usr/bin/env python3
"""Generate a blank wedding expenses workbook for a 2-day wedding (24–25 Feb)."""

from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

wb = Workbook()

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
input_fill = PatternFill("solid", fgColor="FFF8E7")
actual_fill = PatternFill("solid", fgColor="E8F8F5")
note_font = Font(name="Calibri", italic=True, size=10, color="7F8C8D")
label_font = Font(name="Calibri", bold=True, size=11, color="2C3E50")
total_fill = PatternFill("solid", fgColor="1A5276")
total_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
alt_row = PatternFill("solid", fgColor="F7FBFC")
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


def blank_money(cell, fill=input_fill):
    cell.value = None
    money(cell)
    cell.fill = fill


def text_cell(cell, value=None):
    if value is not None:
        cell.value = value
    cell.border = thin
    cell.alignment = Alignment(vertical="center", wrap_text=True)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def blank_input(cell, value=None):
    cell.value = value
    cell.fill = input_fill
    cell.border = thin
    cell.alignment = Alignment(vertical="center", wrap_text=True)


# ============================================================
# SHEET 1: Overview
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

ws["A4"] = "YOUR SETUP (fill yellow cells — all amounts left blank for you)"
ws["A4"].font = section_font
ws["A4"].fill = section_fill
ws.merge_cells("A4:C4")

setup = [
    ("Couple / Event name", None),
    ("Wedding dates", "24–25 Feb"),
    ("Main event date", "25 Feb"),
    ("City / venue city", None),
    ("Expected guest count (main day)", None),
    ("Guests needing stay", None),
    ("Nights of stay", None),
    ("Currency", None),
    ("Total budget target", None),
]

for i, (label, val) in enumerate(setup, start=5):
    ws.cell(row=i, column=1, value=label).font = label_font
    ws.cell(row=i, column=1).border = thin
    c = ws.cell(row=i, column=2)
    blank_input(c, val)
    if label == "Total budget target":
        money(c)
        c.fill = input_fill
        c.value = None

ws["A15"] = "HOW TO USE THIS WORKBOOK"
ws["A15"].font = section_font
ws["A15"].fill = section_fill
ws.merge_cells("A15:F15")

instructions = [
    "1. Fill yellow cells on Overview (guest count, stay nights, total budget, currency).",
    "2. Enter your own Planned / Actual / Paid amounts on Detailed Expenses, Food & Meals, and Guest Stay.",
    "3. Budget Summary rolls up those sheets automatically (blank until you enter numbers).",
    "4. Use Vendors and Payments to track contacts and instalments.",
    "5. Guest List Helper is optional — mark who needs stay and which meals.",
    "6. Category rows and tip text are suggestions only — no amounts are pre-filled.",
]

for i, line in enumerate(instructions, start=16):
    ws.cell(row=i, column=1, value=line).font = Font(name="Calibri", size=11)
    ws.merge_cells(f"A{i}:F{i}")

ws["A23"] = "SUGGESTED CATEGORIES TO BUDGET FOR (fill your own amounts on other sheets)"
ws["A23"].font = section_font
ws["A23"].fill = section_fill
ws.merge_cells("A23:D23")

headers = ["Category", "Why it matters", "Things to confirm", "Your notes"]
for i, h in enumerate(headers, 1):
    ws.cell(row=24, column=i, value=h)
style_header(ws, 24, 1, 4)

splits = [
    (
        "Venue / mandap / décor",
        "Usually one of the largest fixed costs",
        "Indoor backup, power, timing, what’s included in package",
    ),
    (
        "Catering & all meals",
        "Biggest variable cost across a 2-day wedding",
        "Per-meal headcount, taxes, service staff, Jain/kids plates",
    ),
    (
        "Guest stay / accommodation",
        "Outstation guests for 24 & 25 Feb",
        "Breakfast included?, attrition/free rooms, check-in dates",
    ),
    (
        "Outfits & jewellery",
        "Couple + close family looks",
        "Fittings timeline, alterations, rental vs buy",
    ),
    (
        "Photography & videography",
        "Covers both days + deliverables",
        "Hours, album, highlight film, raw footage, overtime",
    ),
    (
        "Travel / cabs / logistics",
        "Airport runs, hotel↔venue shuttles, baraat",
        "Vehicle count, wait-time charges, parking/valet",
    ),
    (
        "Mehendi / sangeet / entertainment",
        "24th energy + 25th reception vibe",
        "DJ/sound load, artist timing, permits",
    ),
    (
        "Invites, favours & misc",
        "Easy to underestimate",
        "Digital vs print RSVP, welcome kits, return gifts",
    ),
    (
        "Makeup / hair / grooming",
        "Trial + event-day slots",
        "Family package, touch-ups, travel to venue",
    ),
    (
        "Contingency buffer",
        "Last-minute guests, weather, extras",
        "Keep a set-aside you don’t spend unless needed",
    ),
]

for i, (cat, why, confirm) in enumerate(splits, start=25):
    text_cell(ws.cell(row=i, column=1), cat)
    text_cell(ws.cell(row=i, column=2), why)
    text_cell(ws.cell(row=i, column=3), confirm)
    blank_input(ws.cell(row=i, column=4))
    if i % 2 == 0:
        for col in range(1, 4):
            ws.cell(row=i, column=col).fill = alt_row

ws["A36"] = "MEALS TO PLAN (enter costs on Food & Meals sheet)"
ws["A36"].font = section_font
ws["A36"].fill = section_fill
ws.merge_cells("A36:C36")

meal_reminders = [
    "24 Feb — welcome/high tea, lunch, evening snacks, dinner (+ optional late-night)",
    "25 Feb — breakfast, mid-morning tea, lunch, main wedding dinner/reception, dessert counter",
    "Both days — beverages, special-diet plates, staff/vendor meals",
]
for i, line in enumerate(meal_reminders, start=37):
    ws.cell(row=i, column=1, value=line).font = Font(name="Calibri", size=11)
    ws.merge_cells(f"A{i}:F{i}")

ws["A41"] = "All money columns elsewhere start blank. Totals and charts populate as you type."
ws["A41"].font = note_font

set_widths(ws, [36, 42, 48, 28, 14, 14])
ws.row_dimensions[1].height = 28
ws.row_dimensions[24].height = 22

# ============================================================
# SHEET 2: Budget Summary
# ============================================================
ws2 = wb.create_sheet("Budget Summary")
ws2.sheet_view.showGridLines = False

ws2["A1"] = "Budget Summary — live roll-up"
ws2["A1"].font = title_font
ws2.merge_cells("A1:G1")

ws2["A2"] = "Pulls from Detailed Expenses, Food & Meals, and Guest Stay. Stays blank/zero until you enter amounts."
ws2["A2"].font = note_font
ws2.merge_cells("A2:G2")

ws2["A4"] = "TARGET BUDGET"
ws2["A4"].font = label_font
ws2["B4"] = "=Overview!B13"
money(ws2["B4"])
ws2["B4"].fill = input_fill

ws2["A5"] = "TOTAL PLANNED"
ws2["A5"].font = label_font
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
ws2["B8"] = "=IF(OR(B4=\"\",B5=\"\"),\"\",B4-B5)"
money(ws2["B8"])

ws2["A9"] = "STILL TO PAY (actual − paid)"
ws2["A9"].font = label_font
ws2["B9"] = "=IF(OR(B6=\"\",B7=\"\"),\"\",B6-B7)"
money(ws2["B9"])

ws2["A11"] = "BY CATEGORY"
ws2["A11"].font = section_font
ws2["A11"].fill = section_fill
ws2.merge_cells("A11:F11")

cat_headers = [
    "Category",
    "Planned",
    "Actual",
    "Paid",
    "Variance (Actual−Planned)",
    "% of planned",
]
for i, h in enumerate(cat_headers, 1):
    ws2.cell(row=12, column=i, value=h)
style_header(ws2, 12, 1, 6)

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

for i, (label, key) in enumerate(categories, start=13):
    text_cell(ws2.cell(row=i, column=1), label)
    if key:
        ws2.cell(
            row=i,
            column=2,
            value=f"=SUMIF('Detailed Expenses'!$A$5:$A$55,\"{key}\",'Detailed Expenses'!$E$5:$E$55)",
        )
        ws2.cell(
            row=i,
            column=3,
            value=f"=SUMIF('Detailed Expenses'!$A$5:$A$55,\"{key}\",'Detailed Expenses'!$F$5:$F$55)",
        )
        ws2.cell(
            row=i,
            column=4,
            value=f"=SUMIF('Detailed Expenses'!$A$5:$A$55,\"{key}\",'Detailed Expenses'!$G$5:$G$55)",
        )
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
    ws2.cell(row=i, column=5, value=f"=IF(AND(B{i}=0,C{i}=0),\"\",C{i}-B{i})")
    money(ws2.cell(row=i, column=5))
    ws2.cell(row=i, column=6, value=f"=IF(B{i}=0,\"\",C{i}/B{i})")
    ws2.cell(row=i, column=6).number_format = "0.0%"
    ws2.cell(row=i, column=6).border = thin

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
    "Contingency set aside?",
]
for i, c in enumerate(checks, start=26):
    text_cell(ws2.cell(row=i, column=1), c)
    ws2.merge_cells(f"A{i}:B{i}")
    blank_input(ws2.cell(row=i, column=3))

dv_yesno = DataValidation(type="list", formula1='"Yes,No,In progress"', allow_blank=True)
ws2.add_data_validation(dv_yesno)
dv_yesno.add("C26:C33")

for i in range(13, 23):
    ws2.cell(row=i, column=8, value=f"=A{i}")
    ws2.cell(row=i, column=9, value=f"=B{i}")
    money(ws2.cell(row=i, column=9))

pie = PieChart()
pie.title = "Planned spend by category"
labels = Reference(ws2, min_col=8, min_row=13, max_row=22)
data = Reference(ws2, min_col=9, min_row=12, max_row=22)
ws2.cell(row=12, column=9, value="Planned")
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
ws3.merge_cells("A1:K1")

ws3["A2"] = "Yellow = planned / vendor / dates. Green = actual & paid. All amount cells start blank."
ws3["A2"].font = note_font
ws3.merge_cells("A2:K2")

det_headers = [
    "Category",
    "Item",
    "Suggestion / tip",
    "Planned",
    "Actual",
    "Paid",
    "Status",
    "Vendor",
    "Due date",
    "Day (24/25/Both)",
    "Notes",
]
for i, h in enumerate(det_headers, 1):
    ws3.cell(row=4, column=i, value=h)
style_header(ws3, 4, 1, 11)
ws3.row_dimensions[4].height = 30
ws3.freeze_panes = "A5"

# (category, item, tip, day)
items = [
    ("Venue & décor", "Wedding venue / banquet hall (25th)", "Ask about outdoor + indoor backup", "25 Feb"),
    ("Venue & décor", "Mandap / stage / floral décor", "Package vs à la carte; confirm flower types", "Both"),
    ("Venue & décor", "Lighting & draping", "Warm lighting helps photos", "Both"),
    ("Venue & décor", "Entrance / photo booth setup", "Optional", "25 Feb"),
    ("Venue & décor", "Generator / power backup", "Useful for outdoor or evening events", "Both"),
    ("Venue & décor", "Furniture / sofa / props rental", "VIP seating & décor accents", "Both"),
    ("Outfits & jewellery", "Bride outfit(s) — ceremony + reception", "Plan fittings early", "Before"),
    ("Outfits & jewellery", "Groom outfit(s) — sherwani / suit", "Include footwear & accessories", "Before"),
    ("Outfits & jewellery", "Family outfits (parents + siblings)", "Coordinate with décor colours", "Before"),
    ("Outfits & jewellery", "Jewellery (own / rental / new)", "Rental can reduce cost for statement pieces", "Before"),
    ("Outfits & jewellery", "Alterations & styling accessories", "Blouse, petticoat, safas, brooches", "Before"),
    ("Photography & video", "Photography (candid + traditional)", "Confirm 2-day coverage & deliverables", "Both"),
    ("Photography & video", "Cinematography / highlight film", "Ask about raw footage access", "Both"),
    ("Photography & video", "Drone / special coverage", "Needs venue permission", "25 Feb"),
    ("Photography & video", "Photo album / prints", "Order after selecting favourites", "Before"),
    ("Travel & cabs", "Airport / station guest pickups", "Track arrivals in a shared sheet", "Both"),
    ("Travel & cabs", "Guest shuttle (hotel ↔ venue) both days", "Tempo traveller / mini-bus often efficient", "Both"),
    ("Travel & cabs", "Baraat / family procession vehicles", "Decorated car + support vehicles", "25 Feb"),
    ("Travel & cabs", "Couple / VIP car (24–25 Feb)", "Optional", "Both"),
    ("Travel & cabs", "Outstation / intercity travel help", "If reimbursing close family", "Both"),
    ("Travel & cabs", "Parking / valet", "Check if venue includes this", "Both"),
    ("Entertainment & events", "DJ / sound system (25th)", "Confirm power load with venue", "25 Feb"),
    ("Entertainment & events", "Sangeet / mehendi artist (24th)", "Dhol, singer, or host + playlist", "24 Feb"),
    ("Entertainment & events", "Emcee / host", "Helps keep timeline on track", "Both"),
    ("Entertainment & events", "Entertainment extras (fireworks, LED)", "Check local permits", "25 Feb"),
    ("Beauty & grooming", "Bridal makeup & hair (trial + day)", "Schedule trial ahead of the wedding", "25 Feb"),
    ("Beauty & grooming", "Family makeup slots", "Bundle if several people", "25 Feb"),
    ("Beauty & grooming", "Groom grooming / barber", "Day-of touch-up", "25 Feb"),
    ("Invites & favours", "Invitations (print + digital)", "Digital RSVP simplifies headcounts", "Before"),
    ("Invites & favours", "Guest favours / return gifts", "Edible favours are popular", "Both"),
    ("Invites & favours", "Welcome kits for staying guests", "Water, snacks, schedule card", "24 Feb"),
    ("Misc & contingency", "Pooja / ceremony essentials", "Pandit, samagri, seating", "25 Feb"),
    ("Misc & contingency", "Wedding planner / coordinator day-of", "Helpful for 2-day logistics", "Both"),
    ("Misc & contingency", "Tips & staff gratitude", "Venue, waiters, drivers, helpers", "Both"),
    ("Misc & contingency", "Contingency buffer", "Do not spend unless needed", "Both"),
    ("Misc & contingency", "Unexpected / overflow", "Last-minute guests, weather, extras", "Both"),
]

dv_status = DataValidation(
    type="list",
    formula1='"Planned,Getting quotes,Booked,Deposit paid,Fully paid,Cancelled"',
    allow_blank=True,
)
ws3.add_data_validation(dv_status)

dv_day = DataValidation(type="list", formula1='"24 Feb,25 Feb,Both,Before"', allow_blank=True)
ws3.add_data_validation(dv_day)

for i, (cat, item, tip, day) in enumerate(items, start=5):
    text_cell(ws3.cell(row=i, column=1), cat)
    text_cell(ws3.cell(row=i, column=2), item)
    text_cell(ws3.cell(row=i, column=3), tip)
    blank_money(ws3.cell(row=i, column=4), input_fill)
    blank_money(ws3.cell(row=i, column=5), actual_fill)
    blank_money(ws3.cell(row=i, column=6), actual_fill)
    blank_input(ws3.cell(row=i, column=7), "Planned")
    blank_input(ws3.cell(row=i, column=8))
    blank_input(ws3.cell(row=i, column=9))
    blank_input(ws3.cell(row=i, column=10), day)
    blank_input(ws3.cell(row=i, column=11))
    if i % 2 == 0:
        ws3.cell(row=i, column=1).fill = alt_row
        ws3.cell(row=i, column=2).fill = alt_row
        ws3.cell(row=i, column=3).fill = alt_row

dv_status.add("G5:G55")
dv_day.add("J5:J55")

last_item_row = 4 + len(items)
total_row = 60
ws3.cell(row=total_row, column=1, value="TOTAL (this sheet)").font = total_font
ws3.cell(row=total_row, column=1).fill = total_fill
ws3.merge_cells(f"A{total_row}:C{total_row}")
for col in range(1, 4):
    ws3.cell(row=total_row, column=col).fill = total_fill
    ws3.cell(row=total_row, column=col).border = thin
# Columns shifted: Planned=D, Actual=E, Paid=F — but Budget Summary still refs E/F/G
# Keep compatibility: use E=Planned, F=Actual, G=Paid in formulas by inserting empty Suggested col? 
# Simpler: update Budget Summary to D/E/F and totals accordingly.
# I removed Suggested column, so:
# D=Planned, E=Actual, F=Paid
# Update Budget Summary and total formulas.

for col, letter in [(4, "D"), (5, "E"), (6, "F")]:
    cell = ws3.cell(row=total_row, column=col, value=f"=SUM({letter}5:{letter}{last_item_row})")
    money(cell)
    cell.fill = total_fill
    cell.font = total_font
for col in range(7, 12):
    ws3.cell(row=total_row, column=col).fill = total_fill
    ws3.cell(row=total_row, column=col).border = thin

# Place totals also on E60/F60/G60 aliases expected by summary?
# Update summary sheet formulas instead (done below after rewrite awareness).
# Actually Budget Summary currently references E60, F60, G60 for Planned/Actual/Paid.
# With new layout Planned=D, Actual=E, Paid=F. Fix summary + put dummy? Better fix summary.

ws3["A62"] = "Also fill Food & Meals and Guest Stay — kept separate for per-meal / per-room planning."
ws3["A62"].font = note_font

set_widths(ws3, [24, 42, 48, 14, 14, 12, 14, 18, 12, 14, 28])

# Fix Budget Summary to match new column layout (D/E/F = Planned/Actual/Paid)
ws2["B5"] = "='Detailed Expenses'!D60+'Food & Meals'!F30+'Guest Stay'!G25"
ws2["B6"] = "='Detailed Expenses'!E60+'Food & Meals'!G30+'Guest Stay'!H25"
ws2["B7"] = "='Detailed Expenses'!F60+'Food & Meals'!I30+'Guest Stay'!J25"

for i, (label, key) in enumerate(categories, start=13):
    if key:
        ws2.cell(
            row=i,
            column=2,
            value=f"=SUMIF('Detailed Expenses'!$A$5:$A$55,\"{key}\",'Detailed Expenses'!$D$5:$D$55)",
        )
        ws2.cell(
            row=i,
            column=3,
            value=f"=SUMIF('Detailed Expenses'!$A$5:$A$55,\"{key}\",'Detailed Expenses'!$E$5:$E$55)",
        )
        ws2.cell(
            row=i,
            column=4,
            value=f"=SUMIF('Detailed Expenses'!$A$5:$A$55,\"{key}\",'Detailed Expenses'!$F$5:$F$55)",
        )

# ============================================================
# SHEET 4: Food & Meals
# ============================================================
ws4 = wb.create_sheet("Food & Meals")
ws4.sheet_view.showGridLines = False

ws4["A1"] = "Food & Catering — meal-by-meal (24 & 25 Feb)"
ws4["A1"].font = title_font
ws4.merge_cells("A1:J1")

ws4["A2"] = "Enter your own headcounts, rates, and totals. All amount cells start blank."
ws4["A2"].font = note_font
ws4.merge_cells("A2:J2")

ws4["A3"] = "Main-day guest count (from Overview):"
ws4["A3"].font = label_font
ws4["B3"] = "=Overview!B9"
ws4["B3"].fill = PatternFill("solid", fgColor="F5F5F5")
ws4["B3"].border = thin

food_headers = [
    "Day",
    "Meal",
    "Suggestion",
    "Headcount",
    "Rate / plate",
    "Planned total",
    "Actual total",
    "Vendor",
    "Paid",
    "Notes",
]
for i, h in enumerate(food_headers, 1):
    ws4.cell(row=5, column=i, value=h)
style_header(ws4, 5, 1, 10, green_header)
ws4.freeze_panes = "A6"

meals = [
    ("24 Feb", "Welcome drinks / high tea", "Light snacks on guest arrival"),
    ("24 Feb", "Lunch", "Often smaller than dinner if guests arrive later"),
    ("24 Feb", "Evening snacks / mehendi bites", "Finger food during mehendi/sangeet"),
    ("24 Feb", "Dinner", "Often the bigger 24th meal if sangeet follows"),
    ("24 Feb", "Late-night snacks", "Optional — after sangeet"),
    ("25 Feb", "Breakfast", "Mostly for staying guests; hotel may include"),
    ("25 Feb", "Mid-morning tea / snacks", "During prep / pooja"),
    ("25 Feb", "Lunch", "Timing depends on ceremony schedule"),
    ("25 Feb", "Wedding dinner / reception (MAIN)", "Usually the highest catering line"),
    ("25 Feb", "Dessert / ice cream counter", "Sometimes separate from dinner package"),
    ("Both", "Water, soft drinks & beverages", "Budget explicitly — often underestimated"),
    ("Both", "Kids / special diet / Jain meals", "Confirm counts about a week prior"),
    ("Both", "Staff / vendor meals", "Photographers, makeup, band, helpers"),
    ("Before", "Tasting session with caterer", "Lock menu after tasting"),
]

for i, (day, meal, tip) in enumerate(meals, start=6):
    text_cell(ws4.cell(row=i, column=1), day)
    text_cell(ws4.cell(row=i, column=2), meal)
    text_cell(ws4.cell(row=i, column=3), tip)
    blank_input(ws4.cell(row=i, column=4))
    blank_money(ws4.cell(row=i, column=5), input_fill)
    # Planned total = headcount * rate when both filled
    c = ws4.cell(row=i, column=6, value=f'=IF(OR(D{i}="",E{i}=""),"",D{i}*E{i})')
    money(c)
    c.fill = PatternFill("solid", fgColor="F5F5F5")
    blank_money(ws4.cell(row=i, column=7), actual_fill)
    blank_input(ws4.cell(row=i, column=8))
    blank_money(ws4.cell(row=i, column=9), actual_fill)
    blank_input(ws4.cell(row=i, column=10))

food_last = 5 + len(meals)

for i in range(food_last + 1, 29):
    for col in range(1, 11):
        ws4.cell(row=i, column=col).border = thin
        if col in (1, 2, 3, 4, 8, 10):
            ws4.cell(row=i, column=col).fill = input_fill
        if col == 5:
            blank_money(ws4.cell(row=i, column=col), input_fill)
        if col == 6:
            ws4.cell(row=i, column=col, value=f'=IF(OR(D{i}="",E{i}=""),"",D{i}*E{i})')
            money(ws4.cell(row=i, column=col))
        if col in (7, 9):
            blank_money(ws4.cell(row=i, column=col), actual_fill)

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
    "Lock final headcounts about 7–10 days before; many caterers charge a guaranteed minimum.",
    "25th dinner often runs higher than RSVPs (plus-ones & drop-ins).",
    "Breakfast may already be included in hotel stay — avoid double-counting.",
    "Confirm veg / non-veg / Jain / kids plates and whether GST & service staff are included.",
]
for i, t in enumerate(tips, start=33):
    ws4.cell(row=i, column=1, value=t).font = Font(name="Calibri", size=11)
    ws4.merge_cells(f"A{i}:J{i}")

ws4["A38"] = "DAY SUBTOTALS"
ws4["A38"].font = section_font
ws4["A38"].fill = section_fill

ws4["A39"] = "24 Feb food (planned)"
ws4["B39"] = '=SUMIF(A6:A28,"24 Feb",F6:F28)'
money(ws4["B39"])
ws4["A40"] = "25 Feb food (planned)"
ws4["B40"] = '=SUMIF(A6:A28,"25 Feb",F6:F28)'
money(ws4["B40"])
ws4["A41"] = "Both / other (planned)"
ws4["B41"] = '=SUMIF(A6:A28,"Both",F6:F28)+SUMIF(A6:A28,"Before",F6:F28)'
money(ws4["B41"])

set_widths(ws4, [12, 36, 48, 12, 14, 14, 14, 18, 12, 28])

# ============================================================
# SHEET 5: Guest Stay
# ============================================================
ws5 = wb.create_sheet("Guest Stay")
ws5.sheet_view.showGridLines = False

ws5["A1"] = "Guest Stay — rooms for 24 & 25 February"
ws5["A1"].font = title_font
ws5.merge_cells("A1:K1")

ws5["A2"] = "Fill rooms, nights, and rates yourself. Planned total = rooms × rate × nights when those three are filled."
ws5["A2"].font = note_font
ws5.merge_cells("A2:K2")

ws5["A3"] = "Guests needing stay:"
ws5["B3"] = "=Overview!B10"
ws5["B3"].border = thin
ws5["C3"] = "Nights (from Overview):"
ws5["D3"] = "=Overview!B11"
ws5["D3"].border = thin

stay_headers = [
    "Hotel / stay",
    "Room type",
    "# Rooms",
    "Check-in",
    "Check-out",
    "Nights",
    "Rate / room / night",
    "Planned total",
    "Actual total",
    "Paid",
    "Guest names / notes",
]
for i, h in enumerate(stay_headers, 1):
    ws5.cell(row=5, column=i, value=h)
style_header(ws5, 5, 1, 11, teal_header)
ws5.freeze_panes = "A6"

stay_rows = [
    ("Primary hotel (guest block)", "Deluxe double", "24 Feb", "26 Feb"),
    ("Primary hotel (guest block)", "Suite / family", "24 Feb", "26 Feb"),
    ("Secondary hotel / overflow", "Standard double", "24 Feb", "26 Feb"),
    ("Couple suite / wedding night", "Suite", "25 Feb", "26 Feb"),
    ("Early arrivals", "Deluxe double", "23 Feb", "26 Feb"),
    ("Driver / staff rooms", "Economy", "24 Feb", "26 Feb"),
]

for i, (hotel, rtype, cin, cout) in enumerate(stay_rows, start=6):
    blank_input(ws5.cell(row=i, column=1), hotel)
    blank_input(ws5.cell(row=i, column=2), rtype)
    blank_input(ws5.cell(row=i, column=3))
    blank_input(ws5.cell(row=i, column=4), cin)
    blank_input(ws5.cell(row=i, column=5), cout)
    blank_input(ws5.cell(row=i, column=6))
    blank_money(ws5.cell(row=i, column=7), input_fill)
    c = ws5.cell(row=i, column=8, value=f'=IF(OR(C{i}="",F{i}="",G{i}=""),"",C{i}*F{i}*G{i})')
    money(c)
    c.fill = PatternFill("solid", fgColor="F5F5F5")
    blank_money(ws5.cell(row=i, column=9), actual_fill)
    blank_money(ws5.cell(row=i, column=10), actual_fill)
    blank_input(ws5.cell(row=i, column=11))

for i in range(12, 24):
    for col in range(1, 12):
        ws5.cell(row=i, column=col).border = thin
        if col in (1, 2, 3, 4, 5, 6, 11):
            ws5.cell(row=i, column=col).fill = input_fill
        if col == 7:
            blank_money(ws5.cell(row=i, column=col), input_fill)
        if col == 8:
            ws5.cell(row=i, column=col, value=f'=IF(OR(C{i}="",F{i}="",G{i}=""),"",C{i}*F{i}*G{i})')
            money(ws5.cell(row=i, column=col))
        if col in (9, 10):
            blank_money(ws5.cell(row=i, column=col), actual_fill)

# Budget Summary expects G25 Planned, H25 Actual, J25 Paid — remap:
# New layout: H=Planned, I=Actual, J=Paid
# Update summary refs to H/I/J and total row accordingly.

ws5.cell(row=25, column=1, value="TOTAL GUEST STAY").font = total_font
ws5.cell(row=25, column=1).fill = total_fill
ws5.merge_cells("A25:G25")
for col in range(1, 8):
    ws5.cell(row=25, column=col).fill = total_fill
    ws5.cell(row=25, column=col).border = thin
for col, letter in [(8, "H"), (9, "I"), (10, "J")]:
    cell = ws5.cell(row=25, column=col, value=f"=SUM({letter}6:{letter}23)")
    money(cell)
    cell.fill = total_fill
    cell.font = total_font
ws5.cell(row=25, column=11).fill = total_fill
ws5.cell(row=25, column=11).border = thin

# Fix Budget Summary stay refs
ws2["B5"] = "='Detailed Expenses'!D60+'Food & Meals'!F30+'Guest Stay'!H25"
ws2["B6"] = "='Detailed Expenses'!E60+'Food & Meals'!G30+'Guest Stay'!I25"
ws2["B7"] = "='Detailed Expenses'!F60+'Food & Meals'!I30+'Guest Stay'!J25"

for i, (label, key) in enumerate(categories, start=13):
    if label.startswith("Guest stay"):
        ws2.cell(row=i, column=2, value="='Guest Stay'!H25")
        ws2.cell(row=i, column=3, value="='Guest Stay'!I25")
        ws2.cell(row=i, column=4, value="='Guest Stay'!J25")

ws5["A27"] = "Tip: Rough room need ≈ guests needing stay ÷ 2 (doubles), plus a few singles/suites."
ws5["A27"].font = note_font
ws5.merge_cells("A27:K27")

ws5["A28"] = "Ask hotels about wedding attrition (complimentary rooms against banquet spend) and breakfast inclusion."
ws5["A28"].font = note_font
ws5.merge_cells("A28:K28")

set_widths(ws5, [28, 16, 10, 12, 12, 10, 18, 14, 14, 12, 32])

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
    "Quote",
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
    blank_input(ws6.cell(row=i, column=1), cat)
    for col in range(2, 11):
        blank_input(ws6.cell(row=i, column=col))
    blank_money(ws6.cell(row=i, column=6), input_fill)
    blank_input(ws6.cell(row=i, column=9), prio)

for i in range(14, 31):
    for col in range(1, 11):
        blank_input(ws6.cell(row=i, column=col))
    blank_money(ws6.cell(row=i, column=6), input_fill)

set_widths(ws6, [22, 22, 16, 14, 24, 12, 12, 14, 14, 28])

# ============================================================
# SHEET 7: Payments
# ============================================================
ws7 = wb.create_sheet("Payments")
ws7.sheet_view.showGridLines = False

ws7["A1"] = "Payment schedule & tracker"
ws7["A1"].font = title_font
ws7.merge_cells("A1:I1")

ws7["A2"] = "Typical pattern: advance on booking → mid instalment → balance before/on event day. Enter your own figures."
ws7["A2"].font = note_font

pay_headers = [
    "Vendor / item",
    "Total due",
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

dv_paid = DataValidation(type="list", formula1='"Yes,No,Partial"', allow_blank=True)
ws7.add_data_validation(dv_paid)
dv_paid.add("D5:D40")
dv_paid.add("F5:F40")
dv_paid.add("H5:H40")

sample_names = [
    "Venue",
    "Caterer (all meals)",
    "Guest hotel block",
    "Decorator",
    "Photographer + video",
    "Makeup",
    "Travel / cabs",
    "Outfits",
]

for i, name in enumerate(sample_names, start=5):
    blank_input(ws7.cell(row=i, column=1), name)
    blank_money(ws7.cell(row=i, column=2), input_fill)
    blank_money(ws7.cell(row=i, column=3), input_fill)
    blank_input(ws7.cell(row=i, column=4))
    blank_money(ws7.cell(row=i, column=5), input_fill)
    blank_input(ws7.cell(row=i, column=6))
    blank_money(ws7.cell(row=i, column=7), input_fill)
    blank_input(ws7.cell(row=i, column=8))
    blank_input(ws7.cell(row=i, column=9))

for i in range(13, 25):
    for col in range(1, 10):
        blank_input(ws7.cell(row=i, column=col))
    for col in (2, 3, 5, 7):
        blank_money(ws7.cell(row=i, column=col), input_fill)

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

ws7["A28"] = "Suggested calendar (no amounts — just reminders)"
ws7["A28"].font = section_font
ws7["A28"].fill = section_fill
ws7.merge_cells("A28:C28")

cal = [
    ("Now → Dec", "Book venue, caterer, hotel block, photographer — pay advances"),
    ("January", "Lock menus, décor concept, outfits fittings — mid instalments"),
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
# SHEET 8: Guest List Helper
# ============================================================
ws8 = wb.create_sheet("Guest List Helper")
ws8.sheet_view.showGridLines = False

ws8["A1"] = "Guest list helper — stay & meal flags"
ws8["A1"].font = title_font
ws8.merge_cells("A1:H1")

ws8["A2"] = "Optional: list guests and mark stay / meals. Counts at the bottom help planning."
ws8["A2"].font = note_font

gh = [
    "Name",
    "Side (Bride/Groom)",
    "City",
    "Needs stay?",
    "Nights",
    "24 Lunch",
    "24 Dinner",
    "25 Breakfast",
    "25 Lunch",
    "25 Dinner",
    "Notes",
]
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
        blank_input(ws8.cell(row=i, column=col))

ws8["A56"] = "COUNTS"
ws8["A56"].font = section_font
ws8["A56"].fill = section_fill

counts = [
    (57, "Guests needing stay", 'D5:D54'),
    (58, "24 Lunch", 'F5:F54'),
    (59, "24 Dinner", 'G5:G54'),
    (60, "25 Breakfast", 'H5:H54'),
    (61, "25 Lunch", 'I5:I54'),
    (62, "25 Dinner", 'J5:J54'),
]
for row, label, rng in counts:
    ws8.cell(row=row, column=1, value=label).font = label_font
    ws8.cell(row=row, column=2, value=f'=COUNTIF({rng},"Yes")')
    ws8.cell(row=row, column=2).border = thin
    ws8.cell(row=row, column=2).fill = PatternFill("solid", fgColor="F5F5F5")

set_widths(ws8, [22, 14, 14, 12, 10, 12, 12, 14, 12, 12, 24])

out = "/workspace/Wedding_Expenses_24-25_Feb.xlsx"
wb.save(out)
print(f"Saved {out}")
print(f"Sheets: {wb.sheetnames}")
