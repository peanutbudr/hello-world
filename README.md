# Wedding Expense Planner (24–25 Feb)

Blank Excel workbook for a **2-day wedding** (main event **25 Feb**). Category rows and tips are included; **all money fields start empty** for you to fill in.

**File:** [`Wedding_Expenses_24-25_Feb.xlsx`](./Wedding_Expenses_24-25_Feb.xlsx)

Works in **Excel**, **Google Sheets** (File → Import), and LibreOffice.

## Sheets

| Sheet | What it’s for |
| --- | --- |
| **Overview** | Guest count, stay nights, budget target (blank), category checklist |
| **Budget Summary** | Live planned vs actual vs paid (fills as you enter numbers) |
| **Detailed Expenses** | Venue, décor, outfits, photo, cabs, entertainment, beauty, misc |
| **Food & Meals** | 24th lunch/dinner + 25th breakfast/lunch/dinner (and more) |
| **Guest Stay** | Hotel room blocks |
| **Vendors** | Contact directory |
| **Payments** | Advance / instalment / final tracker |
| **Guest List Helper** | Optional RSVP flags for stay + each meal |

## How to use

1. Fill yellow cells on **Overview**.
2. Enter **Planned / Actual / Paid** on Detailed, Food, and Stay sheets.
3. Watch **Budget Summary** update.

Regenerate: `python3 create_wedding_budget.py`
