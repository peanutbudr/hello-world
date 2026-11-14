# Wedding Expense Planner (24–25 Feb)

Excel workbook for a **2-day wedding** (main event **25 Feb**).

**File:** [`Wedding_Expenses_24-25_Feb.xlsx`](./Wedding_Expenses_24-25_Feb.xlsx)

Works in **Excel**, **Google Sheets** (File → Import), and LibreOffice.

## Sheets

| Sheet | What it’s for |
| --- | --- |
| **Overview** | Guest count, stay nights, total budget target, suggested % split |
| **Budget Summary** | Live planned vs actual vs paid + pie chart |
| **Detailed Expenses** | Venue, décor, outfits, photo, cabs, entertainment, beauty, misc |
| **Food & Meals** | 24th lunch/dinner + 25th breakfast/lunch/dinner (and more) |
| **Guest Stay** | Hotel room blocks and night-wise cost |
| **Vendors** | Contact directory |
| **Payments** | Advance / instalment / final tracker |
| **Guest List Helper** | Optional RSVP flags for stay + each meal |

## How to use

1. Edit **yellow** cells on **Overview** (especially guest count and budget target).
2. Adjust **Planned** amounts on Detailed / Food / Stay (pre-filled with mid-range ₹ suggestions for ~120 guests).
3. Fill **Actual** and **Paid** as quotes and transfers come in.
4. Watch **Budget Summary** for remaining budget and still-to-pay.

## Notes on suggestions

- Currency defaults to **INR (₹)** — mid-range Indian wedding assumptions.
- Default target budget: **₹25 lakh** for ~100–150 guests (edit freely).
- Food rates assume ~₹800–1,500/plate depending on meal; luxury cities run higher.
- Regenerate the file anytime: `python3 create_wedding_budget.py`
