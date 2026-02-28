# Propfolio — Property Portfolio P&L Builder  v1.0

Upload property PDFs, add manual entries, and generate a fully formatted Excel P&L workbook for each property and your entire portfolio.

---

## Quick Start

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open `http://localhost:8501` in your browser.

---

## Workflow

| Step | What you do |
|---|---|
| ① Setup | Set number of properties, FY start month, FY range, and property details |
| ② Upload PDFs | Drop in rental statements, bank records, utility bills, invoices — auto-parsed |
| ③ Review & Edit | Check editable data tables · Add fixed/recurring expenses via ⚡ Add Entry |
| ④ Generate Excel | Pick a colour theme · Download workbook + Session JSON |

**Monthly update:** Load your saved Session JSON in Setup, upload only new PDFs, download updated JSON when done.

**No JSON?** Use *Restore from Excel* in Setup to rebuild your session from a previously generated workbook.

---

## Supported PDF Types

| Type | What's extracted |
|---|---|
| Rental / Ownership Statement | Rental income, management fees, net EFT amount, itemised bill expenses |
| Bank Statement | Transactions auto-categorised into P&L items (mortgage, repairs, insurance, etc.) |
| Utility Bill | Electricity, water, gas, internet — mapped to the correct utility line |
| Tax Invoice / Notice | Council rates, land tax, strata levies, building insurance, trade invoices |

---

## ⚡ Add Entry (Step 3)

Add any expense not captured in a PDF — fixed, recurring, or one-off:

- **Toggle off** — single manual entry: one category, one month
- **Toggle on → Mode A** — same amount each entry (e.g. Internet $89 × 12 months)
- **Toggle on → Mode B** — total ÷ N entries, split evenly (e.g. Insurance $1,200 ÷ 12)
- **Interval** — every 1 / 3 / 6 months (quarterly Strata, semi-annual reviews, etc.)

---

## Output Excel

- **Property tabs** — Full P&L with monthly columns (FY-grouped, collapsible), FY & CY totals, KPI table (NOI, Net Profit, DSCR)
- **Summary tab** — Portfolio asset table (yield, LVR, equity) + performance summary across all properties and periods
- **3 colour themes** — Navy Professional · Slate & Sage · Charcoal & Amber
- **Semantic row colours** — 🟢 Income · 🔴 Expenses · 🔵 Net/Profit · 🟣 Cash Flow
- **Period colours** — Yellow = active FY · Lt. Yellow = active CY · Grey = inactive · Blue = input cell

---

## Address Validation

For every non-bank PDF, the app extracts the property address and compares it to your Setup address:

| Status | Meaning | Default |
|---|---|---|
| ✅ Matched | Address confirmed | Included |
| ⚠️ Partial | Same suburb, verify number/unit | Included |
| ❌ Mismatch | Different postcode or address | Excluded |
| ⚪ Not found | No address in PDF | Included |

Override any decision with the **Include in P&L** checkbox in Step 2.
