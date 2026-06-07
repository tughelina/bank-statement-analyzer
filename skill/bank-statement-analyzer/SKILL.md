---
name: bank-statement-analyzer
description: Turn a year of bank statements (CSV, Excel, or PDF) into one clean Excel workbook with 6 sheets — all transactions, creditors (detail + summary), debtors/revenue, owner account (salary draws + deposits), and internal transfers. Use when the user wants to analyze, summarize, or get clarity on their bank statements for a year or period, prepare data for an accountant, or build a financial overview from raw statements.
---

# Bank Statement Analyzer

Build ONE professional Excel workbook from a folder of bank statements. Be accurate and reconcile every transaction — never invent numbers.

## Workflow

### 1. Locate & understand the data
- Find the statement files (ask the user for the folder if unknown). Accept CSV, Excel (`.xlsx`), or PDF. CSV/Excel are cleanest; for PDFs, extract the transaction table carefully.
- Detect the columns: **date, description, amount (in/out), running balance, currency, account**. Banks differ — map them explicitly.
- Determine the amount sign convention (usually negative = money out). Confirm against the running balance.
- Report back: period covered, number of transactions, accounts, currencies. Flag anything missing or inconsistent.

### 2. Clarify categories (ask, don't guess — max ~4 questions)
- Which **incoming** transfers are real **revenue** vs. **owner deposits** (capital the user put in themselves)?
- Which **outgoing** transfers are **salary / owner draws** to a private account?
- Which transfers are **internal moves** between the user's own accounts (currency exchange, savings pockets)?
- Output **language and reporting currency**.

### 3. Build the workbook (use Python + `openpyxl`)
Six sheets:
1. **All Transactions** — chronological. Columns: Date · Account · Currency · Type · Description · Income · Expense · Balance · Note. Income tinted green, expense red. Per-currency column totals at the bottom (live `SUMIF`). Do NOT merge currencies/accounts into one running balance — keep the bank's real per-account balance per row.
2. **Creditors — Detail** — every supplier payment, sorted by creditor (A→Z) then date. Merge obvious vendor variants into one name. Refunds as negative (so totals net). All bank fees bundled as one creditor "Bank Fees".
3. **Creditors — Summary** — one row per creditor: count, total, first/last charge date. Alphabetical, grand total.
4. **Debtors** — real revenue only (exclude owner deposits + internal moves). Summary by payer + detail. Per-currency totals if multi-currency.
5. **Owner Account** — Block A salary/draws (business → owner), Block B owner deposits (owner → business), plus net (draws − deposits).
6. **Internal Transfers** — moves between the user's own accounts (currency conversions, savings pockets). No real money in/out.

### 4. Reconcile & report
- Every transaction lands in exactly one category. Print the counts per category and confirm the total matches the input.
- Save the `.xlsx` next to the source files. Give a short summary: revenue, creditors, salary, net owner withdrawal, reconciliation counts.

## Formatting standards
Bold white-on-dark headers, frozen header row, auto-filters, thousands separators with 2 decimals, sensible column widths, coloured sheet tabs.

## Starter script
`template_analysis.py` (in this skill folder) is a configurable skeleton: set the column mapping and category keyword lists at the top, then run `python3 template_analysis.py "<folder>"`. Adapt it to the user's specific bank format and categories — it is a starting point, not a fixed solution.

---
*© 2026 SoulStrategy by Chantal Perrinjaquet — House Of Coaching. Shared as a gift; please keep this credit.*
