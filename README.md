# Financial Data and Close Controls Engine

A hands-on project that models how transaction data moves from raw procurement files into a general ledger, gets allocated across business units, accrues use tax, and passes a month-end control check before the books close.

Built with SQL and Excel VBA, using synthetic data. The goal is to show the accounting logic behind ERP systems such as Oracle NetSuite and SAP, not to replace them.

---

## Why I built it

Corporate accounting runs on data that arrives messy: invoices from an AP platform, a chart of accounts from the ERP, tax rates that vary by district. The work is turning that into something reconciled and supportable.

This project walks through that path end to end, so the reasoning is visible at each step: what gets rejected at the door, how shared costs get split, when tax is owed even though the vendor didn't charge it, and what stops a close when the numbers don't tie.

---

## Repository structure

```text
financial-close-engine/
│
├── data_schema/
│   ├── ap_invoices.csv              # Synthetic AP invoice extract
│   ├── gl_accounts.csv              # Chart of accounts master file
│   └── ca_district_tax_rates.csv    # District tax rates with effective dates
│
├── sql_pipeline/
│   ├── 01_data_validation.sql       # Table creation, constraints, integrity checks
│   └── 02_allocations_and_tax.sql   # Cost allocation and use tax accrual
│
├── vba_automation/
│   └── CloseTieOutCheck.bas         # Month-end tie-out control with validation log
│
└── docs/
    └── sample_output.png            # Screenshot of results
```

---

## What each module does

### 1. Validation layer (`01_data_validation.sql`)

Creates the tables with primary keys, foreign keys, and NOT NULL constraints, so invalid rows never reach the ledger. Then runs exception queries that flag:

- Negative or zero invoice amounts
- Invoice dates outside the open accounting period
- Vendor or GL account codes with no match in the master files
- Possible duplicate invoices (same vendor, amount, and date)

Anything flagged goes to an exceptions table for review rather than being silently dropped.

### 2. Allocation and tax layer (`02_allocations_and_tax.sql`)

**Cost allocation.** Shared overhead sits in a corporate cost center and can't stay there if you want profitability by business unit. This script distributes it across business units using a weighted driver, revenue share by default, and writes the allocation entries. It's the same idea behind SAP's profitability analysis module, modeled in plain SQL.

**Use tax accrual.** When an out-of-state vendor doesn't charge California sales tax on a taxable purchase, the buyer still owes use tax directly to the state. The script identifies those invoices, looks up the correct district rate by location and effective date, and posts the accrual to the use tax liability account.

Rates live in a lookup table with effective dates rather than hardcoded in the query, because California district rates differ by address and change over time.

### 3. Month-end tie-out control (`CloseTieOutCheck.bas`)

An Excel VBA macro that runs before the books close. It compares the total of a supporting schedule against the general ledger control total and requires them to match exactly.

If they don't, it writes the variance to a validation log with a timestamp, the Windows user, and the amounts compared, and returns a failure status instead of a clean result. It's a simple control, and that's the point: a difference should stop the process rather than pass unnoticed.

---

## How to run it

**SQL portion**

1. Open any SQLite environment, such as SQLite Online or DB Browser for SQLite.
2. Run `01_data_validation.sql` to create the tables and load the sample data.
3. Run the exception queries at the bottom of that file to see flagged rows.
4. Run `02_allocations_and_tax.sql` to generate allocation entries and use tax accruals.

**Excel portion**

1. Open Excel and press ALT + F11 for the VBA editor.
2. Import `vba_automation/CloseTieOutCheck.bas`.
3. Enter a schedule total and a control total on the input sheet, then run the macro to see a pass result and a failing result with its log entry.

---

## Scope and limitations

- All data is synthetic. No company data is used.
- This models accounting logic; it isn't a connection to NetSuite, SAP, or any live ERP.
- Tax rates in the sample file are illustrative. Real filings require current rates from the California Department of Tax and Fee Administration.
- The tie-out macro is a single control, not a full close checklist or an audit system.

---

## Background

I'm an IRS Enrolled Agent and CPA exam candidate with a computer science degree. The accounting reasoning here (allocation drivers, use tax nexus, reconciliation controls) comes from my tax credential and accounting coursework, and the implementation from my software engineering background.

**Chaitanya Yarlagadda** | IRS Enrolled Agent | CPA Exam Candidate
