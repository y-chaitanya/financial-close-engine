-- ============================================================
-- 02_allocations_and_tax.sql
-- Allocates shared corporate overhead to business units and
-- accrues California use tax on untaxed out-of-state purchases.
-- Run after 01_data_validation.sql in the same session.
-- ============================================================

DROP TABLE IF EXISTS use_tax_accrual;
DROP TABLE IF EXISTS allocation_entries;
DROP TABLE IF EXISTS taxable_accounts;
DROP TABLE IF EXISTS bu_revenue;

-- ---------- Allocation driver: revenue by business unit ----------

CREATE TABLE bu_revenue (
    business_unit  TEXT PRIMARY KEY,
    period_revenue REAL NOT NULL
);

INSERT INTO bu_revenue VALUES
('UNIT_ALPHA', 6000000.00),
('UNIT_BETA',  4000000.00);

-- ---------- Which accounts represent taxable goods ----------
-- Use tax applies to tangible personal property.
-- Services, rent, and software subscriptions are excluded here.

CREATE TABLE taxable_accounts (
    account_code INTEGER PRIMARY KEY
);

INSERT INTO taxable_accounts VALUES (60100), (60200);

-- ---------- Allocate corporate overhead ----------

CREATE TABLE allocation_entries (
    business_unit     TEXT,
    source_account    INTEGER,
    corporate_amount  REAL,
    revenue_share     REAL,
    allocated_amount  REAL
);

INSERT INTO allocation_entries
SELECT
    r.business_unit,
    c.gl_account_code,
    ROUND(SUM(c.amount), 2),
    ROUND(r.period_revenue / (SELECT SUM(period_revenue) FROM bu_revenue), 4),
    ROUND(SUM(c.amount) * r.period_revenue
          / (SELECT SUM(period_revenue) FROM bu_revenue), 2)
FROM ap_invoices_clean c
CROSS JOIN bu_revenue r
WHERE c.business_unit = 'CORPORATE'
GROUP BY r.business_unit, c.gl_account_code, r.period_revenue;

-- ---------- Accrue use tax ----------

CREATE TABLE use_tax_accrual (
    invoice_id       TEXT,
    vendor_name      TEXT,
    vendor_state     TEXT,
    district_name    TEXT,
    invoice_date     TEXT,
    taxable_amount   REAL,
    tax_rate         REAL,
    use_tax_accrued  REAL,
    credit_account   INTEGER
);

INSERT INTO use_tax_accrual
SELECT
    c.invoice_id,
    c.vendor_name,
    c.vendor_state,
    c.district_name,
    c.invoice_date,
    c.amount,
    d.tax_rate,
    ROUND(c.amount * d.tax_rate, 2),
    22100
FROM ap_invoices_clean c
JOIN taxable_accounts t
  ON t.account_code = c.gl_account_code
JOIN district_tax_rates d
  ON  d.district_name  = c.district_name
  AND c.invoice_date  >= d.effective_date
  AND c.invoice_date  <= COALESCE(d.end_date, '9999-12-31')
WHERE c.sales_tax_charged = 'N'
  AND c.vendor_state <> 'CA';

-- ---------- Results ----------

SELECT * FROM allocation_entries ORDER BY business_unit, source_account;

SELECT
  (SELECT ROUND(SUM(amount), 2) FROM ap_invoices_clean
    WHERE business_unit = 'CORPORATE')        AS corporate_pool,
  (SELECT ROUND(SUM(allocated_amount), 2)
     FROM allocation_entries)                 AS total_allocated;

SELECT * FROM use_tax_accrual ORDER BY invoice_id;

SELECT COUNT(*) AS invoices_accrued,
       ROUND(SUM(use_tax_accrued), 2) AS total_use_tax_liability
FROM use_tax_accrual;
