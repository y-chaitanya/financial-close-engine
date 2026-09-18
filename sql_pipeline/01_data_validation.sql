-- ============================================================
-- 01_data_validation.sql
-- Loads AP data into staging, applies integrity rules,
-- and promotes only clean rows into the ledger table.
-- Accounting period under review: August 2026
-- ============================================================

PRAGMA foreign_keys = ON;

DROP TABLE IF EXISTS validation_exceptions;
DROP TABLE IF EXISTS ap_invoices_clean;
DROP TABLE IF EXISTS ap_invoices_staging;
DROP TABLE IF EXISTS district_tax_rates;
DROP TABLE IF EXISTS gl_accounts;

-- ---------- Master data ----------

CREATE TABLE gl_accounts (
    account_code  INTEGER PRIMARY KEY,
    account_name  TEXT NOT NULL,
    account_type  TEXT NOT NULL
);

INSERT INTO gl_accounts VALUES
(20000,'Accounts Payable','Liability'),
(22100,'Accrued Use Tax Liability','Liability'),
(60100,'Office Supplies Expense','Expense'),
(60200,'Lab Supplies Expense','Expense'),
(61000,'Rent Expense','Expense'),
(62000,'Software Subscriptions Expense','Expense'),
(63000,'Professional Fees Expense','Expense'),
(64000,'Shipping and Freight Expense','Expense');

CREATE TABLE district_tax_rates (
    district_name   TEXT NOT NULL,
    tax_rate        REAL NOT NULL,
    effective_date  TEXT NOT NULL,
    end_date        TEXT
);

INSERT INTO district_tax_rates VALUES
('Statewide Base',0.0725,'2023-01-01',NULL),
('Santa Clarita',0.0925,'2023-01-01','2026-03-31'),
('Santa Clarita',0.0975,'2026-04-01',NULL),
('Los Angeles City',0.0975,'2023-01-01',NULL),
('Sacramento',0.0875,'2023-01-01',NULL);

-- ---------- Staging: raw data, no constraints ----------

CREATE TABLE ap_invoices_staging (
    invoice_id         TEXT,
    vendor_name        TEXT,
    vendor_state       TEXT,
    invoice_date       TEXT,
    amount             REAL,
    gl_account_code    INTEGER,
    business_unit      TEXT,
    district_name      TEXT,
    sales_tax_charged  TEXT
);

INSERT INTO ap_invoices_staging VALUES
('INV-1001','Valley Office Supply','CA','2026-08-03',1250.00,60100,'UNIT_ALPHA','Santa Clarita','Y'),
('INV-1002','Northwind Lab Equipment','NV','2026-08-04',40000.00,60200,'UNIT_ALPHA','Santa Clarita','N'),
('INV-1003','Harbor Property Group','CA','2026-08-01',18500.00,61000,'CORPORATE','Santa Clarita','N'),
('INV-1004','Cascade Software Inc','WA','2026-08-05',7200.00,62000,'CORPORATE','Santa Clarita','N'),
('INV-1005','Bridgeway Consulting','CA','2026-08-07',9500.00,63000,'UNIT_BETA','Los Angeles City','N'),
('INV-1006','Summit Freight Lines','CA','2026-08-08',2340.50,64000,'UNIT_BETA','Los Angeles City','Y'),
('INV-1007','Northwind Lab Equipment','NV','2026-08-11',15750.00,60200,'UNIT_BETA','Santa Clarita','N'),
('INV-1008','Valley Office Supply','CA','2026-08-12',860.25,60100,'UNIT_BETA','Santa Clarita','Y'),
('INV-1009','Pinnacle Analytics','TX','2026-08-14',4800.00,62000,'UNIT_ALPHA','Santa Clarita','N'),
('INV-1010','Summit Freight Lines','CA','2026-08-15',1120.00,64000,'UNIT_ALPHA','Santa Clarita','Y'),
('INV-1011','Bridgeway Consulting','CA','2026-08-18',6200.00,63000,'CORPORATE','Los Angeles City','N'),
('INV-1012','Redwood Facilities','CA','2026-08-19',3400.00,61000,'CORPORATE','Santa Clarita','Y'),
('INV-1013','Pinnacle Analytics','TX','2026-08-21',4800.00,62000,'UNIT_ALPHA','Santa Clarita','N'),
('INV-1014','Lakeshore Instruments','OR','2026-08-22',22600.00,60200,'UNIT_ALPHA','Santa Clarita','N'),
('INV-1015','Valley Office Supply','CA','2026-08-25',-450.00,60100,'UNIT_BETA','Santa Clarita','Y'),
('INV-1016','Cascade Software Inc','WA','2026-05-30',7200.00,62000,'CORPORATE','Santa Clarita','N'),
('INV-1017','Summit Freight Lines','CA','2026-08-27',1875.00,99999,'UNIT_BETA','Los Angeles City','Y'),
('INV-1018','Lakeshore Instruments','OR','2026-08-28',9900.00,60200,'UNIT_BETA','Sacramento','N'),
('INV-1019','Harbor Property Group','CA','2026-08-29',18500.00,61000,'CORPORATE','Santa Clarita','N'),
('INV-1020','Meridian Print Services','CA','2026-08-31',640.00,60100,'UNIT_ALPHA','Santa Clarita','Y');

-- ---------- Ledger table: constraints enforced ----------

CREATE TABLE ap_invoices_clean (
    invoice_id         TEXT PRIMARY KEY,
    vendor_name        TEXT NOT NULL,
    vendor_state       TEXT NOT NULL,
    invoice_date       TEXT NOT NULL,
    amount             REAL NOT NULL CHECK (amount > 0),
    gl_account_code    INTEGER NOT NULL REFERENCES gl_accounts(account_code),
    business_unit      TEXT NOT NULL,
    district_name      TEXT NOT NULL,
    sales_tax_charged  TEXT NOT NULL CHECK (sales_tax_charged IN ('Y','N'))
);

-- ---------- Exception log ----------

CREATE TABLE validation_exceptions (
    invoice_id  TEXT,
    rule_code   TEXT,
    detail      TEXT
);

-- Rule 1: amount must be positive
INSERT INTO validation_exceptions
SELECT invoice_id, 'AMT_NOT_POSITIVE',
       'Amount is ' || amount || '; credits must be processed as vendor credits'
FROM ap_invoices_staging
WHERE amount <= 0;

-- Rule 2: invoice date must fall in the open period
INSERT INTO validation_exceptions
SELECT invoice_id, 'DATE_OUT_OF_PERIOD',
       'Invoice dated ' || invoice_date || '; period under review is 2026-08-01 to 2026-08-31'
FROM ap_invoices_staging
WHERE invoice_date < '2026-08-01' OR invoice_date > '2026-08-31';

-- Rule 3: GL account must exist in the chart of accounts
INSERT INTO validation_exceptions
SELECT s.invoice_id, 'GL_ACCOUNT_NOT_FOUND',
       'Account ' || s.gl_account_code || ' is not in the chart of accounts'
FROM ap_invoices_staging s
LEFT JOIN gl_accounts g ON g.account_code = s.gl_account_code
WHERE g.account_code IS NULL;

-- Rule 4: possible duplicate (same vendor, amount, and account)
INSERT INTO validation_exceptions
SELECT b.invoice_id, 'POSSIBLE_DUPLICATE',
       'Matches ' || a.invoice_id || ' on vendor, amount, and account'
FROM ap_invoices_staging a
JOIN ap_invoices_staging b
  ON  a.vendor_name     = b.vendor_name
  AND a.amount          = b.amount
  AND a.gl_account_code = b.gl_account_code
  AND a.invoice_id      < b.invoice_id;

-- ---------- Promote clean rows only ----------

INSERT INTO ap_invoices_clean
SELECT * FROM ap_invoices_staging
WHERE invoice_id NOT IN (SELECT invoice_id FROM validation_exceptions);

-- ---------- Results ----------

SELECT rule_code, invoice_id, detail
FROM validation_exceptions
ORDER BY rule_code, invoice_id;

SELECT
  (SELECT COUNT(*) FROM ap_invoices_staging)                     AS rows_received,
  (SELECT COUNT(DISTINCT invoice_id) FROM validation_exceptions) AS invoices_flagged,
  (SELECT COUNT(*) FROM ap_invoices_clean)                       AS rows_posted;
