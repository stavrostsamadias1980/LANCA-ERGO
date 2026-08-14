-- =============================================================================
-- ERGO INSURANCE MANAGEMENT SYSTEM - POSTGRESQL DATABASE SCHEMA (ergo_insurance_db)
-- LANCA Insurance Services
-- =============================================================================

-- CORE PRODUCTION TABLES
CREATE TABLE IF NOT EXISTS ergo_statements_1411 (
    id SERIAL PRIMARY KEY,
    month_statement VARCHAR(50),
    receipt_number VARCHAR(50),
    policy_number VARCHAR(50),
    start_date VARCHAR(20),
    end_date VARCHAR(20),
    client_lastname VARCHAR(100),
    client_firstname VARCHAR(100),
    product_code VARCHAR(20),
    tier VARCHAR(50),
    net_total NUMERIC(12,2),
    commission_total NUMERIC(12,2),
    tax_amount NUMERIC(12,2),
    payment_freq NUMERIC(5,2),
    duration_years NUMERIC(5,2),
    policy_year NUMERIC(5,2)
);

CREATE TABLE IF NOT EXISTS ergo_company_payouts (
    id SERIAL PRIMARY KEY,
    payout_date VARCHAR(20),
    deposit_month VARCHAR(50),
    month_statement VARCHAR(50),
    payment_code VARCHAR(150),
    credit_amount NUMERIC(12,2),
    debit_amount NUMERIC(12,2),
    raw_text TEXT
);

CREATE TABLE IF NOT EXISTS ergo_audit_logs (
    id SERIAL PRIMARY KEY,
    timestamp VARCHAR(50),
    username VARCHAR(100),
    action VARCHAR(100),
    details TEXT,
    ip_address VARCHAR(50)
);

CREATE INDEX IF NOT EXISTS idx_ergo_statements_policy ON ergo_statements_1411(policy_number);
CREATE INDEX IF NOT EXISTS idx_ergo_statements_month ON ergo_statements_1411(month_statement);
CREATE INDEX IF NOT EXISTS idx_ergo_payouts_month ON ergo_company_payouts(month_statement);

-- 1. DROP EXISTING TABLES AND VIEWS IF NEEDED
DROP VIEW IF EXISTS v_policy_full_details CASCADE;
DROP VIEW IF EXISTS v_monthly_commissions_summary CASCADE;

DROP TABLE IF EXISTS commissions CASCADE;
DROP TABLE IF EXISTS coverages CASCADE;
DROP TABLE IF EXISTS contracts CASCADE;
DROP TABLE IF EXISTS sales_regulation_rates CASCADE;
DROP TABLE IF EXISTS products CASCADE;

-- 2. PRODUCTS MASTER TABLE
CREATE TABLE products (
    product_code VARCHAR(20) PRIMARY KEY,
    product_name VARCHAR(150) NOT NULL,
    branch VARCHAR(50) NOT NULL,
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_products_branch ON products(branch);

-- 3. CONTRACTS TABLE
CREATE TABLE contracts (
    policy_number VARCHAR(50) PRIMARY KEY,
    client_id VARCHAR(50),
    client_name VARCHAR(150) NOT NULL,
    product_code VARCHAR(20) REFERENCES products(product_code) ON DELETE SET NULL,
    start_date DATE,
    policy_duration INTEGER DEFAULT 1,
    payment_frequency VARCHAR(20) DEFAULT 'Ετήσιο',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_contracts_client ON contracts(client_name);
CREATE INDEX idx_contracts_start_date ON contracts(start_date);

-- 4. COMMISSIONS MASTER TABLE (WITH DEDUPLICATION CONSTRAINT)
CREATE TABLE commissions (
    id BIGSERIAL PRIMARY KEY,
    policy_number VARCHAR(50) REFERENCES contracts(policy_number) ON DELETE CASCADE,
    receipt_number VARCHAR(50),
    statement_date DATE,
    statement_month VARCHAR(15) NOT NULL,
    tier_role VARCHAR(30) NOT NULL, -- 'ΣΥΝΕΡΓΑΤΗΣ' or 'AGENCY'
    policy_year INTEGER DEFAULT 1,
    clean_bk NUMERIC(12, 2) DEFAULT 0.00,
    clean_sk NUMERIC(12, 2) DEFAULT 0.00,
    clean_total NUMERIC(12, 2) DEFAULT 0.00,
    comm_bk NUMERIC(12, 2) DEFAULT 0.00,
    comm_sk NUMERIC(12, 2) DEFAULT 0.00,
    comm_total NUMERIC(12, 2) DEFAULT 0.00,
    comm_rate_pct NUMERIC(6, 4) DEFAULT 0.0000,
    tax NUMERIC(12, 2) DEFAULT 0.00,
    net_payout NUMERIC(12, 2) DEFAULT 0.00,
    imported_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT idx_unique_commission_record UNIQUE(policy_number, receipt_number, statement_month, tier_role)
);

CREATE INDEX idx_commissions_month ON commissions(statement_month);
CREATE INDEX idx_commissions_date ON commissions(statement_date);
CREATE INDEX idx_commissions_tier ON commissions(tier_role);
CREATE INDEX idx_commissions_policy ON commissions(policy_number);

-- 5. COVERAGES TECHNICAL TABLE
CREATE TABLE coverages (
    product_code VARCHAR(20) PRIMARY KEY REFERENCES products(product_code) ON DELETE CASCADE,
    max_coverage_limit VARCHAR(100),
    hospital_room_tier VARCHAR(50),
    hospital_network VARCHAR(150),
    deductible_options VARCHAR(100),
    emergency_limit VARCHAR(100),
    diagnostic_exams VARCHAR(150),
    hospital_allowance VARCHAR(150),
    assistance_services VARCHAR(150),
    riders VARCHAR(255),
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 6. SALES REGULATION RATES TABLE (ΔΑΣ RULES)
CREATE TABLE sales_regulation_rates (
    id BIGSERIAL PRIMARY KEY,
    product_code VARCHAR(20) REFERENCES products(product_code) ON DELETE CASCADE,
    policy_year_label VARCHAR(30) NOT NULL, -- '1ο Έτος', '2ο Έτος', etc.
    rate_cat_a NUMERIC(6, 4) DEFAULT 0.0000,
    rate_cat_b NUMERIC(6, 4) DEFAULT 0.0000,
    rate_cat_c NUMERIC(6, 4) DEFAULT 0.0000,
    unit_override_pct NUMERIC(6, 4) DEFAULT 0.1000,
    agency_override_pct NUMERIC(6, 4) DEFAULT 0.1000,
    notes TEXT
);

-- =============================================================================
-- SEED DATA INSERTIONS
-- =============================================================================

INSERT INTO products (product_code, product_name, branch) VALUES
('020118', 'ERGO Health Care (Simple, Advanced, Superior)', 'Υγεία'),
('020119', 'ERGO Best Health', 'Υγεία'),
('110118', 'ERGO Life - Ισόβια Ασφάλιση Θανάτου', 'Ζωή'),
('110318', 'ERGO Life - Πρόσκαιρη Ασφάλιση Θανάτου', 'Ζωή'),
('990119', 'ERGO My Saving Simple & Junior', 'Αποταμίευση'),
('030122', 'ERGO My Fund Flex Plan', 'Unit-Linked'),
('030222', 'ERGO My Fund Invest Plan', 'Unit-Linked'),
('GROUP01', 'ERGO My People - Βασική Ασφάλιση Θανάτου', 'Ομαδικά'),
('GROUP02', 'ERGO My People - Συμπληρωματικές Καλύψεις', 'Ομαδικά')
ON CONFLICT (product_code) DO NOTHING;

INSERT INTO coverages (product_code, max_coverage_limit, hospital_room_tier, hospital_network, deductible_options, emergency_limit, diagnostic_exams, hospital_allowance, assistance_services, riders) VALUES
('020118', '€500.000 / ασφαλιστικό έτος', 'Α'' Θέση (Μονόκλινο)', '100% Δίκτυο 4U & Συμβεβλημένα', '€0 / €500 / €1.500 / €3.000', '€1.000 Επείγοντα 24/7', 'Περιλαμβάνεται (€2.000)', '€150/ημέρα Νοσοκομειακό + Χειρουργικό', 'ERGOLIFE Assistance 24/7', 'Θάνατος/ΜΑ Ατυχήματος, ΑΠΑ'),
('020119', '€1.000.000 / ασφαλιστικό έτος', 'Α'' Θέση (Μονόκλινο)', 'Πλήρες Δίκτυο ERGO Παγκόσμια', '€1.500 / €3.000', '€1.000 Επείγοντα 24/7', 'Περιλαμβάνεται πλήρως', '€150/ημέρα', 'ERGOLIFE Assistance 24/7', 'Πλήρης κάλυψη εξωτερικού'),
('110118', 'Κεφάλαιο Ζωής €50.000+', 'N/A', 'Ελεύθερη Επιλογή Ιατρών', 'N/A', 'N/A', 'Ιατροφαρμακευτικά Ατυχήματος', 'Νοσοκομειακό Ατυχήματος', 'ERGOLIFE Assistance 24/7', 'Θάνατος / ΜΑ Ατυχήματος, ΑΠΑ'),
('110318', 'Κεφάλαιο Ζωής €30.000+', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'ERGOLIFE Assistance 24/7', 'Βασική Ασφάλιση Θανάτου'),
('990119', 'Εγγυημένο Κεφάλαιο στη Λήξη', 'N/A', 'N/A', 'Τακτικές Καταβολές', 'N/A', 'N/A', 'N/A', 'Ετήσια Ενημέρωση', 'Προστασία Αποταμίευσης'),
('030122', 'Επενδυτικό Χαρτοφυλάκιο Flex', 'N/A', 'N/A', 'Τακτικό Ασφάλιστρο', 'N/A', 'N/A', 'N/A', 'Online Παρακολούθηση', 'Επιλογή Αμοιβαίων'),
('GROUP01', 'Κεφάλαιο Ασφάλισης Προσωπικού', 'Β'' Θέση', 'Δίκτυο 4U Group', '€500', '€500 Επείγοντα', 'Πρωτοβάθμια Περίθαλψη', 'Επίδομα Νοσηλείας', 'ERGOLIFE Group 24/7', 'Bonus +5% Νέων Ομαδικών')
ON CONFLICT (product_code) DO NOTHING;

-- =============================================================================
-- POSTGRESQL ANALYTICS VIEWS
-- =============================================================================

-- VIEW 1: MONTHLY COMMISSIONS & OVERRIDES SUMMARY
CREATE OR REPLACE VIEW v_monthly_commissions_summary AS
SELECT 
    statement_month AS month_code,
    COUNT(DISTINCT policy_number) AS active_policies_count,
    COUNT(*) AS total_transactions,
    SUM(CASE WHEN tier_role = 'ΣΥΝΕΡΓΑΤΗΣ' THEN clean_total ELSE 0 END) AS net_premiums,
    SUM(CASE WHEN tier_role = 'ΣΥΝΕΡΓΑΤΗΣ' THEN comm_total ELSE 0 END) AS comm_agent_total,
    ROUND(
        CASE WHEN SUM(CASE WHEN tier_role = 'ΣΥΝΕΡΓΑΤΗΣ' THEN clean_total ELSE 0 END) > 0 
             THEN (SUM(CASE WHEN tier_role = 'ΣΥΝΕΡΓΑΤΗΣ' THEN comm_total ELSE 0 END) / SUM(CASE WHEN tier_role = 'ΣΥΝΕΡΓΑΤΗΣ' THEN clean_total ELSE 0 END)) * 100 
             ELSE 0 END, 2
    ) AS comm_agent_avg_pct,
    SUM(CASE WHEN tier_role = 'AGENCY' THEN comm_total ELSE 0 END) AS comm_agency_override_total,
    ROUND(
        CASE WHEN SUM(CASE WHEN tier_role = 'ΣΥΝΕΡΓΑΤΗΣ' THEN clean_total ELSE 0 END) > 0 
             THEN (SUM(CASE WHEN tier_role = 'AGENCY' THEN comm_total ELSE 0 END) / SUM(CASE WHEN tier_role = 'ΣΥΝΕΡΓΑΤΗΣ' THEN clean_total ELSE 0 END)) * 100 
             ELSE 0 END, 2
    ) AS comm_agency_avg_pct,
    SUM(comm_total) AS combined_network_payout,
    ROUND(
        CASE WHEN SUM(CASE WHEN tier_role = 'ΣΥΝΕΡΓΑΤΗΣ' THEN clean_total ELSE 0 END) > 0 
             THEN (SUM(comm_total) / SUM(CASE WHEN tier_role = 'ΣΥΝΕΡΓΑΤΗΣ' THEN clean_total ELSE 0 END)) * 100 
             ELSE 0 END, 2
    ) AS combined_network_avg_pct
FROM commissions
GROUP BY statement_month
ORDER BY statement_month ASC;

-- VIEW 2: FULL POLICY DETAILS WITH COVERAGES & FINANCIALS
CREATE OR REPLACE VIEW v_policy_full_details AS
SELECT 
    c.policy_number,
    c.client_name,
    p.product_name,
    p.branch,
    cm.statement_date,
    cm.statement_month,
    cm.policy_year,
    cm.clean_total AS net_premium,
    SUM(CASE WHEN cm.tier_role = 'ΣΥΝΕΡΓΑΤΗΣ' THEN cm.comm_total ELSE 0 END) AS comm_agent,
    ROUND(
        CASE WHEN cm.clean_total > 0 THEN (SUM(CASE WHEN cm.tier_role = 'ΣΥΝΕΡΓΑΤΗΣ' THEN cm.comm_total ELSE 0 END) / cm.clean_total) * 100 ELSE 0 END, 2
    ) AS comm_agent_pct,
    SUM(CASE WHEN cm.tier_role = 'AGENCY' THEN cm.comm_total ELSE 0 END) AS comm_agency_override,
    ROUND(
        CASE WHEN cm.clean_total > 0 THEN (SUM(CASE WHEN cm.tier_role = 'AGENCY' THEN cm.comm_total ELSE 0 END) / cm.clean_total) * 100 ELSE 0 END, 2
    ) AS comm_agency_pct,
    cov.max_coverage_limit,
    cov.hospital_room_tier,
    cov.hospital_network,
    cov.deductible_options,
    cov.emergency_limit,
    cov.diagnostic_exams,
    cov.hospital_allowance,
    cov.assistance_services,
    cov.riders
FROM contracts c
JOIN products p ON c.product_code = p.product_code
JOIN commissions cm ON c.policy_number = cm.policy_number
LEFT JOIN coverages cov ON p.product_code = cov.product_code
GROUP BY 
    c.policy_number, c.client_name, p.product_name, p.branch, 
    cm.statement_date, cm.statement_month, cm.policy_year, cm.clean_total,
    cov.max_coverage_limit, cov.hospital_room_tier, cov.hospital_network, 
    cov.deductible_options, cov.emergency_limit, cov.diagnostic_exams, 
    cov.hospital_allowance, cov.assistance_services, cov.riders
ORDER BY cm.statement_date ASC, c.policy_number ASC;

-- SUCCESS MESSAGE
SELECT 'PostgreSQL ergo_insurance_db schema and views initialized successfully!' AS status;


-- =============================================================================
-- ENHANCED 10-MODULE SCHEMA (3NF RELATIONAL TABLES)
-- =============================================================================

CREATE TABLE IF NOT EXISTS clients (
    client_id VARCHAR(50) PRIMARY KEY,
    ergo_client_code VARCHAR(30),
    full_name VARCHAR(150) NOT NULL,
    afm VARCHAR(20),
    phone_mobile VARCHAR(30),
    phone_landline VARCHAR(30),
    email VARCHAR(100),
    address_street VARCHAR(150),
    city VARCHAR(100),
    postal_code VARCHAR(20),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS insured_persons (
    insured_id VARCHAR(50) PRIMARY KEY,
    client_id VARCHAR(50) NOT NULL,
    full_name VARCHAR(150) NOT NULL,
    birth_date DATE,
    gender VARCHAR(10),
    relationship_type VARCHAR(30) DEFAULT 'PRIMARY',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (client_id) REFERENCES clients(client_id)
);

CREATE TABLE IF NOT EXISTS insurance_products (
    product_id VARCHAR(50) PRIMARY KEY,
    product_code VARCHAR(30) NOT NULL,
    product_name VARCHAR(150) NOT NULL,
    branch_category VARCHAR(50) NOT NULL,
    hospital_class VARCHAR(20),
    max_coverage_limit NUMERIC(12,2),
    default_comm_rate_first_year NUMERIC(5,4),
    default_comm_rate_renewal NUMERIC(5,4),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS policies (
    policy_number VARCHAR(30) PRIMARY KEY,
    client_id VARCHAR(50) NOT NULL,
    primary_insured_id VARCHAR(50),
    producer_partner_code VARCHAR(20) NOT NULL,
    agency_partner_code VARCHAR(20) NOT NULL,
    product_id VARCHAR(50) NOT NULL,
    issue_date DATE,
    start_date DATE,
    expiry_date DATE,
    payment_frequency VARCHAR(30) DEFAULT 'Ετήσια',
    duration_years INT DEFAULT 1,
    current_policy_year INT DEFAULT 1,
    status VARCHAR(30) DEFAULT 'ACTIVE',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (client_id) REFERENCES clients(client_id),
    FOREIGN KEY (product_id) REFERENCES insurance_products(product_id)
);

CREATE TABLE IF NOT EXISTS policy_coverages (
    coverage_id VARCHAR(50) PRIMARY KEY,
    policy_number VARCHAR(30) NOT NULL,
    coverage_code VARCHAR(30) NOT NULL,
    coverage_description VARCHAR(200) NOT NULL,
    insured_capital NUMERIC(12,2) DEFAULT 0.00,
    deductible_amount NUMERIC(10,2) DEFAULT 0.00,
    hospital_class INT DEFAULT 1,
    net_premium NUMERIC(10,2) NOT NULL,
    annual_net_premium NUMERIC(10,2),
    producer_commission_rate NUMERIC(5,4),
    producer_commission_amount NUMERIC(10,2) NOT NULL,
    agency_overriding_amount NUMERIC(10,2) NOT NULL,
    statement_month VARCHAR(10),
    receipt_number VARCHAR(30),
    FOREIGN KEY (policy_number) REFERENCES policies(policy_number)
);

CREATE TABLE IF NOT EXISTS financial_movements (
    movement_id VARCHAR(50) PRIMARY KEY,
    policy_number VARCHAR(30) NOT NULL,
    receipt_number VARCHAR(30) NOT NULL,
    statement_month VARCHAR(10) NOT NULL,
    statement_file_ref VARCHAR(100),
    movement_date VARCHAR(20) NOT NULL,
    iso_date VARCHAR(20),
    movement_type VARCHAR(50) NOT NULL,
    client_name VARCHAR(150),
    package_name VARCHAR(150),
    gross_premium NUMERIC(10,2) NOT NULL,
    net_premium_basic NUMERIC(10,2) DEFAULT 0.00,
    net_premium_supp NUMERIC(10,2) DEFAULT 0.00,
    net_premium_total NUMERIC(10,2) NOT NULL,
    policy_fee NUMERIC(8,2) DEFAULT 0.00,
    tax_amount NUMERIC(8,2) DEFAULT 0.00,
    producer_partner_code VARCHAR(20) NOT NULL,
    producer_commission_amount NUMERIC(10,2) NOT NULL,
    producer_commission_rate NUMERIC(5,4),
    agency_partner_code VARCHAR(20) NOT NULL,
    agency_overriding_amount NUMERIC(10,2) NOT NULL,
    agency_overriding_rate NUMERIC(5,4) DEFAULT 0.2000,
    total_office_revenue NUMERIC(10,2) NOT NULL,
    is_zero_offset INT DEFAULT 0,
    reconciliation_status VARCHAR(30) DEFAULT 'MATCHED_IN_ACCOUNT_57',
    notes TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (policy_number) REFERENCES policies(policy_number)
);

CREATE TABLE IF NOT EXISTS account_57_transactions (
    transaction_id VARCHAR(50) PRIMARY KEY,
    transaction_date VARCHAR(20) NOT NULL,
    iso_date VARCHAR(20),
    description VARCHAR(250) NOT NULL,
    branch_category VARCHAR(50) NOT NULL,
    debit_amount NUMERIC(12,2) DEFAULT 0.00,
    credit_amount NUMERIC(12,2) DEFAULT 0.00,
    running_balance NUMERIC(12,2) NOT NULL,
    matched_statement_month VARCHAR(10),
    is_reconciled INT DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS monthly_reconciliations (
    reconciliation_id VARCHAR(50) PRIMARY KEY,
    statement_month VARCHAR(10) NOT NULL UNIQUE,
    statement_producer_comm NUMERIC(10,2) NOT NULL,
    statement_agency_overriding NUMERIC(10,2) NOT NULL,
    statement_total_amount NUMERIC(10,2) NOT NULL,
    account_57_release_date VARCHAR(20),
    account_57_release_month VARCHAR(20),
    account_57_released_amount NUMERIC(10,2) NOT NULL,
    variance_amount NUMERIC(10,2) DEFAULT 0.00,
    match_status VARCHAR(30) DEFAULT 'PERFECT_MATCH',
    notes TEXT,
    verified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
