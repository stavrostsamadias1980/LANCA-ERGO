-- =============================================================================
-- ERGO INSURANCE MANAGEMENT SYSTEM - POSTGRESQL DATABASE SCHEMA (ergo_insurance_db)
-- LANCA Insurance Services
-- =============================================================================

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
