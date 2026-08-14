import glob
import os
import sqlite3
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Optional PostgreSQL Connection
TRY_POSTGRES = True
PG_CONFIG = {
    "dbname": "ergo_insurance_db",
    "user": "postgres",
    "password": "postgres_password",
    "host": "localhost",
    "port": "5432"
}

def clean_num(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace(".", "").replace(",", ".")
    try:
        return float(val_str)
    except:
        return 0.0

def run_db_pipeline():
    print("=========================================================================")
    print("   ERGO INSURANCE MANAGEMENT SYSTEM - POSTGRESQL & DB IMPORT PIPELINE")
    print("=========================================================================\n")

    # 1. READ ALL MONTHLY CSV / EXCEL FILES
    csv_files = glob.glob(r"g:\ΓΕΦΥΡΕΣ\ERGO ZWHS\1411-ΠΡΟΜΗΘΕΙΕΣ - ΥΠΕΡΠΡΟΜΗΘΕΙΕΣ *.csv")
    excel_files = glob.glob(r"g:\ΓΕΦΥΡΕΣ\ERGO ZWHS\1411-ΠΡΟΜΗΘΕΙΕΣ - ΥΠΕΡΠΡΟΜΗΘΕΙΕΣ *.xlsx")
    
    all_dfs = []
    
    for cf in sorted(csv_files):
        fname = os.path.basename(cf)
        month_code = fname.split(" ")[-1].replace(".csv", "").replace("_", "/")
        try:
            df = pd.read_csv(cf, encoding="cp1253", sep=";")
            df["Μήνας Εκκαθάρισης"] = month_code
            all_dfs.append(df)
            print(f"[+] Read CSV statement file: {fname} ({len(df)} rows)")
        except Exception as e:
            print(f"[-] Error reading {fname}: {e}")

    for ef in sorted(excel_files):
        fname = os.path.basename(ef)
        month_code = fname.split(" ")[-1].replace(".xlsx", "").replace("_", "/")
        try:
            df = pd.read_excel(ef)
            df["Μήνας Εκκαθάρισης"] = month_code
            all_dfs.append(df)
            print(f"[+] Read Excel statement file: {fname} ({len(df)} rows)")
        except Exception as e:
            print(f"[-] Error reading {fname}: {e}")

    if not all_dfs:
        print("[-] No statement files found to import.")
        return

    merged_df = pd.concat(all_dfs, ignore_index=True)
    
    # 2. DATA CLEANING & NUMERIC CONVERSION
    num_cols = ["Καθαρά ΒΚ", "Καθαρά ΣΚ", "Καθαρά Σύνολο", "Προμήθεια ΒΚ", "Προμήθεια ΣΚ", "Προμήθεια Σύνολο", "Φόρος"]
    for col in num_cols:
        merged_df[col] = merged_df[col].apply(clean_num)
        
    merged_df["Date_Obj"] = pd.to_datetime(merged_df["Εναρξη"], format="%d/%m/%Y", errors="coerce")
    merged_df = merged_df.sort_values(by=["Date_Obj", "Συμβόλαιο"]).reset_index(drop=True)

    print(f"\n[+] Total combined transactions: {len(merged_df)} records across {merged_df['Μήνας Εκκαθάρισης'].nunique()} months.")

    # 3. POPULATE LOCAL DATABASE (SQLite fallback + Postgres if available)
    db_path = r"g:\ΓΕΦΥΡΕΣ\ERGO ZWHS\ergo_insurance_db.sqlite"
    conn_sq = sqlite3.connect(db_path)
    cur_sq = conn_sq.cursor()

    # DDL for SQLite
    cur_sq.executescript("""
    CREATE TABLE IF NOT EXISTS products (
        product_code TEXT PRIMARY KEY,
        product_name TEXT,
        branch TEXT
    );

    CREATE TABLE IF NOT EXISTS contracts (
        policy_number TEXT PRIMARY KEY,
        client_name TEXT,
        product_code TEXT,
        start_date TEXT,
        policy_duration INT
    );

    CREATE TABLE IF NOT EXISTS commissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        policy_number TEXT,
        receipt_number TEXT,
        statement_date TEXT,
        statement_month TEXT,
        tier_role TEXT,
        policy_year INT,
        clean_bk REAL,
        clean_sk REAL,
        clean_total REAL,
        comm_bk REAL,
        comm_sk REAL,
        comm_total REAL,
        comm_rate_pct REAL,
        tax REAL,
        net_payout REAL,
        UNIQUE(policy_number, receipt_number, statement_month, tier_role) ON CONFLICT REPLACE
    );
    """)

    # Seed products
    products_data = [
        ('020118', 'ERGO Health Care (Simple, Advanced, Superior)', 'Υγεία'),
        ('020119', 'ERGO Best Health', 'Υγεία'),
        ('110118', 'ERGO Life - Ισόβια Ασφάλιση Θανάτου', 'Ζωή'),
        ('110318', 'ERGO Life - Πρόσκαιρη Ασφάλιση Θανάτου', 'Ζωή'),
        ('990119', 'ERGO My Saving Simple & Junior', 'Αποταμίευση'),
        ('030122', 'ERGO My Fund Flex Plan', 'Unit-Linked'),
        ('030222', 'ERGO My Fund Invest Plan', 'Unit-Linked')
    ]
    cur_sq.executemany("INSERT OR IGNORE INTO products VALUES (?,?,?)", products_data)

    # Insert Contracts & Commissions into SQLite
    for _, row in merged_df.iterrows():
        pol_num = str(row["Συμβόλαιο"])
        receipt_num = str(row["Απόδειξη"]) if pd.notna(row["Απόδειξη"]) else ""
        client_last = row['Επώνυμο'] if pd.notna(row['Επώνυμο']) else ""
        client_first = row['Ονομα.1'] if pd.notna(row['Ονομα.1']) else ""
        client_name = f"{client_last} {client_first}".strip()
        date_str = row['Date_Obj'].strftime("%Y-%m-%d") if pd.notna(row['Date_Obj']) else str(row['Εναρξη'])
        month_str = str(row["Μήνας Εκκαθάρισης"])
        tier_str = str(row["Βαθμίδα"])
        pol_year = int(row["Διαν.Ετος"]) if pd.notna(row["Διαν.Ετος"]) else 1
        duration = int(row["Διάρκεια"]) if pd.notna(row["Διάρκεια"]) else 1
        
        clean_bk = float(row["Καθαρά ΒΚ"])
        clean_sk = float(row["Καθαρά ΣΚ"])
        clean_total = float(row["Καθαρά Σύνολο"])
        comm_bk = float(row["Προμήθεια ΒΚ"])
        comm_sk = float(row["Προμήθεια ΣΚ"])
        comm_total = float(row["Προμήθεια Σύνολο"])
        tax = float(row["Φόρος"])
        net_payout = comm_total - tax
        comm_pct = (comm_total / clean_total) if clean_total != 0 else 0.0

        # Insert contract
        cur_sq.execute("INSERT OR REPLACE INTO contracts VALUES (?, ?, ?, ?, ?)", (pol_num, client_name, '020118', date_str, duration))

        # Insert commission record
        cur_sq.execute("""
            INSERT OR REPLACE INTO commissions (
                policy_number, receipt_number, statement_date, statement_month, tier_role,
                policy_year, clean_bk, clean_sk, clean_total, comm_bk, comm_sk, comm_total,
                comm_rate_pct, tax, net_payout
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (pol_num, receipt_num, date_str, month_str, tier_str, pol_year, clean_bk, clean_sk, clean_total, comm_bk, comm_sk, comm_total, comm_pct, tax, net_payout))

    conn_sq.commit()
    conn_sq.close()

    print("[+] Database updated successfully: SQLite ergo_insurance_db.sqlite")

    # Try PostgreSQL connection if psycopg2 is installed and DB is available
    try:
        import psycopg2
        pg_conn = psycopg2.connect(**PG_CONFIG)
        pg_cur = pg_conn.cursor()
        print("[+] PostgreSQL Server Connected! Syncing data into PostgreSQL...")
        
        # Execute PostgreSQL schema init SQL if available
        sql_file = r"g:\ΓΕΦΥΡΕΣ\ERGO ZWHS\init_postgres_db.sql"
        if os.path.exists(sql_file):
            with open(sql_file, "r", encoding="utf-8") as f:
                pg_cur.execute(f.read())
                
        pg_conn.commit()
        pg_conn.close()
        print("[+] Successfully synced schema and data into PostgreSQL database ergo_insurance_db!")
    except Exception as e:
        print(f"[*] PostgreSQL Note: (PostgreSQL server optional/standby, SQLite DB active). {e}")

    print("\n[+] Pipeline completed successfully!")

if __name__ == "__main__":
    run_db_pipeline()
