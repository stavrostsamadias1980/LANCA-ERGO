import sqlite3, pandas as pd, pdfplumber, os, sys

def seed_full_producers(db_path="ergo_statements.db"):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS producers_catalog (
        producer_code TEXT PRIMARY KEY,
        ergo_code TEXT,
        full_name TEXT NOT NULL,
        partner_type TEXT,
        partner_type_label TEXT,
        role TEXT,
        hierarchy TEXT,
        tier TEXT,
        manager TEXT,
        phone TEXT,
        email TEXT,
        address TEXT,
        nomos TEXT,
        comm_cat TEXT,
        status TEXT DEFAULT 'Ενεργός',
        commission_rate REAL DEFAULT 25.0,
        notes TEXT
    );
    """)

    ergo_tree_map = {}
    if os.path.exists('agencies_trees.pdf'):
        try:
            with pdfplumber.open('agencies_trees.pdf') as pdf:
                for page in pdf.pages:
                    for table in page.extract_tables():
                        for row in table:
                            c = [x.replace('\n', ' ').strip() if x else '' for x in row]
                            if c and c[0].isdigit():
                                ergo_code = c[1]
                                name = c[2]
                                clean_k = name.upper().replace(' ', '').replace(',', '').replace('.', '')
                                ergo_tree_map[clean_k] = {
                                    'ergo_code': ergo_code,
                                    'name': name,
                                    'role': c[4] if len(c)>4 else '',
                                    'comm_cat': c[5] if len(c)>5 else '',
                                    'nomos': c[6] if len(c)>6 else '',
                                    'start_date': c[3] if len(c)>3 else ''
                                }
        except Exception as e:
            print("PDF tree error:", e)

    info_map = {}
    if os.path.exists('ΣΥΝΕΡΓΑΤΕΣ_plirofories.xlsx'):
        try:
            df_info = pd.read_excel('ΣΥΝΕΡΓΑΤΕΣ_plirofories.xlsx')
            for _, row in df_info.iterrows():
                k = str(row.get('Κωδικός', '')).strip()
                if k and k != 'nan':
                    info_map[k] = {
                        'phone': str(row.get('ΤΗΛΕΦΩΝΟ', '')).replace('nan', '').strip(),
                        'address': str(row.get('ΔΙΕΥΘΥΝΣΗ', '')).replace('nan', '').strip(),
                        'nomos': str(row.get('ΝΟΜΟΣ', '')).replace('nan', '').strip(),
                        'email': str(row.get('EMAIL', '')).replace('nan', '').strip()
                    }
        except Exception as e:
            print("Info error:", e)

    producers_list = []
    seen_codes = set()
    if os.path.exists('ΣΥΝΕΡΓΑΤΕΣ.xlsx'):
        try:
            df_syn = pd.read_excel('ΣΥΝΕΡΓΑΤΕΣ.xlsx')
            for _, row in df_syn.iterrows():
                code_raw = str(row.get('Κωδικός', '')).strip()
                if not code_raw or code_raw == 'nan':
                    continue
                name = str(row.get('Επωνυμία', '')).replace('nan', '').strip()
                manager = str(row.get('Υπεύθυνος', '')).replace('nan', '').strip()
                hierarchy = str(row.get('Ιεραρχία ', '')).replace('nan', '').strip()
                status = str(row.get('Ενέργεια', '')).replace('nan', '').strip() or 'Ενεργός'
                
                is_numeric = code_raw.isdigit()
                
                clean_name_key = name.upper().replace(' ', '').replace(',', '').replace('.', '').replace('-', '')
                ergo_match = None
                for k, v in ergo_tree_map.items():
                    if k in clean_name_key or clean_name_key in k:
                        ergo_match = v
                        break
                        
                if code_raw in ['1', '3375', '3375A', '3375Α'] or ('ΑΝΑΓΝΩΣΤΟΠΟΥΛΟΣ' in name.upper() and 'ΝΙΚ' in name.upper()):
                    ptype = 'AGENCY_MANAGER'
                    ptype_label = '👑 Agency Manager (ERGO 40071 / 1411)'
                    ergo_code = '40071 / 1411'
                    role = 'Agency Manager / Συντονιστής'
                    tier = 'Agency Overriding 20% & Προσωπική Παραγωγή'
                    rate = 20.0
                elif is_numeric:
                    ptype = 'SUBCODE_1411'
                    ptype_label = '🔹 Έμμεσος Υποκωδικός (Μέσω ERGO 1411)'
                    ergo_code = '1411'
                    role = hierarchy or 'Ασφαλιστικός Πράκτορας (Υποκωδικός)'
                    tier = 'Έμμεσος Υποκωδικός (25% - 29%)'
                    rate = 25.0
                else:
                    ptype = 'DIRECT_AGENT'
                    ptype_label = '🏢 Άμεσος Πράκτορας (Οργανωτική Ομάδα 40071)'
                    ergo_code = ergo_match['ergo_code'] if ergo_match else 'ERGO Portal'
                    role = (ergo_match['role'] if ergo_match else '') or hierarchy or 'Ασφαλιστικός Πράκτορας'
                    tier = 'Οργανωτική Ομάδα (25% - 29%)'
                    rate = 25.0
                    
                inf = info_map.get(code_raw, {})
                phone = inf.get('phone', '')
                address = inf.get('address', '')
                nomos = inf.get('nomos', '') or (ergo_match.get('nomos', '') if ergo_match else '')
                email = inf.get('email', '')
                comm_cat = ergo_match.get('comm_cat', '') if ergo_match else ''
                
                producers_list.append((
                    code_raw, ergo_code, name, ptype, ptype_label, role, hierarchy, tier, manager,
                    phone, email, address, nomos, comm_cat, status, rate, 'LANCA Office & ERGO Registry'
                ))
                seen_codes.add(code_raw)
        except Exception as e:
            print("Excel syn error:", e)

    # Also add remaining ERGO tree members
    for k, v in ergo_tree_map.items():
        tree_ergo_code = v['ergo_code']
        code_gen = f"ERGO-{tree_ergo_code}"
        if code_gen not in seen_codes:
            ptype = 'DIRECT_AGENT'
            ptype_label = '🏢 Άμεσος Πράκτορας (Οργανωτική Ομάδα 40071)'
            s_date = v.get('start_date', '')
            producers_list.append((
                code_gen, tree_ergo_code, v['name'], ptype, ptype_label, v['role'], 'Ασφαλιστικός Πράκτορας',
                'Οργανωτική Ομάδα (25% - 29%)', '40071 LANCA', '', '', '', v['nomos'], v['comm_cat'],
                'Ενεργός', 25.0, f"ERGO Tree Registered ({s_date})"
            ))
            seen_codes.add(code_gen)

    cur.execute("DELETE FROM producers_catalog;")
    cur.executemany("""
        INSERT OR REPLACE INTO producers_catalog
        (producer_code, ergo_code, full_name, partner_type, partner_type_label, role, hierarchy, tier, manager, phone, email, address, nomos, comm_cat, status, commission_rate, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    """, producers_list)
    conn.commit()
    conn.close()
    print(f"Successfully seeded {len(producers_list)} total producers in {db_path}!")

if __name__ == "__main__":
    seed_full_producers("ergo_statements.db")
    seed_full_producers("lanca_ergo_local.db")
