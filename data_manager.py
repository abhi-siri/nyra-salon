import os
import sqlite3
import pandas as pd
import requests
import io
from datetime import datetime

GOOGLE_DRIVE_VIEW_URL = "https://docs.google.com/spreadsheets/d/1xqPkIAPaeEmvwvSOA6MqE9hTaMECtGUC/edit?usp=sharing"
GOOGLE_DRIVE_EXPORT_URL = "https://docs.google.com/spreadsheets/d/1xqPkIAPaeEmvwvSOA6MqE9hTaMECtGUC/export?format=xlsx"

LOCAL_EXCEL_PATH = os.path.join(os.path.dirname(__file__), "nyra_data.xlsx")
DB_PATH = os.path.join(os.path.dirname(__file__), "nyra_salon.db")

def init_db():
    """Initialize local SQLite database with automatic column migration."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Daily entry table with full rich columns
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS daily_entry (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            s_no INTEGER,
            date TEXT NOT NULL,
            items TEXT,
            category TEXT,
            money_received REAL NOT NULL,
            cost REAL DEFAULT 0.0,
            profit REAL DEFAULT 0.0,
            payment_mode TEXT NOT NULL,
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Auto-migrate table columns if older schema exists
    cursor.execute("PRAGMA table_info(daily_entry)")
    existing_cols = [col[1] for col in cursor.fetchall()]
    
    for new_col, col_type in [
        ('category', 'TEXT'),
        ('cost', 'REAL DEFAULT 0.0'),
        ('profit', 'REAL DEFAULT 0.0'),
        ('notes', 'TEXT')
    ]:
        if new_col not in existing_cols:
            try:
                cursor.execute(f"ALTER TABLE daily_entry ADD COLUMN {new_col} {col_type}")
            except Exception as e:
                print(f"Column migration warning for {new_col}: {e}")
                
    # Inventory table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_name TEXT NOT NULL,
            quantity TEXT NOT NULL,
            notes TEXT
        )
    ''')
    
    conn.commit()
    
    # Seed database from excel if table is empty
    cursor.execute("SELECT COUNT(*) FROM daily_entry")
    if cursor.fetchone()[0] == 0 and os.path.exists(LOCAL_EXCEL_PATH):
        try:
            xls = pd.ExcelFile(LOCAL_EXCEL_PATH)
            dfs_to_combine = []
            for sname in ['Earnings', 'Daily Entry', 'Daily entry']:
                if sname in xls.sheet_names:
                    d = pd.read_excel(LOCAL_EXCEL_PATH, sheet_name=sname)
                    if not d.empty:
                        dfs_to_combine.append(d)
                        
            if dfs_to_combine:
                combined_df = pd.concat(dfs_to_combine, ignore_index=True)
                combined_df.columns = [str(c).strip() for c in combined_df.columns]
                
                s_no_counter = 1
                for idx, row in combined_df.iterrows():
                    m_rec = row.get('Money Received') if pd.notnull(row.get('Money Received')) else row.get('Money Received (Rs.)')
                    p_mode = row.get('Payment Mode')
                    items_val = row.get('Items') if pd.notnull(row.get('Items')) else row.get('Services')
                    
                    if pd.notnull(m_rec) or pd.notnull(p_mode) or pd.notnull(items_val):
                        date_val = str(row['Date']).split(' ')[0] if pd.notnull(row.get('Date')) else datetime.now().strftime('%Y-%m-%d')
                        items_str = str(items_val) if pd.notnull(items_val) else ''
                        cat_str = str(row.get('Category')) if pd.notnull(row.get('Category')) else ''
                        money_val = float(m_rec) if pd.notnull(m_rec) else 0.0
                        cost_val = float(row.get('Cost (Rs.)')) if pd.notnull(row.get('Cost (Rs.)')) else 0.0
                        profit_val = float(row.get('Profit (Rs.)')) if pd.notnull(row.get('Profit (Rs.)')) else (money_val - cost_val)
                        mode_val = str(p_mode) if pd.notnull(p_mode) else 'Cash'
                        s_no_val = int(row['S.No']) if pd.notnull(row.get('S.No')) else s_no_counter
                        notes_val = str(row.get('Notes')) if pd.notnull(row.get('Notes')) else ''
                        
                        cursor.execute('''
                            INSERT INTO daily_entry (s_no, date, items, category, money_received, cost, profit, payment_mode, notes)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (s_no_val, date_val, items_str, cat_str, money_val, cost_val, profit_val, mode_val, notes_val))
                        s_no_counter += 1
                conn.commit()
        except Exception as e:
            print(f"Error seeding daily entry: {e}")
            
    cursor.execute("SELECT COUNT(*) FROM inventory")
    if cursor.fetchone()[0] == 0 and os.path.exists(LOCAL_EXCEL_PATH):
        try:
            df_inv = pd.read_excel(LOCAL_EXCEL_PATH, sheet_name='Inventory')
            col1 = df_inv.columns[0]
            col2 = df_inv.columns[1]
            
            cursor.execute("INSERT INTO inventory (item_name, quantity) VALUES (?, ?)", (str(col1).strip(), str(col2).strip()))
            
            for idx, row in df_inv.iterrows():
                if pd.notnull(row[col1]) or pd.notnull(row[col2]):
                    item_name = str(row[col1]).strip() if pd.notnull(row[col1]) else ""
                    qty = str(row[col2]).strip() if pd.notnull(row[col2]) else ""
                    cursor.execute("INSERT INTO inventory (item_name, quantity) VALUES (?, ?)", (item_name, qty))
            conn.commit()
        except Exception as e:
            print(f"Error seeding inventory: {e}")
            
    conn.close()

def load_menu_list():
    """Load Price List & Margins from local Excel file or default source."""
    if os.path.exists(LOCAL_EXCEL_PATH):
        df = pd.read_excel(LOCAL_EXCEL_PATH, sheet_name='Price List & Margins')
    else:
        df_dict = fetch_from_google_drive()
        df = df_dict.get('Price List & Margins', pd.DataFrame())
    
    if 'Variant' in df.columns:
        df['Variant'] = df['Variant'].fillna('')
    if 'Notes' in df.columns:
        df['Notes'] = df['Notes'].fillna('')
    if 'Cost per Service (Rs.)' in df.columns:
        df['Cost per Service (Rs.)'] = df['Cost per Service (Rs.)'].fillna(0.0)
    if 'Profit (Rs.)' in df.columns:
        df['Profit (Rs.)'] = df['Profit (Rs.)'].fillna(df['Selling Price (Rs.)'] - df['Cost per Service (Rs.)'])
    if 'Margin %' in df.columns:
        df['Margin %'] = df['Margin %'].fillna(0.0)
        
    return df

def fetch_from_google_drive():
    """Fetch live data from Google Drive link export and automatically fallback 'Daily Entry' if 0 rows."""
    try:
        res = requests.get(GOOGLE_DRIVE_EXPORT_URL, timeout=10)
        if res.status_code == 200:
            bytes_data = io.BytesIO(res.content)
            xls = pd.ExcelFile(bytes_data)
            data_dict = {}
            for sheet in xls.sheet_names:
                data_dict[sheet] = pd.read_excel(bytes_data, sheet_name=sheet)
                
            # If 'Daily Entry' or 'Daily entry' is empty (0 rows), fallback to 'Earnings' sheet data!
            for entry_name in ['Daily Entry', 'Daily entry']:
                if entry_name in data_dict:
                    if data_dict[entry_name].empty and 'Earnings' in data_dict and not data_dict['Earnings'].empty:
                        data_dict[entry_name] = data_dict['Earnings'].copy()
                        
            # Ensure 'Daily Entry' exists in data_dict
            if 'Daily Entry' not in data_dict and 'Earnings' in data_dict:
                data_dict['Daily Entry'] = data_dict['Earnings'].copy()
                
            return data_dict
        else:
            print(f"Failed to fetch from Google Drive: {res.status_code}")
            return {}
    except Exception as e:
        print(f"Error fetching from Google Drive: {e}")
        return {}

def get_earnings_df():
    """Get all daily entries from SQLite database."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query('''
        SELECT id, s_no as 'S.No', date as 'Date', items as 'Items/Services', category as 'Category', 
               money_received as 'Money Received (Rs.)', cost as 'Cost (Rs.)', profit as 'Profit (Rs.)', 
               payment_mode as 'Payment Mode', notes as 'Notes', created_at as 'Created At' 
        FROM daily_entry ORDER BY date DESC, id DESC
    ''', conn)
    conn.close()
    return df

def add_earning_entry(date_str, items_str, money_received, payment_mode, category="", cost=0.0, profit=0.0, notes="", webhook_url=None):
    """Add a new daily entry into SQLite database and optional Google Sheet webhook."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("SELECT MAX(s_no) FROM daily_entry")
    max_sno = cursor.fetchone()[0]
    next_sno = (max_sno + 1) if max_sno is not None else 1
    
    if profit == 0.0 and cost > 0.0:
        profit = money_received - cost
        
    cursor.execute('''
        INSERT INTO daily_entry (s_no, date, items, category, money_received, cost, profit, payment_mode, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (next_sno, date_str, items_str, category, float(money_received), float(cost), float(profit), payment_mode, notes))
    
    conn.commit()
    conn.close()
    
    # Optional Webhook push to Google Sheet
    if webhook_url and webhook_url.strip():
        try:
            payload = {
                "s_no": next_sno,
                "date": date_str,
                "items": items_str,
                "category": category,
                "money_received": float(money_received),
                "cost": float(cost),
                "profit": float(profit),
                "payment_mode": payment_mode,
                "notes": notes
            }
            requests.post(webhook_url.strip(), json=payload, timeout=5)
        except Exception as ex:
            print(f"Webhook push error: {ex}")

def delete_earning_entry(entry_id):
    """Delete an entry by ID."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM daily_entry WHERE id = ?", (entry_id,))
    conn.commit()
    conn.close()

def get_inventory_df():
    """Get all inventory items."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT id, item_name as 'Item Name', quantity as 'Quantity/Stock', notes as 'Notes' FROM inventory", conn)
    conn.close()
    return df

def add_inventory_item(item_name, quantity, notes=""):
    """Add a new inventory item."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO inventory (item_name, quantity, notes) VALUES (?, ?, ?)", (item_name, quantity, notes))
    conn.commit()
    conn.close()

def generate_excel_export():
    """Generate Excel file bytes matching the exact format with rich 'Daily Entry' sheet."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Instructions
        if os.path.exists(LOCAL_EXCEL_PATH):
            try:
                inst_df = pd.read_excel(LOCAL_EXCEL_PATH, sheet_name='Instructions')
                inst_df.to_excel(writer, sheet_name='Instructions', index=False)
            except Exception:
                pd.DataFrame({'Instructions': ['NYRA Unisex Salon Price List & Daily Entry Ledger']}).to_excel(writer, sheet_name='Instructions', index=False)
        else:
            pd.DataFrame({'Instructions': ['NYRA Unisex Salon Price List & Daily Entry Ledger']}).to_excel(writer, sheet_name='Instructions', index=False)
            
        # Sheet 2: Price List & Margins
        menu_df = load_menu_list()
        menu_df.to_excel(writer, sheet_name='Price List & Margins', index=False)
        
        # Sheet 3: Daily Entry (Full Rich Columns)
        daily_df = get_earnings_df()
        export_daily = daily_df[['S.No', 'Date', 'Items/Services', 'Category', 'Money Received (Rs.)', 'Cost (Rs.)', 'Profit (Rs.)', 'Payment Mode', 'Notes']].copy()
        export_daily.columns = ['S.No', 'Date', 'Items', 'Category', 'Money Received', 'Cost', 'Profit', 'Payment Mode', 'Notes']
        export_daily.to_excel(writer, sheet_name='Daily Entry', index=False)
        
        # Sheet 4: Category Summary
        cat_summary = menu_df.groupby('Category').agg(
            Number_of_Services=('Service', 'count'),
            Average_Selling_Price=('Selling Price (Rs.)', 'mean'),
            Average_Cost=('Cost per Service (Rs.)', 'mean'),
            Average_Profit=('Profit (Rs.)', 'mean'),
            Total_Profit_Sold_Once=('Profit (Rs.)', 'sum')
        ).reset_index()
        cat_summary.columns = [
            'Category', 'Number of Services', 'Average Selling Price (Rs.)', 
            'Average Cost per Service (Rs.)', 'Average Profit per Service (Rs.)', 
            'Total Profit if Sold Once Each (Rs.)'
        ]
        cat_summary.to_excel(writer, sheet_name='Category Summary', index=False)
        
        # Sheet 5: Inventory
        inv_df = get_inventory_df()
        export_inv = inv_df[['Item Name', 'Quantity/Stock']].copy()
        export_inv.to_excel(writer, sheet_name='Inventory', index=False)
        
    output.seek(0)
    return output.getvalue()
