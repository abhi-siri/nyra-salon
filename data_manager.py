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
    """Initialize local SQLite database for earnings and custom menu/inventory tracking."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Earnings table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS earnings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            s_no INTEGER,
            date TEXT NOT NULL,
            items TEXT,
            money_received REAL NOT NULL,
            payment_mode TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
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
    cursor.execute("SELECT COUNT(*) FROM earnings")
    if cursor.fetchone()[0] == 0 and os.path.exists(LOCAL_EXCEL_PATH):
        try:
            df = pd.read_excel(LOCAL_EXCEL_PATH, sheet_name='Earnings')
            # Clean column names
            df.columns = [c.strip() for c in df.columns]
            for idx, row in df.iterrows():
                if pd.notnull(row.get('Money Received')) or pd.notnull(row.get('Payment Mode')) or pd.notnull(row.get('Items')):
                    date_val = str(row['Date']).split(' ')[0] if pd.notnull(row.get('Date')) else datetime.now().strftime('%Y-%m-%d')
                    items_val = str(row['Items']) if pd.notnull(row.get('Items')) else ''
                    money_val = float(row['Money Received']) if pd.notnull(row.get('Money Received')) else 0.0
                    mode_val = str(row['Payment Mode']) if pd.notnull(row.get('Payment Mode')) else 'Cash'
                    s_no_val = int(row['S.No']) if pd.notnull(row.get('S.No')) else (idx + 1)
                    
                    cursor.execute('''
                        INSERT INTO earnings (s_no, date, items, money_received, payment_mode)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (s_no_val, date_val, items_val, money_val, mode_val))
            conn.commit()
        except Exception as e:
            print(f"Error seeding earnings: {e}")
            
    cursor.execute("SELECT COUNT(*) FROM inventory")
    if cursor.fetchone()[0] == 0 and os.path.exists(LOCAL_EXCEL_PATH):
        try:
            df_inv = pd.read_excel(LOCAL_EXCEL_PATH, sheet_name='Inventory')
            col1 = df_inv.columns[0]
            col2 = df_inv.columns[1]
            
            # Header item
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
        # Fallback fetch from Google Drive URL
        df_dict = fetch_from_google_drive()
        df = df_dict.get('Price List & Margins', pd.DataFrame())
    
    # Fill NaN values for clean rendering
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
    """Fetch live data from Google Drive link export."""
    try:
        res = requests.get(GOOGLE_DRIVE_EXPORT_URL, timeout=10)
        if res.status_code == 200:
            bytes_data = io.BytesIO(res.content)
            xls = pd.ExcelFile(bytes_data)
            data_dict = {}
            for sheet in xls.sheet_names:
                data_dict[sheet] = pd.read_excel(bytes_data, sheet_name=sheet)
            return data_dict
        else:
            print(f"Failed to fetch from Google Drive: {res.status_code}")
            return {}
    except Exception as e:
        print(f"Error fetching from Google Drive: {e}")
        return {}

def get_earnings_df():
    """Get all daily earnings entries from SQLite database."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT id, s_no as 'S.No', date as 'Date', items as 'Items', money_received as 'Money Received', payment_mode as 'Payment Mode', created_at as 'Created At' FROM earnings ORDER BY date DESC, id DESC", conn)
    conn.close()
    return df

def add_earning_entry(date_str, items_str, money_received, payment_mode):
    """Add a new service transaction / daily earning entry."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Calculate next S.No
    cursor.execute("SELECT MAX(s_no) FROM earnings")
    max_sno = cursor.fetchone()[0]
    next_sno = (max_sno + 1) if max_sno is not None else 1
    
    cursor.execute('''
        INSERT INTO earnings (s_no, date, items, money_received, payment_mode)
        VALUES (?, ?, ?, ?, ?)
    ''', (next_sno, date_str, items_str, float(money_received), payment_mode))
    
    conn.commit()
    conn.close()

def delete_earning_entry(entry_id):
    """Delete an entry by ID."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM earnings WHERE id = ?", (entry_id,))
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

def update_inventory_item(item_id, item_name, quantity, notes=""):
    """Update existing inventory item."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("UPDATE inventory SET item_name = ?, quantity = ?, notes = ? WHERE id = ?", (item_name, quantity, notes, item_id))
    conn.commit()
    conn.close()

def generate_excel_export():
    """Generate Excel file bytes matching the exact 5-sheet format of the salon workbook."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Instructions
        if os.path.exists(LOCAL_EXCEL_PATH):
            try:
                inst_df = pd.read_excel(LOCAL_EXCEL_PATH, sheet_name='Instructions')
                inst_df.to_excel(writer, sheet_name='Instructions', index=False)
            except Exception:
                pd.DataFrame({'Instructions': ['NYRA Unisex Salon Price List & Earnings Ledger']}).to_excel(writer, sheet_name='Instructions', index=False)
        else:
            pd.DataFrame({'Instructions': ['NYRA Unisex Salon Price List & Earnings Ledger']}).to_excel(writer, sheet_name='Instructions', index=False)
            
        # Sheet 2: Price List & Margins
        menu_df = load_menu_list()
        menu_df.to_excel(writer, sheet_name='Price List & Margins', index=False)
        
        # Sheet 3: Earnings
        earnings_df = get_earnings_df()
        export_earnings = earnings_df[['S.No', 'Date', 'Items', 'Money Received', 'Payment Mode']].copy()
        export_earnings.to_excel(writer, sheet_name='Earnings', index=False)
        
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
