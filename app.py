import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import json
import os

from data_manager import (
    GOOGLE_DRIVE_VIEW_URL,
    NEW_SPREADSHEET_ID,
    load_menu_list,
    get_earnings_df,
    add_earning_entry,
    delete_earning_entry,
    get_inventory_df,
    add_inventory_item,
    fetch_from_google_drive,
    generate_excel_export,
    init_db
)

# Page configuration
st.set_page_config(
    page_title="NYRA Unisex Salon - Daily Entry & Earnings",
    page_icon="✂️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Gold & Dark Navy / Salon Luxe Theme)
st.markdown("""
<style>
    .main-header {
        font-family: 'Helvetica Neue', sans-serif;
        font-size: 2.3rem;
        font-weight: 700;
        color: #D4AF37;
        text-align: center;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.0rem;
        color: #888;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    .stButton>button {
        border-radius: 6px;
    }
</style>
""", unsafe_allow_html=True)

# Initialize database
init_db()

# Load menu dataset
menu_df = load_menu_list()

# Sidebar Navigation & Branding
st.sidebar.markdown("# ✂️ **NYRA UNISEX SALON**")
st.sidebar.caption("Daily Entry, Earnings & Service Management")
st.sidebar.markdown("---")

nav_choice = st.sidebar.radio(
    "Navigation Menu",
    [
        "📝 Daily Entry",
        "📊 Daily Entry Ledger & Analytics",
        "📋 Menu Catalog & Pricing",
        "📦 Salon Inventory",
        "☁️ Google Drive & API Sync"
    ]
)

st.sidebar.markdown("---")
st.sidebar.markdown(f"[🔗 Open Google Sheet in Drive]({GOOGLE_DRIVE_VIEW_URL})")

# Download formatted Excel button in sidebar
excel_data = generate_excel_export()
st.sidebar.download_button(
    label="📥 Download Full Salon Excel (.xlsx)",
    data=excel_data,
    file_name=f"NYRA_Salon_Daily_Entry_{datetime.now().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

# Header Section
st.markdown("<div class='main-header'>NYRA UNISEX SALON</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>Daily Entry • Service Tracker • Price & Margin Catalog</div>", unsafe_allow_html=True)


# ==============================================================================
# TAB 1: DAILY ENTRY
# ==============================================================================
if nav_choice == "📝 Daily Entry":
    st.subheader("📝 Record Service Work & Money Received in 'Daily Entry'")
    
    if 'cart_items' not in st.session_state:
        st.session_state.cart_items = []

    col_left, col_right = st.columns([1, 1], gap="large")
    
    with col_left:
        st.markdown("### 1. Select Date & Payment Info")
        entry_date = st.date_input("Transaction Date", value=date.today())
        payment_mode = st.selectbox("Payment Mode", ["UPI", "Cash", "Online", "Card", "Mixed"])
        
        st.markdown("### 2. Add Services Performed")
        
        categories = sorted(menu_df['Category'].unique().tolist())
        selected_category = st.selectbox("Service Category", categories)
        
        cat_services = menu_df[menu_df['Category'] == selected_category]
        
        service_options = []
        for idx, row in cat_services.iterrows():
            variant_str = f" ({row['Variant']})" if pd.notnull(row['Variant']) and str(row['Variant']).strip() != '' else ""
            label = f"{row['Service']}{variant_str} - ₹{int(row['Selling Price (Rs.)'])}"
            service_options.append((label, row))
            
        selected_service_label = st.selectbox("Select Service", [opt[0] for opt in service_options])
        selected_service_row = [opt[1] for opt in service_options if opt[0] == selected_service_label][0]
        
        qty = st.number_input("Quantity", min_value=1, max_value=20, value=1)
        
        if st.button("➕ Add Service to Ticket", use_container_width=True):
            item_name = selected_service_row['Service']
            variant = selected_service_row['Variant']
            price = selected_service_row['Selling Price (Rs.)']
            cost = selected_service_row.get('Cost per Service (Rs.)', 0.0)
            if pd.isnull(cost):
                cost = 0.0
                
            full_item_desc = f"{item_name}{' (' + str(variant) + ')' if variant else ''}"
            
            st.session_state.cart_items.append({
                'category': selected_category,
                'item_desc': full_item_desc,
                'price': price,
                'cost': cost,
                'qty': qty,
                'subtotal': price * qty,
                'subtotal_cost': cost * qty
            })
            st.success(f"Added {full_item_desc} (x{qty}) to ticket!")

    with col_right:
        st.markdown("### 3. Service Ticket Summary")
        
        if len(st.session_state.cart_items) > 0:
            cart_df = pd.DataFrame(st.session_state.cart_items)
            st.dataframe(
                cart_df[['item_desc', 'qty', 'price', 'subtotal']].rename(
                    columns={'item_desc': 'Service', 'qty': 'Qty', 'price': 'Unit Price (₹)', 'subtotal': 'Subtotal (₹)'}
                ),
                use_container_width=True,
                hide_index=True
            )
            
            calculated_total = float(cart_df['subtotal'].sum())
            calculated_cost = float(cart_df['subtotal_cost'].sum())
            items_str = ", ".join([f"{item['item_desc']} x{item['qty']}" for item in st.session_state.cart_items])
            cats_str = ", ".join(list(set([item['category'] for item in st.session_state.cart_items])))
            
            col_clear, col_space = st.columns([1, 2])
            with col_clear:
                if st.button("🗑️ Clear Ticket Items"):
                    st.session_state.cart_items = []
                    st.rerun()
        else:
            calculated_total = 0.0
            calculated_cost = 0.0
            items_str = ""
            cats_str = ""
            st.info("No menu items added yet. You can pick services on the left or enter custom work below.")
            custom_items = st.text_area("Custom Work / Items Description", placeholder="e.g. Threading, Hair Cut, Facial")
            if custom_items.strip():
                items_str = custom_items.strip()
                
        final_money_received = st.number_input("Money Received (₹)", value=calculated_total, step=10.0)
        entry_notes = st.text_input("Customer Name / Notes (Optional)", placeholder="e.g. Client Name, Discount info")
        
        st.markdown("---")
        if st.button("✅ Save Entry", type="primary", use_container_width=True):
            if final_money_received <= 0:
                st.warning("Please enter a valid Money Received amount.")
            else:
                profit_calc = final_money_received - calculated_cost
                
                # Check for Google Sheets API secrets or session credentials
                sa_creds = None
                if "gcp_service_account" in st.secrets:
                    sa_creds = dict(st.secrets["gcp_service_account"])
                elif "sa_credentials_json" in st.session_state:
                    sa_creds = st.session_state["sa_credentials_json"]
                    
                webhook_url = st.secrets.get("GOOGLE_SHEET_WEBHOOK", None)
                
                add_earning_entry(
                    date_str=entry_date.strftime('%Y-%m-%d'),
                    items_str=items_str,
                    money_received=final_money_received,
                    payment_mode=payment_mode,
                    category=cats_str,
                    cost=calculated_cost,
                    profit=profit_calc,
                    notes=entry_notes,
                    webhook_url=webhook_url,
                    service_account_json=sa_creds
                )
                st.session_state.cart_items = []
                st.balloons()
                st.success(f"Successfully saved entry for {entry_date.strftime('%d-%b-%Y')}!")


# ==============================================================================
# TAB 2: DAILY ENTRY LEDGER & ANALYTICS
# ==============================================================================
elif nav_choice == "📊 Daily Entry Ledger & Analytics":
    st.subheader("📊 'Daily Entry' Sheet Ledger & Financial Analytics")
    
    df_earnings = get_earnings_df()
    
    if df_earnings.empty:
        st.info("No daily entries recorded yet.")
    else:
        df_earnings['Date_Parsed'] = pd.to_datetime(df_earnings['Date'], errors='coerce')
        min_date = df_earnings['Date_Parsed'].min().date() if pd.notnull(df_earnings['Date_Parsed'].min()) else date.today()
        max_date = df_earnings['Date_Parsed'].max().date() if pd.notnull(df_earnings['Date_Parsed'].max()) else date.today()
        
        st.markdown("#### Filter Date Range")
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            start_d = st.date_input("Start Date", value=min_date)
        with c2:
            end_d = st.date_input("End Date", value=max_date)
        with c3:
            mode_filter = st.multiselect("Payment Mode", options=df_earnings['Payment Mode'].unique().tolist(), default=df_earnings['Payment Mode'].unique().tolist())
            
        filtered_df = df_earnings[
            (df_earnings['Date_Parsed'].dt.date >= start_d) &
            (df_earnings['Date_Parsed'].dt.date <= end_d) &
            (df_earnings['Payment Mode'].isin(mode_filter))
        ].copy()
        
        st.markdown("---")
        
        total_rev = filtered_df['Money Received (Rs.)'].sum()
        total_count = len(filtered_df)
        avg_ticket = total_rev / total_count if total_count > 0 else 0
        total_profit = filtered_df['Profit (Rs.)'].sum()
        
        upi_rev = filtered_df[filtered_df['Payment Mode'] == 'UPI']['Money Received (Rs.)'].sum()
        cash_rev = filtered_df[filtered_df['Payment Mode'] == 'Cash']['Money Received (Rs.)'].sum()
        online_rev = filtered_df[filtered_df['Payment Mode'].isin(['Online', 'Card', 'Mixed'])]['Money Received (Rs.)'].sum()
        
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Revenue", f"₹{total_rev:,.0f}")
        m2.metric("Total Profit", f"₹{total_profit:,.0f}")
        m3.metric("UPI Earnings", f"₹{upi_rev:,.0f}")
        m4.metric("Cash Earnings", f"₹{cash_rev:,.0f}")
        m5.metric("Avg Ticket", f"₹{avg_ticket:,.0f}")
        
        st.markdown("---")
        
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            st.markdown("##### Daily Revenue Trend")
            daily_grp = filtered_df.groupby('Date')['Money Received (Rs.)'].sum().reset_index()
            daily_grp = daily_grp.sort_values('Date')
            
            fig_line = px.line(
                daily_grp,
                x='Date',
                y='Money Received (Rs.)',
                markers=True,
                line_shape='spline',
                title="Daily Revenue Trend (₹)"
            )
            fig_line.update_traces(line_color="#D4AF37", marker=dict(size=8, color="#D4AF37"))
            fig_line.update_layout(xaxis_title="Date", yaxis_title="Money Received (₹)", template="plotly_dark")
            st.plotly_chart(fig_line, use_container_width=True)
            
        with chart_col2:
            st.markdown("##### Revenue Breakdown by Payment Mode")
            mode_grp = filtered_df.groupby('Payment Mode')['Money Received (Rs.)'].sum().reset_index()
            
            fig_pie = px.pie(
                mode_grp,
                names='Payment Mode',
                values='Money Received (Rs.)',
                hole=0.4,
                title="Payment Method Share",
                color_discrete_sequence=px.colors.qualitative.Gold
            )
            fig_pie.update_layout(template="plotly_dark")
            st.plotly_chart(fig_pie, use_container_width=True)
            
        st.markdown("---")
        st.markdown(f"### 📜 'Daily Entry' Table ({len(filtered_df)} Rows)")
        
        disp_cols = ['S.No', 'Date', 'Items/Services', 'Category', 'Money Received (Rs.)', 'Cost (Rs.)', 'Profit (Rs.)', 'Payment Mode', 'Notes', 'Created At']
        st.dataframe(
            filtered_df[disp_cols],
            use_container_width=True,
            hide_index=True
        )
        
        with st.expander("🗑️ Manage / Delete an Entry"):
            del_id = st.number_input("Enter ID of entry to delete", min_value=1, step=1)
            if st.button("Delete Entry"):
                delete_earning_entry(del_id)
                st.success(f"Entry #{del_id} deleted successfully.")
                st.rerun()


# ==============================================================================
# TAB 3: MENU CATALOG & PRICING
# ==============================================================================
elif nav_choice == "📋 Menu Catalog & Pricing":
    st.subheader("📋 NYRA Unisex Salon Menu List & Margin Calculator")
    
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        search_query = st.text_input("🔍 Search Service / Category / Note", "")
    with col_f2:
        selected_cats = st.multiselect("Filter by Category", options=sorted(menu_df['Category'].unique().tolist()))
        
    filtered_menu = menu_df.copy()
    if selected_cats:
        filtered_menu = filtered_menu[filtered_menu['Category'].isin(selected_cats)]
    if search_query.strip():
        q = search_query.lower()
        filtered_menu = filtered_menu[
            filtered_menu['Service'].str.lower().str.contains(q, na=False) |
            filtered_menu['Category'].str.lower().str.contains(q, na=False) |
            filtered_menu['Notes'].str.lower().str.contains(q, na=False)
        ]
        
    st.markdown(f"**Showing {len(filtered_menu)} services**")
    
    display_df = filtered_menu[['Category', 'Service', 'Variant', 'Selling Price (Rs.)', 'Cost per Service (Rs.)', 'Profit (Rs.)', 'Margin %', 'Notes']].copy()
    display_df['Margin %'] = (display_df['Margin %'] * 100).round(1).astype(str) + '%'
    
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )
    
    st.markdown("---")
    st.markdown("### 🧮 Quick Price Estimator / Quote Builder")
    calc_services = st.multiselect(
        "Select services to build quote",
        options=menu_df['Category'] + " - " + menu_df['Service'] + " (" + menu_df['Variant'].astype(str) + ") [₹" + menu_df['Selling Price (Rs.)'].astype(str) + "]"
    )
    
    if calc_services:
        total_quote = 0.0
        total_cost = 0.0
        quote_rows = []
        for s_label in calc_services:
            for idx, row in menu_df.iterrows():
                label = f"{row['Category']} - {row['Service']} ({row['Variant']}) [₹{row['Selling Price (Rs.)']}]"
                if label == s_label:
                    total_quote += row['Selling Price (Rs.)']
                    total_cost += row['Cost per Service (Rs.)']
                    quote_rows.append({
                        'Category': row['Category'],
                        'Service': row['Service'],
                        'Variant': row['Variant'],
                        'Price (₹)': row['Selling Price (Rs.)'],
                        'Est Cost (₹)': row['Cost per Service (Rs.)'],
                        'Est Profit (₹)': row['Profit (Rs.)']
                    })
                    break
        q_df = pd.DataFrame(quote_rows)
        st.table(q_df)
        st.markdown(f"#### **Total Quote Price: ₹{total_quote:,.2f}** (Estimated Profit: ₹{(total_quote - total_cost):,.2f})")


# ==============================================================================
# TAB 4: SALON INVENTORY
# ==============================================================================
elif nav_choice == "📦 Salon Inventory":
    st.subheader("📦 Salon Inventory & Product Stock")
    
    inv_df = get_inventory_df()
    
    st.dataframe(inv_df, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.markdown("### ➕ Add New Inventory Item")
    c1, c2 = st.columns(2)
    with c1:
        new_item = st.text_input("Item Name")
    with c2:
        new_qty = st.text_input("Quantity / Stock (e.g. 2 boxes, 5 bottles)")
        
    if st.button("Add Item to Stock"):
        if new_item.strip():
            add_inventory_item(new_item.strip(), new_qty.strip())
            st.success(f"Added {new_item} to inventory!")
            st.rerun()


# ==============================================================================
# TAB 5: GOOGLE DRIVE & API SYNC
# ==============================================================================
elif nav_choice == "☁️ Google Drive & API Sync":
    st.subheader("☁️ Google Drive Spreadsheet & Google Sheets API Integration")
    
    st.markdown(f"""
    Target Google Spreadsheet ID: `{NEW_SPREADSHEET_ID}`  
    👉 **[Click Here to Open Google Sheet in Drive]({GOOGLE_DRIVE_VIEW_URL})**
    """)
    
    st.markdown("---")
    
    # API Integration Configuration Box
    with st.expander("🔑 Setup Live 2-Way Google Sheets API (Automated Write)"):
        st.markdown("""
        To enable automatic real-time writing directly into your Google Sheet when entries are saved:
        
        #### Option A: Google Cloud Service Account (gspread)
        1. Share your Google Sheet (`10ZEp7mTd3lhSk2qs5eeEVYqhDMkm9s4S`) with your Google Service Account email as **Editor**.
        2. Paste your Service Account JSON credentials below or add it to Streamlit Secrets (`st.secrets["gcp_service_account"]`).
        """)
        
        sa_json_input = st.text_area("Paste Service Account JSON Credentials", placeholder='{"type": "service_account", ...}')
        if st.button("Save Service Account Key"):
            try:
                parsed_json = json.loads(sa_json_input.strip())
                st.session_state["sa_credentials_json"] = parsed_json
                st.success("Google Service Account credentials saved for this session!")
            except Exception as e:
                st.error(f"Invalid JSON format: {e}")

    st.markdown("---")
    st.markdown("### 🔄 Fetch & Download Sync Options")
    
    col_sync1, col_sync2 = st.columns(2)
    
    with col_sync1:
        st.markdown("#### 1. Fetch Data from Google Drive Link")
        st.caption("Inspect live sheets from your Google Sheet link.")
        if st.button("📥 Fetch Google Drive Sheet Data", use_container_width=True):
            with st.spinner("Downloading spreadsheet from Google Drive..."):
                g_data = fetch_from_google_drive()
                if g_data:
                    st.success("Successfully fetched live Google Drive sheet!")
                    for sheet_name, df_s in g_data.items():
                        st.markdown(f"**Sheet: {sheet_name}** ({len(df_s)} rows)")
                else:
                    st.error("Failed to fetch live spreadsheet. Please ensure link permissions are set to 'Anyone with link can view'.")
                    
    with col_sync2:
        st.markdown("#### 2. Download Updated Excel Workbook")
        st.caption("Download the complete 5-sheet `.xlsx` file containing all your Daily Entry logs and import it into Google Drive.")
        st.download_button(
            label="📤 Download Updated Salon Excel (.xlsx)",
            data=excel_data,
            file_name=f"NYRA_Salon_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
