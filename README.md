# ✂️ NYRA Unisex Salon - Streamlit Application

A complete web application for **NYRA Unisex Salon** to track daily service entries, earnings, work performed, service menu pricing with profit margins, inventory levels, and live synchronization with Google Drive Excel sheets.

---

## 🌟 Key Features

1. **📝 Daily Entry & Work Logger**:
   - Record daily services rendered by selecting items directly from the catalog.
   - Enter money received, payment mode (`UPI`, `Cash`, `Online`, `Card`), and service notes.
   - Calculates totals automatically.

2. **📊 Earnings & Work Done Analytics**:
   - Filter earnings by custom date ranges and payment methods.
   - View KPI metric cards (Total Revenue, UPI vs Cash breakdown, Average ticket value).
   - Interactive line and donut charts powered by Plotly.
   - Manage and inspect past transaction logs.

3. **📋 Menu Catalog & Price / Margin Calculator**:
   - Interactive searchable catalog of 160+ salon services across Threading, Waxing, Facials, Hair Spa, Color, Makeup, and Packages.
   - Shows Selling Price, Cost per Service, Profit, Margin %, and Notes.
   - Built-in **Quick Quote Builder** to calculate estimates for clients on the fly.

4. **📦 Salon Inventory Manager**:
   - Track salon supplies stock (wax, facial kits, thread, colors, tools).
   - Add new items or edit stock quantities.

5. **☁️ Google Drive Sync & Excel Export**:
   - Direct link to your [Google Drive Spreadsheet](https://docs.google.com/spreadsheets/d/1xqPkIAPaeEmvwvSOA6MqE9hTaMECtGUC/edit?usp=sharing).
   - One-click fetch to import live Google Sheet data.
   - One-click download of a formatted 5-sheet `.xlsx` file containing all salon records.

---

## 🚀 How to Run Locally

### 1. Prerequisite
Ensure Python 3.8+ is installed.

### 2. Install Dependencies
Open your terminal / command prompt in this directory and run:
```bash
pip install -r requirements.txt
```

### 3. Launch App
```bash
streamlit run app.py
```
The app will open automatically in your browser at `http://localhost:8501`.

---

## 🌐 How to Host / Deploy on Streamlit Cloud (Free)

1. **Push to GitHub**:
   - Create a GitHub repository (e.g. `nyra-salon-app`).
   - Upload all files from this directory (`app.py`, `data_manager.py`, `nyra_data.xlsx`, `requirements.txt`, `.streamlit/config.toml`).

2. **Deploy on Streamlit Community Cloud**:
   - Go to [share.streamlit.io](https://share.streamlit.io) and log in with your GitHub account.
   - Click **"New app"**.
   - Select your repository (`nyra-salon-app`), branch (`main`), and main file path (`app.py`).
   - Click **"Deploy!"**.

Your NYRA Unisex Salon site will be live on a custom Streamlit link (e.g., `https://nyra-salon.streamlit.app`) accessible on any phone, tablet, or desktop!
