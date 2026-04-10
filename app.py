"""
Business Intelligence Dashboard – Real‑time analytics for companies.
Connect to SQL, Excel, CSV, visualize KPIs, sales trends, inventory, custom reports.
Multi‑language: English, Spanish, French, Haitian Creole.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import sqlite3
import os
import io
from datetime import datetime, timedelta
import time

# ------------------------------
# PAGE CONFIG & LOGIN
# ------------------------------
st.set_page_config(page_title="BI Dashboard – GlobalInternet.py", layout="wide")

def show_haitian_flag(width=100):
    st.image("https://flagcdn.com/w320/ht.png", width=width)

# Authentication
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔐 Login Required")
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        show_haitian_flag(150)
        st.markdown("<h2 style='text-align: center;'>Business Intelligence Dashboard</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center;'>by GlobalInternet.py</p>", unsafe_allow_html=True)
        password_input = st.text_input("Enter password to access", type="password")
        if st.button("Login"):
            if password_input == "20082010":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password. Access denied.")
    st.stop()

# ------------------------------
# AFTER LOGIN – MAIN APP
# ------------------------------
col_flag, col_title = st.columns([1, 3])
with col_flag:
    show_haitian_flag(120)
with col_title:
    st.markdown("<h1>📊 Business Intelligence Dashboard</h1>", unsafe_allow_html=True)
    st.markdown("*Real‑time analytics for companies – connect SQL, Excel, CSV*")

# ------------------------------
# SIDEBAR – INFO & LOGOUT
# ------------------------------
with st.sidebar:
    st.markdown("## 🇭🇹 GlobalInternet.py")
    show_haitian_flag(80)
    st.markdown("### BI Dashboard")
    st.markdown("---")
    st.markdown("**Founder & Developer:**")
    st.markdown("Gesner Deslandes")
    st.markdown("📞 **WhatsApp:** [509 4738-5663](https://wa.me/50947385663)")
    st.markdown("📧 **Email:** deslandes78@gmail.com")
    st.markdown("🌐 **Website:** [www.globalinternet.py](https://www.globalinternet.py)")
    st.markdown("---")
    st.markdown("### 💰 Price")
    st.markdown("**$1,200 USD** (one‑time license)")
    st.markdown("---")
    st.markdown("### © 2025 GlobalInternet.py")
    st.markdown("All Rights Reserved")
    st.markdown("---")
    if st.button("🚪 Logout", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

# ------------------------------
# MULTI-LANGUAGE SUPPORT
# ------------------------------
LANGUAGES = {
    "English": "en",
    "Español": "es",
    "Français": "fr",
    "Kreyòl Ayisyen": "ht"
}

TEXTS = {
    "en": {
        "data_source": "📂 Data Source",
        "upload_csv": "Upload CSV or Excel file",
        "use_sample": "Use sample data (sales & inventory)",
        "upload_sqlite": "Upload SQLite database file (.db)",
        "sample_data": "Sample Retail Data",
        "kpi_sales": "Total Sales",
        "kpi_orders": "Total Orders",
        "kpi_avg_order": "Avg Order Value",
        "kpi_inventory": "Total Inventory Value",
        "sales_trend": "Sales Trend Over Time",
        "top_products": "Top 5 Products by Revenue",
        "inventory_status": "Inventory by Category",
        "sales_by_region": "Sales by Region",
        "date_filter": "Date Range",
        "apply_filter": "Apply Filter",
        "download_report": "📥 Download Report (CSV)",
        "report_generated": "Report generated!",
        "no_data": "No data available. Please upload a file or use sample data."
    },
    "es": {
        "data_source": "📂 Fuente de datos",
        "upload_csv": "Subir archivo CSV o Excel",
        "use_sample": "Usar datos de ejemplo (ventas e inventario)",
        "upload_sqlite": "Subir archivo de base de datos SQLite (.db)",
        "sample_data": "Datos de ejemplo de ventas",
        "kpi_sales": "Ventas totales",
        "kpi_orders": "Pedidos totales",
        "kpi_avg_order": "Valor promedio de pedido",
        "kpi_inventory": "Valor total del inventario",
        "sales_trend": "Tendencia de ventas en el tiempo",
        "top_products": "Top 5 productos por ingresos",
        "inventory_status": "Inventario por categoría",
        "sales_by_region": "Ventas por región",
        "date_filter": "Rango de fechas",
        "apply_filter": "Aplicar filtro",
        "download_report": "📥 Descargar informe (CSV)",
        "report_generated": "¡Informe generado!",
        "no_data": "No hay datos disponibles. Suba un archivo o use datos de ejemplo."
    },
    "fr": {
        "data_source": "📂 Source de données",
        "upload_csv": "Télécharger un fichier CSV ou Excel",
        "use_sample": "Utiliser des données d'exemple (ventes et inventaire)",
        "upload_sqlite": "Télécharger une base de données SQLite (.db)",
        "sample_data": "Données de vente exemple",
        "kpi_sales": "Ventes totales",
        "kpi_orders": "Commandes totales",
        "kpi_avg_order": "Valeur moyenne de commande",
        "kpi_inventory": "Valeur totale de l'inventaire",
        "sales_trend": "Tendance des ventes dans le temps",
        "top_products": "Top 5 produits par chiffre d'affaires",
        "inventory_status": "Inventaire par catégorie",
        "sales_by_region": "Ventes par région",
        "date_filter": "Plage de dates",
        "apply_filter": "Appliquer le filtre",
        "download_report": "📥 Télécharger le rapport (CSV)",
        "report_generated": "Rapport généré !",
        "no_data": "Aucune donnée disponible. Téléchargez un fichier ou utilisez les données d'exemple."
    },
    "ht": {
        "data_source": "📂 Sous done",
        "upload_csv": "Telechaje fichye CSV oswa Excel",
        "use_sample": "Sèvi ak done egzanp (lavant ak envantè)",
        "upload_sqlite": "Telechaje baz done SQLite (.db)",
        "sample_data": "Done lavant egzanp",
        "kpi_sales": "Lavant total",
        "kpi_orders": "Kòmand total",
        "kpi_avg_order": "Valè mwayèn kòmand",
        "kpi_inventory": "Valè total envantè",
        "sales_trend": "Tandans lavant sou tan",
        "top_products": "Top 5 pwodwi pa revni",
        "inventory_status": "Envantè pa kategori",
        "sales_by_region": "Lavant pa rejyon",
        "date_filter": "Ranje dat",
        "apply_filter": "Aplike filt",
        "download_report": "📥 Telechaje rapò (CSV)",
        "report_generated": "Rapò kreye!",
        "no_data": "Pa gen done disponib. Telechaje yon fichye oswa itilize done egzanp."
    }
}

def get_text(key):
    lang = st.session_state.get("language", "en")
    return TEXTS[lang].get(key, key)

# Language selector
lang = st.sidebar.selectbox("🌐 Language", list(LANGUAGES.keys()))
st.session_state["language"] = LANGUAGES[lang]

# ------------------------------
# DATA LOADING FUNCTION
# ------------------------------
@st.cache_data
def load_sample_data():
    """Generate realistic sample sales and inventory data."""
    np.random.seed(42)
    dates = pd.date_range(start='2024-01-01', end='2025-03-31', freq='D')
    products = ['Laptop', 'Phone', 'Tablet', 'Monitor', 'Keyboard', 'Mouse', 'Headphones', 'Charger']
    regions = ['North', 'South', 'East', 'West', 'Central']
    categories = ['Electronics', 'Accessories', 'Peripherals']
    product_cat = {p: np.random.choice(categories) for p in products}
    
    sales_data = []
    for date in dates:
        for _ in range(np.random.randint(50, 200)):
            product = np.random.choice(products)
            quantity = np.random.randint(1, 10)
            price = np.random.uniform(20, 1200)
            sales_data.append({
                'date': date,
                'product': product,
                'category': product_cat[product],
                'region': np.random.choice(regions),
                'quantity': quantity,
                'price': price,
                'revenue': quantity * price
            })
    sales_df = pd.DataFrame(sales_data)
    
    # Inventory data
    inventory = []
    for product in products:
        inventory.append({
            'product': product,
            'category': product_cat[product],
            'stock_quantity': np.random.randint(20, 500),
            'unit_cost': np.random.uniform(15, 800),
            'reorder_level': np.random.randint(10, 50)
        })
    inventory_df = pd.DataFrame(inventory)
    inventory_df['inventory_value'] = inventory_df['stock_quantity'] * inventory_df['unit_cost']
    
    return sales_df, inventory_df

def load_from_csv(file):
    df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
    # Ensure date column exists
    if 'date' in df.columns:
        df['date'] = pd.to_datetime(df['date'])
    return df

def load_from_sqlite(file):
    conn = sqlite3.connect(file)
    # Try to find tables with sales and inventory
    tables = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table';", conn)
    sales_df = None
    inventory_df = None
    for table in tables['name']:
        df = pd.read_sql(f"SELECT * FROM '{table}';", conn)
        if 'date' in df.columns or 'revenue' in df.columns:
            sales_df = df
        if 'stock_quantity' in df.columns or 'inventory_value' in df.columns:
            inventory_df = df
    conn.close()
    return sales_df, inventory_df

# ------------------------------
# INITIALIZE DATA
# ------------------------------
if "sales_df" not in st.session_state:
    st.session_state.sales_df = None
if "inventory_df" not in st.session_state:
    st.session_state.inventory_df = None
if "filtered_sales" not in st.session_state:
    st.session_state.filtered_sales = None

# ------------------------------
# DATA SOURCE SELECTION
# ------------------------------
st.sidebar.markdown(f"## {get_text('data_source')}")
data_option = st.sidebar.radio(
    "Choose data source:",
    [get_text('sample_data'), get_text('upload_csv'), get_text('upload_sqlite')]
)

if data_option == get_text('sample_data'):
    if st.sidebar.button("Load Sample Data"):
        sales, inventory = load_sample_data()
        st.session_state.sales_df = sales
        st.session_state.inventory_df = inventory
        st.session_state.filtered_sales = sales.copy()
        st.rerun()

elif data_option == get_text('upload_csv'):
    uploaded_file = st.sidebar.file_uploader(get_text('upload_csv'), type=['csv', 'xlsx', 'xls'])
    if uploaded_file:
        df = load_from_csv(uploaded_file)
        st.session_state.sales_df = df
        st.session_state.inventory_df = None  # inventory may be separate
        st.session_state.filtered_sales = df.copy()
        st.rerun()

elif data_option == get_text('upload_sqlite'):
    uploaded_db = st.sidebar.file_uploader(get_text('upload_sqlite'), type=['db'])
    if uploaded_db:
        # Save temporarily
        with open("temp.db", "wb") as f:
            f.write(uploaded_db.getbuffer())
        sales, inventory = load_from_sqlite("temp.db")
        st.session_state.sales_df = sales
        st.session_state.inventory_df = inventory
        if sales is not None:
            st.session_state.filtered_sales = sales.copy()
        st.rerun()

# ------------------------------
# DASHBOARD DISPLAY
# ------------------------------
if st.session_state.sales_df is not None:
    sales = st.session_state.filtered_sales if st.session_state.filtered_sales is not None else st.session_state.sales_df
    inventory = st.session_state.inventory_df
    
    # Date filter
    if 'date' in sales.columns:
        sales['date'] = pd.to_datetime(sales['date'])
        min_date = sales['date'].min()
        max_date = sales['date'].max()
        date_range = st.sidebar.date_input(
            get_text('date_filter'),
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )
        if len(date_range) == 2:
            start, end = date_range
            mask = (sales['date'] >= pd.Timestamp(start)) & (sales['date'] <= pd.Timestamp(end))
            filtered = sales[mask].copy()
            st.session_state.filtered_sales = filtered
            st.rerun()
    
    # KPIs
    col1, col2, col3, col4 = st.columns(4)
    total_revenue = filtered['revenue'].sum() if 'revenue' in filtered.columns else 0
    total_orders = len(filtered)
    avg_order = total_revenue / total_orders if total_orders > 0 else 0
    total_inventory = inventory['inventory_value'].sum() if inventory is not None else 0
    
    with col1:
        st.metric(get_text('kpi_sales'), f"${total_revenue:,.0f}")
    with col2:
        st.metric(get_text('kpi_orders'), f"{total_orders:,}")
    with col3:
        st.metric(get_text('kpi_avg_order'), f"${avg_order:,.2f}")
    with col4:
        st.metric(get_text('kpi_inventory'), f"${total_inventory:,.0f}")
    
    # Charts
    tab1, tab2, tab3, tab4 = st.tabs(["📈 Sales Trend", "🏆 Top Products", "📦 Inventory", "🗺️ Regional Sales"])
    
    with tab1:
        if 'date' in filtered.columns:
            daily = filtered.groupby('date')['revenue'].sum().reset_index()
            fig = px.line(daily, x='date', y='revenue', title=get_text('sales_trend'))
            st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        if 'product' in filtered.columns:
            product_sales = filtered.groupby('product')['revenue'].sum().sort_values(ascending=False).head(5)
            fig = px.bar(x=product_sales.index, y=product_sales.values, title=get_text('top_products'))
            st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        if inventory is not None:
            inv_by_cat = inventory.groupby('category')['inventory_value'].sum().reset_index()
            fig = px.pie(inv_by_cat, values='inventory_value', names='category', title=get_text('inventory_status'))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Inventory data not available in this dataset.")
    
    with tab4:
        if 'region' in filtered.columns:
            region_sales = filtered.groupby('region')['revenue'].sum().reset_index()
            fig = px.bar(region_sales, x='region', y='revenue', title=get_text('sales_by_region'))
            st.plotly_chart(fig, use_container_width=True)
    
    # Download report
    st.markdown("---")
    if st.button(get_text('download_report')):
        csv = filtered.to_csv(index=False)
        st.download_button(
            label="📥 Click to download CSV",
            data=csv,
            file_name=f"bi_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
else:
    st.info(get_text('no_data'))
    if st.button("Load Sample Data Now"):
        sales, inventory = load_sample_data()
        st.session_state.sales_df = sales
        st.session_state.inventory_df = inventory
        st.session_state.filtered_sales = sales.copy()
        st.rerun()

# Footer
st.markdown("---")
st.markdown("📊 **Custom reports available** – contact us to tailor dashboards to your exact needs.")
