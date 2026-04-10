"""
Business Intelligence Dashboard – Real‑time analytics for companies.
Connect to SQL, Excel, CSV, visualize KPIs, sales trends, inventory, custom reports.
FULLY MULTI‑LANGUAGE: English, Spanish, French, Haitian Creole.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import sqlite3
import os
from datetime import datetime

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
# MULTI-LANGUAGE DICTIONARY
# ------------------------------
LANGUAGES = {
    "English": "en",
    "Español": "es",
    "Français": "fr",
    "Kreyòl Ayisyen": "ht"
}

TEXTS = {
    "en": {
        "app_title": "Business Intelligence Dashboard",
        "app_subtitle": "Real‑time analytics for companies – connect SQL, Excel, CSV",
        "sidebar_company": "GlobalInternet.py",
        "sidebar_product": "BI Dashboard",
        "founder": "Founder & Developer",
        "name": "Gesner Deslandes",
        "whatsapp": "WhatsApp",
        "email": "Email",
        "website": "Website",
        "price_label": "Price",
        "price_value": "$1,200 USD (one‑time license)",
        "copyright": "All Rights Reserved",
        "logout_btn": "Logout",
        "data_source": "📂 Data Source",
        "sample_data_option": "Sample Retail Data",
        "csv_option": "Upload CSV or Excel file",
        "sqlite_option": "Upload SQLite database (.db)",
        "load_sample_btn": "Load Sample Data",
        "upload_csv_btn": "Upload CSV/Excel",
        "upload_sqlite_btn": "Upload SQLite",
        "date_filter": "Date Range",
        "apply_filter": "Apply Filter",
        "kpi_total_sales": "Total Sales",
        "kpi_total_orders": "Total Orders",
        "kpi_avg_order": "Avg Order Value",
        "kpi_inventory_value": "Total Inventory Value",
        "tab_sales_trend": "📈 Sales Trend",
        "tab_top_products": "🏆 Top Products",
        "tab_inventory": "📦 Inventory",
        "tab_regional": "🗺️ Regional Sales",
        "sales_trend_title": "Sales Trend Over Time",
        "top_products_title": "Top 5 Products by Revenue",
        "inventory_title": "Inventory by Category",
        "regional_title": "Sales by Region",
        "download_report_btn": "📥 Download Report (CSV)",
        "download_click": "📥 Click to download CSV",
        "report_generated": "Report generated!",
        "no_data_msg": "No data available. Please upload a file or use sample data.",
        "load_sample_now": "Load Sample Data Now",
        "inventory_not_available": "Inventory data not available in this dataset.",
        "footer_note": "📊 **Custom reports available** – contact us to tailor dashboards to your exact needs.",
        "choose_data_source": "Choose data source:",
        "date_filter_hint": "Select start and end date"
    },
    "es": {
        "app_title": "Panel de Inteligencia de Negocios",
        "app_subtitle": "Analítica en tiempo real para empresas – conecta SQL, Excel, CSV",
        "sidebar_company": "GlobalInternet.py",
        "sidebar_product": "Panel BI",
        "founder": "Fundador y Desarrollador",
        "name": "Gesner Deslandes",
        "whatsapp": "WhatsApp",
        "email": "Correo",
        "website": "Sitio web",
        "price_label": "Precio",
        "price_value": "$1,200 USD (licencia única)",
        "copyright": "Todos los derechos reservados",
        "logout_btn": "Cerrar sesión",
        "data_source": "📂 Fuente de datos",
        "sample_data_option": "Datos de ejemplo de ventas",
        "csv_option": "Subir archivo CSV o Excel",
        "sqlite_option": "Subir base de datos SQLite (.db)",
        "load_sample_btn": "Cargar datos de ejemplo",
        "upload_csv_btn": "Subir CSV/Excel",
        "upload_sqlite_btn": "Subir SQLite",
        "date_filter": "Rango de fechas",
        "apply_filter": "Aplicar filtro",
        "kpi_total_sales": "Ventas totales",
        "kpi_total_orders": "Pedidos totales",
        "kpi_avg_order": "Valor promedio de pedido",
        "kpi_inventory_value": "Valor total del inventario",
        "tab_sales_trend": "📈 Tendencia de ventas",
        "tab_top_products": "🏆 Mejores productos",
        "tab_inventory": "📦 Inventario",
        "tab_regional": "🗺️ Ventas por región",
        "sales_trend_title": "Tendencia de ventas en el tiempo",
        "top_products_title": "Top 5 productos por ingresos",
        "inventory_title": "Inventario por categoría",
        "regional_title": "Ventas por región",
        "download_report_btn": "📥 Descargar informe (CSV)",
        "download_click": "📥 Haga clic para descargar CSV",
        "report_generated": "¡Informe generado!",
        "no_data_msg": "No hay datos disponibles. Suba un archivo o use datos de ejemplo.",
        "load_sample_now": "Cargar datos de ejemplo ahora",
        "inventory_not_available": "Datos de inventario no disponibles en este conjunto de datos.",
        "footer_note": "📊 **Informes personalizados disponibles** – contáctenos para adaptar los paneles a sus necesidades.",
        "choose_data_source": "Elija la fuente de datos:",
        "date_filter_hint": "Seleccione fecha de inicio y fin"
    },
    "fr": {
        "app_title": "Tableau de bord décisionnel",
        "app_subtitle": "Analytique en temps réel pour entreprises – connectez SQL, Excel, CSV",
        "sidebar_company": "GlobalInternet.py",
        "sidebar_product": "Tableau de bord BI",
        "founder": "Fondateur et développeur",
        "name": "Gesner Deslandes",
        "whatsapp": "WhatsApp",
        "email": "Email",
        "website": "Site web",
        "price_label": "Prix",
        "price_value": "1 200 $ USD (licence unique)",
        "copyright": "Tous droits réservés",
        "logout_btn": "Déconnexion",
        "data_source": "📂 Source de données",
        "sample_data_option": "Données de vente exemple",
        "csv_option": "Télécharger un fichier CSV ou Excel",
        "sqlite_option": "Télécharger une base SQLite (.db)",
        "load_sample_btn": "Charger les données exemple",
        "upload_csv_btn": "Télécharger CSV/Excel",
        "upload_sqlite_btn": "Télécharger SQLite",
        "date_filter": "Plage de dates",
        "apply_filter": "Appliquer le filtre",
        "kpi_total_sales": "Ventes totales",
        "kpi_total_orders": "Commandes totales",
        "kpi_avg_order": "Valeur moyenne de commande",
        "kpi_inventory_value": "Valeur totale de l'inventaire",
        "tab_sales_trend": "📈 Tendance des ventes",
        "tab_top_products": "🏆 Meilleurs produits",
        "tab_inventory": "📦 Inventaire",
        "tab_regional": "🗺️ Ventes par région",
        "sales_trend_title": "Tendance des ventes dans le temps",
        "top_products_title": "Top 5 produits par chiffre d'affaires",
        "inventory_title": "Inventaire par catégorie",
        "regional_title": "Ventes par région",
        "download_report_btn": "📥 Télécharger le rapport (CSV)",
        "download_click": "📥 Cliquez pour télécharger CSV",
        "report_generated": "Rapport généré !",
        "no_data_msg": "Aucune donnée disponible. Téléchargez un fichier ou utilisez les données d'exemple.",
        "load_sample_now": "Charger les données exemple maintenant",
        "inventory_not_available": "Données d'inventaire non disponibles dans cet ensemble.",
        "footer_note": "📊 **Rapports personnalisés disponibles** – contactez-nous pour adapter les tableaux de bord à vos besoins.",
        "choose_data_source": "Choisissez la source de données :",
        "date_filter_hint": "Sélectionnez la date de début et de fin"
    },
    "ht": {
        "app_title": "Tablodbò Entelijan Biznis",
        "app_subtitle": "Analiz an tan reyèl pou konpayi – konekte SQL, Excel, CSV",
        "sidebar_company": "GlobalInternet.py",
        "sidebar_product": "Tablodbò BI",
        "founder": "Fondatè ak Devlopè",
        "name": "Gesner Deslandes",
        "whatsapp": "WhatsApp",
        "email": "Imèl",
        "website": "Sitwèb",
        "price_label": "Pri",
        "price_value": "1,200 $ USD (peman inik)",
        "copyright": "Tout dwa rezève",
        "logout_btn": "Dekonekte",
        "data_source": "📂 Sous done",
        "sample_data_option": "Done lavant egzanp",
        "csv_option": "Telechaje fichye CSV oswa Excel",
        "sqlite_option": "Telechaje baz done SQLite (.db)",
        "load_sample_btn": "Chaje done egzanp",
        "upload_csv_btn": "Telechaje CSV/Excel",
        "upload_sqlite_btn": "Telechaje SQLite",
        "date_filter": "Ranje dat",
        "apply_filter": "Aplike filt",
        "kpi_total_sales": "Lavant total",
        "kpi_total_orders": "Kòmand total",
        "kpi_avg_order": "Valè mwayèn kòmand",
        "kpi_inventory_value": "Valè total envantè",
        "tab_sales_trend": "📈 Tandans lavant",
        "tab_top_products": "🏆 Pwodwi tèt yo",
        "tab_inventory": "📦 Envantè",
        "tab_regional": "🗺️ Lavant pa rejyon",
        "sales_trend_title": "Tandans lavant sou tan",
        "top_products_title": "Top 5 pwodwi pa revni",
        "inventory_title": "Envantè pa kategori",
        "regional_title": "Lavant pa rejyon",
        "download_report_btn": "📥 Telechaje rapò (CSV)",
        "download_click": "📥 Klike pou telechaje CSV",
        "report_generated": "Rapò kreye!",
        "no_data_msg": "Pa gen done disponib. Telechaje yon fichye oswa itilize done egzanp.",
        "load_sample_now": "Chaje done egzanp kounye a",
        "inventory_not_available": "Done envantè pa disponib nan seri done sa a.",
        "footer_note": "📊 **Rapò pèsonalize disponib** – kontakte nou pou adapte tablodbò yo nan bezwen ou yo.",
        "choose_data_source": "Chwazi sous done:",
        "date_filter_hint": "Chwazi dat kòmansman ak fini"
    }
}

def get_text(key):
    lang = st.session_state.get("language", "en")
    return TEXTS[lang].get(key, key)

# ------------------------------
# AFTER LOGIN – MAIN APP
# ------------------------------
col_flag, col_title = st.columns([1, 3])
with col_flag:
    show_haitian_flag(120)
with col_title:
    st.markdown(f"<h1>{get_text('app_title')}</h1>", unsafe_allow_html=True)
    st.markdown(f"*{get_text('app_subtitle')}*")

# ------------------------------
# SIDEBAR – INFO & LOGOUT & LANGUAGE
# ------------------------------
with st.sidebar:
    st.markdown(f"## 🇭🇹 {get_text('sidebar_company')}")
    show_haitian_flag(80)
    st.markdown(f"### {get_text('sidebar_product')}")
    st.markdown("---")
    st.markdown(f"**{get_text('founder')}:**")
    st.markdown(get_text('name'))
    st.markdown(f"📞 **{get_text('whatsapp')}:** [509 4738-5663](https://wa.me/50947385663)")
    st.markdown(f"📧 **{get_text('email')}:** deslandes78@gmail.com")
    st.markdown(f"🌐 **{get_text('website')}:** [www.globalinternet.py](https://www.globalinternet.py)")
    st.markdown("---")
    st.markdown(f"### {get_text('price_label')}")
    st.markdown(f"**{get_text('price_value')}**")
    st.markdown("---")
    st.markdown(f"### © 2025 GlobalInternet.py")
    st.markdown(get_text('copyright'))
    st.markdown("---")
    # Language selector
    lang_choice = st.selectbox("🌐 Language", list(LANGUAGES.keys()))
    st.session_state["language"] = LANGUAGES[lang_choice]
    st.markdown("---")
    if st.button(get_text('logout_btn'), use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

# ------------------------------
# DATA LOADING FUNCTIONS
# ------------------------------
@st.cache_data
def load_sample_data():
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
    if file.name.endswith('.csv'):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    if 'date' in df.columns:
        df['date'] = pd.to_datetime(df['date'])
    return df

def load_from_sqlite(file):
    conn = sqlite3.connect(file)
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
# INITIALIZE SESSION STATE
# ------------------------------
if "sales_df" not in st.session_state:
    st.session_state.sales_df = None
if "inventory_df" not in st.session_state:
    st.session_state.inventory_df = None
if "filtered_sales" not in st.session_state:
    st.session_state.filtered_sales = None

# ------------------------------
# DATA SOURCE SELECTION (with translations)
# ------------------------------
st.sidebar.markdown(f"## {get_text('data_source')}")
data_option = st.sidebar.radio(
    get_text('choose_data_source'),
    [get_text('sample_data_option'), get_text('csv_option'), get_text('sqlite_option')]
)

if data_option == get_text('sample_data_option'):
    if st.sidebar.button(get_text('load_sample_btn'), use_container_width=True):
        sales, inventory = load_sample_data()
        st.session_state.sales_df = sales
        st.session_state.inventory_df = inventory
        st.session_state.filtered_sales = sales.copy()
        st.rerun()

elif data_option == get_text('csv_option'):
    uploaded_file = st.sidebar.file_uploader(get_text('csv_option'), type=['csv', 'xlsx', 'xls'])
    if uploaded_file:
        df = load_from_csv(uploaded_file)
        st.session_state.sales_df = df
        st.session_state.inventory_df = None
        st.session_state.filtered_sales = df.copy()
        st.rerun()

elif data_option == get_text('sqlite_option'):
    uploaded_db = st.sidebar.file_uploader(get_text('sqlite_option'), type=['db'])
    if uploaded_db:
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
        col1, col2 = st.sidebar.columns(2)
        with col1:
            start_date = st.date_input(get_text('date_filter'), min_date, min_value=min_date, max_value=max_date)
        with col2:
            end_date = st.date_input("", max_date, min_value=min_date, max_value=max_date)  # second part
        if st.sidebar.button(get_text('apply_filter'), use_container_width=True):
            mask = (sales['date'] >= pd.Timestamp(start_date)) & (sales['date'] <= pd.Timestamp(end_date))
            st.session_state.filtered_sales = sales[mask].copy()
            st.rerun()
        filtered = st.session_state.filtered_sales if st.session_state.filtered_sales is not None else sales
    else:
        filtered = sales
    
    # KPIs
    col1, col2, col3, col4 = st.columns(4)
    total_revenue = filtered['revenue'].sum() if 'revenue' in filtered.columns else 0
    total_orders = len(filtered)
    avg_order = total_revenue / total_orders if total_orders > 0 else 0
    total_inventory = inventory['inventory_value'].sum() if inventory is not None else 0
    
    with col1:
        st.metric(get_text('kpi_total_sales'), f"${total_revenue:,.0f}")
    with col2:
        st.metric(get_text('kpi_total_orders'), f"{total_orders:,}")
    with col3:
        st.metric(get_text('kpi_avg_order'), f"${avg_order:,.2f}")
    with col4:
        st.metric(get_text('kpi_inventory_value'), f"${total_inventory:,.0f}")
    
    # Tabs with translated names
    tab1, tab2, tab3, tab4 = st.tabs([
        get_text('tab_sales_trend'),
        get_text('tab_top_products'),
        get_text('tab_inventory'),
        get_text('tab_regional')
    ])
    
    with tab1:
        if 'date' in filtered.columns:
            daily = filtered.groupby('date')['revenue'].sum().reset_index()
            fig = px.line(daily, x='date', y='revenue', title=get_text('sales_trend_title'))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(get_text('no_data_msg'))
    
    with tab2:
        if 'product' in filtered.columns:
            product_sales = filtered.groupby('product')['revenue'].sum().sort_values(ascending=False).head(5)
            fig = px.bar(x=product_sales.index, y=product_sales.values, title=get_text('top_products_title'))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(get_text('no_data_msg'))
    
    with tab3:
        if inventory is not None:
            inv_by_cat = inventory.groupby('category')['inventory_value'].sum().reset_index()
            fig = px.pie(inv_by_cat, values='inventory_value', names='category', title=get_text('inventory_title'))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(get_text('inventory_not_available'))
    
    with tab4:
        if 'region' in filtered.columns:
            region_sales = filtered.groupby('region')['revenue'].sum().reset_index()
            fig = px.bar(region_sales, x='region', y='revenue', title=get_text('regional_title'))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(get_text('no_data_msg'))
    
    # Download report
    st.markdown("---")
    if st.button(get_text('download_report_btn')):
        csv = filtered.to_csv(index=False)
        st.download_button(
            label=get_text('download_click'),
            data=csv,
            file_name=f"bi_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
else:
    st.info(get_text('no_data_msg'))
    if st.button(get_text('load_sample_now')):
        sales, inventory = load_sample_data()
        st.session_state.sales_df = sales
        st.session_state.inventory_df = inventory
        st.session_state.filtered_sales = sales.copy()
        st.rerun()

# Footer
st.markdown("---")
st.markdown(get_text('footer_note'))
