import os
import re
import json
import base64
import sqlite3
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz
from openai import OpenAI


# ======================================================
# CONFIG IBRIDA: STREAMLIT SECRETS + CONFIG.PY LOCALE
# ======================================================

def get_secret(key, default=None):
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass

    try:
        import config
        if hasattr(config, key):
            return getattr(config, key)
    except Exception:
        pass

    return default


OPENAI_API_KEY = get_secret("OPENAI_API_KEY")
DB_PATH = get_secret("DB_PATH", "procurement_intelligence.db")
MATCH_THRESHOLD = int(get_secret("MATCH_THRESHOLD", 78))
LOGIN_PASSWORD = get_secret("LOGIN_PASSWORD", "1234")
LOGO_PATH = get_secret("LOGO_PATH", "logo.png")


# ======================================================
# CONFIG BASE
# ======================================================

APP_NAME = "ALMOND INTELLIGENCE"
USER_NAME = "Amandorla"

os.makedirs("uploads", exist_ok=True)
os.makedirs("output", exist_ok=True)
os.makedirs("logs", exist_ok=True)

st.set_page_config(
    page_title=APP_NAME,
    layout="wide",
    initial_sidebar_state="expanded"
)

if not OPENAI_API_KEY:
    st.error("OPENAI_API_KEY non trovata. Inseriscila in config.py oppure nei secrets di Streamlit.")
    st.stop()


# ======================================================
# SFONDO
# ======================================================

def get_base64(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()


background_image = get_base64("sfondo.png")


# ======================================================
# CSS DARK CORPORATE
# ======================================================

st.markdown(f"""
<style>

.stApp {{
    background-image: url("data:image/png;base64,{background_image}");
    background-size: cover;
    background-position: center;
    background-attachment: fixed;
    color: white;
}}

.stApp::before {{
    content: "";
    position: fixed;
    inset: 0;
    background: rgba(2, 10, 20, 0.80);
    z-index: -1;
}}

[data-testid="stSidebar"] {{
    background-color: rgba(6, 18, 33, 0.95);
}}

h1, h2, h3, h4, h5, h6, p, label, span {{
    color: white !important;
}}

.stButton > button,
.stDownloadButton > button {{
    background-color: #b30000 !important;
    color: white !important;
    border: 1px solid #ff4d4d !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
}}

.stButton > button:hover,
.stDownloadButton > button:hover {{
    background-color: #d00000 !important;
    color: white !important;
}}

div[data-testid="stMetric"] {{
    background-color: rgba(8, 33, 61, 0.85);
    border: 1px solid rgba(0, 210, 255, 0.35);
    border-radius: 14px;
    padding: 18px;
}}

.card {{
    background: rgba(8, 33, 61, 0.82);
    border: 1px solid rgba(0, 210, 255, 0.25);
    border-radius: 18px;
    padding: 22px;
    margin-bottom: 18px;
}}

.hero {{
    background: linear-gradient(90deg, rgba(8,33,61,0.95), rgba(0,80,120,0.35));
    border: 1px solid rgba(0,210,255,0.35);
    border-radius: 22px;
    padding: 30px;
    margin-bottom: 25px;
}}

.small-muted {{
    color: #b8c7d9 !important;
    font-size: 14px;
}}

</style>
""", unsafe_allow_html=True)


# ======================================================
# LOGIN
# ======================================================

def login_screen():
    st.markdown("<br><br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1])

    with col2:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=180)

        st.markdown(f"""
        <div class="card">
            <h2>{APP_NAME}</h2>
            <p class="small-muted">Vendor Intelligence • Market Intelligence • AI Matching</p>
        </div>
        """, unsafe_allow_html=True)

        password = st.text_input("Password", type="password")

        if st.button("Accedi", use_container_width=True):
            if password == LOGIN_PASSWORD:
                st.session_state["logged_in"] = True
                st.rerun()
            else:
                st.error("Password errata")


if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    login_screen()
    st.stop()


# ======================================================
# DATABASE
# ======================================================

def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS suppliers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        package_name TEXT,
        supplier_name TEXT,
        vat_number TEXT,
        email TEXT,
        phone TEXT,
        address_nl1 TEXT,
        address_nl2 TEXT,
        website TEXT,
        source_file TEXT,
        created_at TEXT
    )
    """)

    conn.commit()
    conn.close()


def insert_supplier(row):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
    INSERT INTO suppliers (
        package_name,
        supplier_name,
        vat_number,
        email,
        phone,
        address_nl1,
        address_nl2,
        website,
        source_file,
        created_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        row.get("package_name", ""),
        row.get("supplier_name", ""),
        row.get("vat_number", ""),
        row.get("email", ""),
        row.get("phone", ""),
        row.get("address_nl1", ""),
        row.get("address_nl2", ""),
        row.get("website", ""),
        row.get("source_file", ""),
        datetime.now().isoformat(timespec="seconds")
    ))

    conn.commit()
    conn.close()


def load_suppliers():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM suppliers", conn)
    conn.close()
    return df


init_db()


# ======================================================
# LETTURA EXCEL INTELLIGENTE
# ======================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    value = str(value).replace("\n", " ").replace("\t", " ").strip()
    value = re.sub(r"\s+", " ", value)
    return value


def read_excel_raw(file):
    return pd.read_excel(file, dtype=str, header=None).fillna("")


def normalize_col(c):
    return str(c).strip().lower()


def find_col(df, keywords):
    for col in df.columns:
        col_norm = normalize_col(col)
        for kw in keywords:
            if kw in col_norm:
                return col
    return None


def detect_columns(df):
    return {
        "package_col": find_col(df, [
            "pacchetto", "lavorazione", "descrizione", "categoria",
            "scope", "risorsa merceologico", "merceologico"
        ]),
        "supplier_col": find_col(df, [
            "fornitore", "supplier", "vendor", "ragione sociale",
            "azienda", "impresa", "nome societa", "nome società",
            "societa", "società"
        ]),
        "vat_col": find_col(df, [
            "piva", "p.iva", "partita iva", "vat"
        ]),
        "email_col": find_col(df, [
            "email", "mail", "e-mail"
        ]),
        "phone_col": find_col(df, [
            "telefono", "phone", "tel", "cellulare"
        ]),
        "address1_col": find_col(df, [
            "nl1", "indirizzo", "address"
        ]),
        "address2_col": find_col(df, [
            "nl2"
        ]),
        "website_col": find_col(df, [
            "sito", "website", "url", "web"
        ])
    }


# ======================================================
# IMPORT VENDOR INTELLIGENTE
# ======================================================

def import_vendor_excel(file):

    raw_df = read_excel_raw(file)

    best_header_row = None
    best_score = -1

    keywords = [
        "merceologico", "societa", "società", "mail",
        "telefono", "fornitore", "azienda"
    ]

    for idx in range(min(len(raw_df), 40)):

        row_values = [clean_text(v).lower() for v in raw_df.iloc[idx].tolist()]
        joined = " ".join(row_values)

        score = 0

        for kw in keywords:
            if kw in joined:
                score += 1

        if score > best_score:
            best_score = score
            best_header_row = idx

    headers = [clean_text(v) for v in raw_df.iloc[best_header_row].tolist()]

    df = raw_df.iloc[best_header_row + 1:].copy()
    df.columns = headers
    df = df.fillna("")

    cols = detect_columns(df)

    count = 0

    for _, r in df.iterrows():

        package = clean_text(r.get(cols["package_col"], ""))
        supplier = clean_text(r.get(cols["supplier_col"], ""))

        if not package or not supplier:
            continue

        insert_supplier({
            "package_name": package,
            "supplier_name": supplier,
            "vat_number": clean_text(r.get(cols["vat_col"], "")) if cols["vat_col"] else "",
            "email": clean_text(r.get(cols["email_col"], "")) if cols["email_col"] else "",
            "phone": clean_text(r.get(cols["phone_col"], "")) if cols["phone_col"] else "",
            "address_nl1": clean_text(r.get(cols["address1_col"], "")) if cols["address1_col"] else "",
            "address_nl2": clean_text(r.get(cols["address2_col"], "")) if cols["address2_col"] else "",
            "website": clean_text(r.get(cols["website_col"], "")) if cols["website_col"] else "",
            "source_file": file.name
        })

        count += 1

    return count


# ======================================================
# MATCHING
# ======================================================

def match_package(package, memory):

    results = []

    for _, row in memory.iterrows():

        score = fuzz.token_set_ratio(
            str(package).lower(),
            str(row["package_name"]).lower()
        )

        if score >= MATCH_THRESHOLD:
            item = row.to_dict()
            item["score"] = score
            results.append(item)

    return sorted(results, key=lambda x: x["score"], reverse=True)


# ======================================================
# SIDEBAR CON TASTI
# ======================================================

with st.sidebar:

    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=160)

    st.markdown(f"### {APP_NAME}")
    st.caption("Procurement dashboard")

    st.divider()

    if "section" not in st.session_state:
        st.session_state["section"] = "Dashboard"

    if st.button("Dashboard", use_container_width=True):
        st.session_state["section"] = "Dashboard"

    if st.button("Import Vendor Storiche", use_container_width=True):
        st.session_state["section"] = "Import Vendor Storiche"

    if st.button("Genera Vendor Gara", use_container_width=True):
        st.session_state["section"] = "Genera Vendor Gara"

    if st.button("Scouting", use_container_width=True):
        st.session_state["section"] = "Scouting"

    if st.button("Database", use_container_width=True):
        st.session_state["section"] = "Database"

    section = st.session_state["section"]

    st.divider()

    st.markdown(f"Utente: **{USER_NAME}**")

    if st.button("Logout", use_container_width=True):
        st.session_state["logged_in"] = False
        st.rerun()


# ======================================================
# HEADER
# ======================================================

st.markdown(f"""
<div class="hero">
    <h1>{APP_NAME}</h1>
    <p class="small-muted">Vendor Intelligence • Market Intelligence • AI Matching</p>
</div>
""", unsafe_allow_html=True)


# ======================================================
# DASHBOARD
# ======================================================

memory = load_suppliers()

if section == "Dashboard":

    col1, col2 = st.columns(2)

    total_suppliers = memory["supplier_name"].nunique() if not memory.empty else 0

    if not memory.empty:
        memory["created_at_dt"] = pd.to_datetime(memory["created_at"], errors="coerce")

        last_30 = memory[
            memory["created_at_dt"] >= datetime.now() - timedelta(days=30)
        ]["supplier_name"].nunique()
    else:
        last_30 = 0

    with col1:
        st.metric("Fornitori totali nel database", total_suppliers)

    with col2:
        st.metric("Fornitori aggiunti ultimi 30 giorni", last_30)


# ======================================================
# IMPORT VENDOR
# ======================================================

elif section == "Import Vendor Storiche":

    st.header("Import Vendor Storiche")

    files = st.file_uploader(
        "Carica una o più vendor list Excel compilate",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if st.button("Importa nel database", use_container_width=True):

        if not files:
            st.warning("Carica almeno un file Excel.")

        else:

            total = 0

            for file in files:

                try:
                    count = import_vendor_excel(file)
                    total += count

                    st.success("Vendor importata correttamente")
                    st.info(f"Numero fornitori importati: {count}")
                    st.success("Database aggiornato")

                except Exception as e:
                    st.error(f"Errore su {file.name}: {e}")

            st.info(f"Totale fornitori importati: {total}")


# ======================================================
# GENERA VENDOR
# ======================================================

elif section == "Genera Vendor Gara":

    st.header("Genera Vendor Gara")

    st.info("Funzione in lavorazione")


# ======================================================
# SCOUTING
# ======================================================

elif section == "Scouting":

    st.header("Scouting")

    st.info("Versione iniziale scouting")


# ======================================================
# DATABASE
# ======================================================

elif section == "Database":

    st.header("Database storico vendor")

    st.metric("Righe totali", len(memory))

    if not memory.empty:

        st.dataframe(memory, use_container_width=True)

        csv = memory.to_csv(index=False).encode("utf-8")

        st.download_button(
            "Scarica database CSV",
            data=csv,
            file_name="database_vendor.csv",
            mime="text/csv",
            use_container_width=True
        )

    else:
        st.info("Database ancora vuoto.")
