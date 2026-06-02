import os
import sqlite3
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st
from rapidfuzz import fuzz

from config import OPENAI_API_KEY, DB_PATH, MATCH_THRESHOLD, LOGIN_PASSWORD, LOGO_PATH


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


# ======================================================
# CSS DARK CORPORATE
# ======================================================

st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #071526 0%, #0b2238 45%, #08111f 100%);
    color: white;
}

[data-testid="stSidebar"] {
    background-color: #061221;
}

h1, h2, h3, h4, h5, h6, p, label, span {
    color: white !important;
}

.stButton > button,
.stDownloadButton > button {
    background-color: #08213d !important;
    color: white !important;
    border: 1px solid #1e90ff !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background-color: #0d335f !important;
    color: white !important;
}

div[data-testid="stMetric"] {
    background-color: rgba(8, 33, 61, 0.85);
    border: 1px solid rgba(0, 210, 255, 0.35);
    border-radius: 14px;
    padding: 18px;
}

.card {
    background: rgba(8, 33, 61, 0.82);
    border: 1px solid rgba(0, 210, 255, 0.25);
    border-radius: 18px;
    padding: 22px;
    margin-bottom: 18px;
}

.hero {
    background: linear-gradient(90deg, rgba(8,33,61,0.95), rgba(0,80,120,0.35));
    border: 1px solid rgba(0,210,255,0.35);
    border-radius: 22px;
    padding: 30px;
    margin-bottom: 25px;
}

.small-muted {
    color: #b8c7d9 !important;
    font-size: 14px;
}
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
# FUNZIONI EXCEL
# ======================================================

def read_excel(file):
    return pd.read_excel(file, dtype=str).fillna("")


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
        "package_col": find_col(df, ["pacchetto", "lavorazione", "descrizione", "categoria", "scope"]),
        "supplier_col": find_col(df, ["fornitore", "supplier", "vendor", "ragione sociale", "azienda", "impresa"]),
        "vat_col": find_col(df, ["piva", "p.iva", "partita iva", "vat"]),
        "email_col": find_col(df, ["email", "mail", "e-mail"]),
        "phone_col": find_col(df, ["telefono", "phone", "tel"]),
        "address1_col": find_col(df, ["nl1", "indirizzo", "address"]),
        "address2_col": find_col(df, ["nl2"]),
        "website_col": find_col(df, ["sito", "website", "url", "web"])
    }


def import_vendor_excel(file):
    df = read_excel(file)
    cols = detect_columns(df)

    if not cols["package_col"]:
        raise Exception("Colonna pacchetto/lavorazione non trovata.")
    if not cols["supplier_col"]:
        raise Exception("Colonna fornitore non trovata.")

    count = 0

    for _, r in df.iterrows():
        package = str(r.get(cols["package_col"], "")).strip()
        supplier = str(r.get(cols["supplier_col"], "")).strip()

        if not package or not supplier:
            continue

        insert_supplier({
            "package_name": package,
            "supplier_name": supplier,
            "vat_number": str(r.get(cols["vat_col"], "")).strip() if cols["vat_col"] else "",
            "email": str(r.get(cols["email_col"], "")).strip() if cols["email_col"] else "",
            "phone": str(r.get(cols["phone_col"], "")).strip() if cols["phone_col"] else "",
            "address_nl1": str(r.get(cols["address1_col"], "")).strip() if cols["address1_col"] else "",
            "address_nl2": str(r.get(cols["address2_col"], "")).strip() if cols["address2_col"] else "",
            "website": str(r.get(cols["website_col"], "")).strip() if cols["website_col"] else "",
            "source_file": file.name
        })

        count += 1

    return count


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


def generate_vendor_from_template(template_file):
    template_df = read_excel(template_file)
    cols = detect_columns(template_df)

    if not cols["package_col"]:
        raise Exception("Nel template non trovo la colonna pacchetto/lavorazione.")

    memory = load_suppliers()

    if memory.empty:
        raise Exception("Database storico vuoto. Importa prima almeno una vendor.")

    output_rows = []
    preview = []

    for _, row in template_df.iterrows():
        package = str(row.get(cols["package_col"], "")).strip()

        if not package:
            continue

        matches = match_package(package, memory)

        seen = set()
        clean_matches = []

        for m in matches:
            key = str(m.get("vat_number") or m.get("supplier_name")).lower().strip()
            if key and key not in seen:
                seen.add(key)
                clean_matches.append(m)

        preview.append({
            "Pacchetto": package,
            "Fornitori trovati": len(clean_matches),
            "Miglior match": clean_matches[0]["package_name"] if clean_matches else "",
            "Score": clean_matches[0]["score"] if clean_matches else ""
        })

        if clean_matches:
            for supplier in clean_matches:
                new_row = row.copy()
                if cols["supplier_col"]:
                    new_row[cols["supplier_col"]] = supplier.get("supplier_name", "")
                if cols["vat_col"]:
                    new_row[cols["vat_col"]] = supplier.get("vat_number", "")
                if cols["email_col"]:
                    new_row[cols["email_col"]] = supplier.get("email", "")
                if cols["phone_col"]:
                    new_row[cols["phone_col"]] = supplier.get("phone", "")
                if cols["address1_col"]:
                    new_row[cols["address1_col"]] = supplier.get("address_nl1", "")
                if cols["address2_col"]:
                    new_row[cols["address2_col"]] = supplier.get("address_nl2", "")
                if cols["website_col"]:
                    new_row[cols["website_col"]] = supplier.get("website", "")
                output_rows.append(new_row)
        else:
            output_rows.append(row)

    output_df = pd.DataFrame(output_rows)
    preview_df = pd.DataFrame(preview)

    output_path = os.path.join(
        "output",
        f"vendor_compilata_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    output_df.to_excel(output_path, index=False)

    return output_path, preview_df


# ======================================================
# SIDEBAR
# ======================================================

with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=160)

    st.markdown(f"### {APP_NAME}")
    st.caption("Procurement dashboard")
    st.divider()

    section = st.radio(
        "Menu",
        [
            "Dashboard",
            "Import Vendor Storiche",
            "Genera Vendor Gara",
            "Scouting",
            "Database"
        ]
    )

    st.divider()
    st.markdown(f"Utente: **{USER_NAME}**")

    if st.button("Logout"):
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

    st.markdown("""
    <div class="card">
        <h3>Flusso operativo</h3>
        <p>1. Importa vendor storiche Excel</p>
        <p>2. Genera vendor list per nuova gara</p>
        <p>3. Usa scouting per cercare nuovi fornitori</p>
        <p>4. Alimenta progressivamente la memoria storica locale</p>
    </div>
    """, unsafe_allow_html=True)


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

    if st.button("Importa nel database"):
        if not files:
            st.warning("Carica almeno un file Excel.")
        else:
            total = 0
            for file in files:
                try:
                    count = import_vendor_excel(file)
                    total += count
                    st.success(f"{file.name}: importate {count} righe.")
                except Exception as e:
                    st.error(f"Errore su {file.name}: {e}")

            st.info(f"Totale righe importate: {total}")


# ======================================================
# GENERA VENDOR
# ======================================================

elif section == "Genera Vendor Gara":
    st.header("Genera Vendor Gara")

    template = st.file_uploader(
        "Carica template vendor list della gara",
        type=["xlsx"]
    )

    if st.button("Genera vendor compilata"):
        if not template:
            st.warning("Carica prima il template.")
        else:
            try:
                output_path, preview_df = generate_vendor_from_template(template)

                st.subheader("Anteprima matching")
                st.dataframe(preview_df, use_container_width=True)

                with open(output_path, "rb") as f:
                    st.download_button(
                        "Scarica Excel compilato",
                        data=f,
                        file_name=os.path.basename(output_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                st.success("Vendor list generata correttamente.")

            except Exception as e:
                st.error(str(e))


# ======================================================
# SCOUTING
# ======================================================

elif section == "Scouting":
    st.header("Scouting")

    st.info("Versione iniziale: lo scouting web automatico sarà ampliato dopo. Per ora usiamo il motore database + predisposizione OpenAI.")

    package = st.text_input("Inserisci lavorazione/pacchetto")

    if st.button("Cerca nel database storico"):
        if not package:
            st.warning("Inserisci una lavorazione.")
        else:
            results = match_package(package, memory)

            if results:
                st.success(f"Trovati {len(results)} possibili fornitori già presenti.")
                st.dataframe(pd.DataFrame(results), use_container_width=True)
            else:
                st.warning("Nessun fornitore trovato nel database storico.")


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
            mime="text/csv"
        )
    else:
        st.info("Database ancora vuoto.")
