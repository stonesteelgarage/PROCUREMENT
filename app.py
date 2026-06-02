import os
import re
import json
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
    background-color: #b30000 !important;
    color: white !important;
    border: 1px solid #ff4d4d !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background-color: #d00000 !important;
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


def detect_header_row_and_columns_local(raw_df):
    best = None

    header_keywords = [
        "pacchetto", "lavorazione", "risorsa merceologico", "merceologico",
        "fornitore", "nome societa", "nome società", "societa", "società",
        "telefono", "mail", "email", "partita iva", "piva"
    ]

    for idx in range(min(len(raw_df), 40)):
        row_values = [clean_text(v).lower() for v in raw_df.iloc[idx].tolist()]
        joined = " | ".join(row_values)

        score = 0

        for kw in header_keywords:
            if kw in joined:
                score += 1

        non_empty = sum(1 for v in row_values if v)
        if non_empty >= 3:
            score += 1

        if best is None or score > best["score"]:
            best = {
                "row_index": idx,
                "score": score,
                "values": row_values
            }

    if not best or best["score"] < 2:
        return None

    header_idx = best["row_index"]
    headers = [clean_text(v) for v in raw_df.iloc[header_idx].tolist()]

    df = raw_df.iloc[header_idx + 1:].copy()
    df.columns = headers
    df = df.loc[:, [c for c in df.columns if str(c).strip() != ""]]
    df = df.fillna("")

    cols = detect_columns(df)

    return {
        "header_row_index": header_idx,
        "df": df,
        "cols": cols
    }


def ask_openai_excel_mapping(raw_df):
    client = OpenAI(api_key=OPENAI_API_KEY)

    preview_rows = []
    max_rows = min(len(raw_df), 25)
    max_cols = min(raw_df.shape[1], 25)

    for i in range(max_rows):
        row = []
        for j in range(max_cols):
            val = clean_text(raw_df.iat[i, j])
            if val:
                row.append({
                    "col_index": j,
                    "value": val[:120]
                })

        if row:
            preview_rows.append({
                "row_index": i,
                "cells": row
            })

    prompt = f"""
Sei un esperto di vendor list procurement in Excel.

Devi analizzare una preview di un file Excel senza intestazioni standard.
Devi capire:
1. qual è la riga delle intestazioni;
2. quale colonna contiene la lavorazione/pacchetto;
3. quale colonna contiene il nome fornitore/società;
4. se presenti, telefono, email, partita IVA, indirizzi, sito.

La numerazione di righe e colonne parte da 0.

Preview:
{json.dumps(preview_rows, ensure_ascii=False)}

Rispondi SOLO con JSON valido, senza markdown, con questa struttura:
{{
  "header_row_index": 0,
  "package_col_index": null,
  "supplier_col_index": null,
  "vat_col_index": null,
  "email_col_index": null,
  "phone_col_index": null,
  "address1_col_index": null,
  "address2_col_index": null,
  "website_col_index": null,
  "notes": ""
}}
"""

    response = client.responses.create(
        model="gpt-4.1-mini",
        input=prompt
    )

    txt = response.output_text.strip()
    txt = txt.replace("```json", "").replace("```", "").strip()

    return json.loads(txt)


def dataframe_from_ai_mapping(raw_df, mapping):
    header_idx = mapping.get("header_row_index")

    if header_idx is None:
        raise Exception("OpenAI non ha individuato la riga intestazioni.")

    headers = [clean_text(v) for v in raw_df.iloc[header_idx].tolist()]

    df = raw_df.iloc[header_idx + 1:].copy()
    df.columns = headers
    df = df.fillna("")
    df = df.loc[:, [c for c in df.columns if str(c).strip() != ""]]

    def col_name(index):
        if index is None:
            return None

        try:
            index = int(index)
            if index < len(headers):
                return headers[index]
        except Exception:
            return None

        return None

    cols = {
        "package_col": col_name(mapping.get("package_col_index")),
        "supplier_col": col_name(mapping.get("supplier_col_index")),
        "vat_col": col_name(mapping.get("vat_col_index")),
        "email_col": col_name(mapping.get("email_col_index")),
        "phone_col": col_name(mapping.get("phone_col_index")),
        "address1_col": col_name(mapping.get("address1_col_index")),
        "address2_col": col_name(mapping.get("address2_col_index")),
        "website_col": col_name(mapping.get("website_col_index")),
    }

    return df, cols


def smart_read_vendor_excel(file):
    raw_df = read_excel_raw(file)

    local = detect_header_row_and_columns_local(raw_df)

    if local:
        df = local["df"]
        cols = local["cols"]

        if cols.get("package_col") and cols.get("supplier_col"):
            return df, cols, f"Riconoscimento automatico locale: intestazioni alla riga Excel {local['header_row_index'] + 1}"

    mapping = ask_openai_excel_mapping(raw_df)
    df, cols = dataframe_from_ai_mapping(raw_df, mapping)

    if not cols.get("package_col") or not cols.get("supplier_col"):
        raise Exception(f"OpenAI non ha individuato pacchetto e fornitore. Dettagli: {mapping}")

    return df, cols, f"Riconoscimento OpenAI: intestazioni alla riga Excel {mapping.get('header_row_index') + 1}"


def smart_read_template_excel(file):
    raw_df = read_excel_raw(file)

    local = detect_header_row_and_columns_local(raw_df)

    if local:
        df = local["df"]
        cols = local["cols"]

        if cols.get("package_col"):
            return df, cols, f"Riconoscimento automatico locale: intestazioni alla riga Excel {local['header_row_index'] + 1}"

    mapping = ask_openai_excel_mapping(raw_df)
    df, cols = dataframe_from_ai_mapping(raw_df, mapping)

    if not cols.get("package_col"):
        raise Exception(f"OpenAI non ha individuato la colonna pacchetto. Dettagli: {mapping}")

    return df, cols, f"Riconoscimento OpenAI: intestazioni alla riga Excel {mapping.get('header_row_index') + 1}"


def extract_email_from_text(text):
    text = clean_text(text)
    found = re.findall(
        r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
        text
    )
    return "; ".join(dict.fromkeys(found))


def extract_phone_from_text(text):
    text = clean_text(text)
    found = re.findall(
        r"(?:\+39\s?)?(?:0\d{1,4}[\s./-]?\d{5,8}|3\d{2}[\s./-]?\d{6,7})",
        text
    )
    return "; ".join(dict.fromkeys(found))


def import_vendor_excel(file):
    df, cols, detection_msg = smart_read_vendor_excel(file)

    count = 0

    for _, r in df.iterrows():
        package = clean_text(r.get(cols["package_col"], ""))
        supplier = clean_text(r.get(cols["supplier_col"], ""))

        if not package or not supplier:
            continue

        if supplier.lower() in [
            "nome societa'", "nome società", "nome societa",
            "fornitore", "supplier", "vendor", "azienda", "impresa"
        ]:
            continue

        email = clean_text(r.get(cols["email_col"], "")) if cols.get("email_col") else ""
        phone = clean_text(r.get(cols["phone_col"], "")) if cols.get("phone_col") else ""

        row_text = " ".join([clean_text(v) for v in r.tolist()])

        if not email:
            email = extract_email_from_text(row_text)

        if not phone:
            phone = extract_phone_from_text(row_text)

        insert_supplier({
            "package_name": package,
            "supplier_name": supplier,
            "vat_number": clean_text(r.get(cols["vat_col"], "")) if cols.get("vat_col") else "",
            "email": email,
            "phone": phone,
            "address_nl1": clean_text(r.get(cols["address1_col"], "")) if cols.get("address1_col") else "",
            "address_nl2": clean_text(r.get(cols["address2_col"], "")) if cols.get("address2_col") else "",
            "website": clean_text(r.get(cols["website_col"], "")) if cols.get("website_col") else "",
            "source_file": file.name
        })

        count += 1

    return count, detection_msg


# ======================================================
# MATCHING E GENERAZIONE VENDOR
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


def generate_vendor_from_template(template_file):
    template_df, cols, detection_msg = smart_read_template_excel(template_file)

    if not cols["package_col"]:
        raise Exception("Nel template non trovo la colonna pacchetto/lavorazione.")

    memory = load_suppliers()

    if memory.empty:
        raise Exception("Database storico vuoto. Importa prima almeno una vendor.")

    output_rows = []
    preview = []

    for _, row in template_df.iterrows():
        package = clean_text(row.get(cols["package_col"], ""))

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

    return output_path, preview_df, detection_msg


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

    if st.button("Importa nel database", use_container_width=True):
        if not files:
            st.warning("Carica almeno un file Excel.")
        else:
            total = 0

            for file in files:
                try:
                    count, msg = import_vendor_excel(file)
                    total += count

                    st.success("Vendor importata correttamente")
                    st.info(f"Numero fornitori importati: {count}")
                    st.success("Database aggiornato")
                    st.caption(msg)

                except Exception as e:
                    st.error(f"Errore su {file.name}: {e}")

            st.info(f"Totale fornitori importati: {total}")


# ======================================================
# GENERA VENDOR
# ======================================================

elif section == "Genera Vendor Gara":
    st.header("Genera Vendor Gara")

    template = st.file_uploader(
        "Carica template vendor list della gara",
        type=["xlsx"]
    )

    if st.button("Genera vendor compilata", use_container_width=True):
        if not template:
            st.warning("Carica prima il template.")
        else:
            try:
                output_path, preview_df, msg = generate_vendor_from_template(template)

                st.caption(msg)

                st.subheader("Anteprima matching")
                st.dataframe(preview_df, use_container_width=True)

                with open(output_path, "rb") as f:
                    st.download_button(
                        "Scarica Excel compilato",
                        data=f,
                        file_name=os.path.basename(output_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
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

    if st.button("Cerca nel database storico", use_container_width=True):
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
            mime="text/csv",
            use_container_width=True
        )
    else:
        st.info("Database ancora vuoto.")
