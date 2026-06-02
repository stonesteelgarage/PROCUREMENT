import os
import re
import json
import sqlite3
import pandas as pd
import streamlit as st
from openai import OpenAI

# =========================
# CONFIG
# =========================

DB_PATH = "suppliers.db"

def get_secret(name, default=""):
    try:
        return st.secrets[name]
    except Exception:
        try:
            import config
            return getattr(config, name, default)
        except Exception:
            return default

OPENAI_API_KEY = get_secret("OPENAI_API_KEY")

client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None


# =========================
# UTILS
# =========================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def normalize(value):
    value = clean_text(value).lower()
    value = re.sub(r"[^a-z0-9àèéìòù]+", "", value)
    return value

def fallback_supplier_from_email_or_domain(email, sito=""):
    base = email or sito
    if not base:
        return ""
    m = re.search(r"@([A-Za-z0-9.-]+)", base)
    domain = m.group(1) if m else base
    domain = domain.replace("www.", "").split("/")[0]
    name = domain.split(".")[0]
    return name.upper()


# =========================
# DATABASE
# =========================

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS suppliers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_name TEXT,
            package_name TEXT,
            package_code TEXT,
            vat_number TEXT,
            email TEXT,
            phone TEXT,
            contact_person TEXT,
            address_1 TEXT,
            address_2 TEXT,
            website TEXT,
            source_file TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(supplier_name, package_name, email)
        )
    """)

    conn.commit()
    conn.close()


def save_supplier(row):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
        INSERT OR IGNORE INTO suppliers (
            supplier_name,
            package_name,
            package_code,
            vat_number,
            email,
            phone,
            contact_person,
            address_1,
            address_2,
            website,
            source_file
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        row.get("supplier_name", ""),
        row.get("package_name", ""),
        row.get("package_code", ""),
        row.get("vat_number", ""),
        row.get("email", ""),
        row.get("phone", ""),
        row.get("contact_person", ""),
        row.get("address_1", ""),
        row.get("address_2", ""),
        row.get("website", ""),
        row.get("source_file", "")
    ))

    conn.commit()
    inserted = cur.rowcount
    conn.close()
    return inserted


# =========================
# EXCEL READING
# =========================

def detect_header_row(raw_df):
    keywords = [
        "fornitore", "societa", "società", "nome societa", "nome società",
        "mail", "email", "telefono", "referente", "risorsa", "merceologico",
        "piva", "partita iva", "indirizzo"
    ]

    best_row = 0
    best_score = 0

    for i in range(min(30, len(raw_df))):
        row_text = " ".join([clean_text(x).lower() for x in raw_df.iloc[i].tolist()])
        score = sum(1 for k in keywords if k in row_text)
        if score > best_score:
            best_score = score
            best_row = i

    return best_row


def ask_openai_column_mapping(columns, sample_rows):
    if not client:
        return {}

    prompt = f"""
Sei un assistente esperto in vendor list procurement Excel.

Devi mappare le colonne reali del file Excel verso questi campi standard:

- supplier_name
- package_name
- package_code
- vat_number
- email
- phone
- contact_person
- address_1
- address_2
- website

Colonne disponibili:
{columns}

Esempio righe:
{sample_rows}

Rispondi SOLO con JSON valido.
Le chiavi devono essere i campi standard.
I valori devono essere il nome esatto della colonna Excel.
Se una colonna non esiste, usa stringa vuota.
"""

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": "Rispondi solo con JSON valido, senza markdown."},
            {"role": "user", "content": prompt}
        ],
        temperature=0
    )

    text = response.choices[0].message.content.strip()
    return json.loads(text)


def fallback_column_mapping(columns):
    mapping = {}

    for col in columns:
        c = str(col).lower().strip()

        if "nome societ" in c or "fornitore" in c or "ragione sociale" in c:
            mapping["supplier_name"] = col
        elif "risorsa" in c or "merceologico" in c or "lavorazione" in c or "pacchetto" in c:
            mapping["package_name"] = col
        elif "codice" in c and ("pacchetto" in c or "risorsa" in c or "wbs" in c):
            mapping["package_code"] = col
        elif "iva" in c or "p.iva" in c or "partita" in c:
            mapping["vat_number"] = col
        elif "mail" in c or "email" in c:
            mapping["email"] = col
        elif "telefono" in c or "tel" in c or "cell" in c:
            mapping["phone"] = col
        elif "referente" in c or "contatto" in c:
            mapping["contact_person"] = col
        elif "indirizzo" in c and "2" not in c:
            mapping["address_1"] = col
        elif "indirizzo" in c and "2" in c:
            mapping["address_2"] = col
        elif "sito" in c or "website" in c or "web" in c:
            mapping["website"] = col

    return mapping


def read_vendor_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    all_rows = []

    for sheet_name in xls.sheet_names:
        raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

        if raw.dropna(how="all").empty:
            continue

        header_row = detect_header_row(raw)

        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name,
            header=header_row
        )

        df = df.dropna(how="all")
        df.columns = [clean_text(c) for c in df.columns]

        columns = list(df.columns)
        sample_rows = df.head(10).fillna("").to_dict(orient="records")

        try:
            mapping = ask_openai_column_mapping(columns, sample_rows)
        except Exception:
            mapping = fallback_column_mapping(columns)

        if not mapping:
            mapping = fallback_column_mapping(columns)

        for _, r in df.iterrows():
            supplier_name = clean_text(r.get(mapping.get("supplier_name", ""), ""))
            email = clean_text(r.get(mapping.get("email", ""), ""))
            website = clean_text(r.get(mapping.get("website", ""), ""))

            if not supplier_name:
                supplier_name = fallback_supplier_from_email_or_domain(email, website)

            package_name = clean_text(r.get(mapping.get("package_name", ""), ""))

            if not supplier_name or not package_name:
                continue

            row = {
                "supplier_name": supplier_name,
                "package_name": package_name,
                "package_code": clean_text(r.get(mapping.get("package_code", ""), "")),
                "vat_number": clean_text(r.get(mapping.get("vat_number", ""), "")),
                "email": email,
                "phone": clean_text(r.get(mapping.get("phone", ""), "")),
                "contact_person": clean_text(r.get(mapping.get("contact_person", ""), "")),
                "address_1": clean_text(r.get(mapping.get("address_1", ""), "")),
                "address_2": clean_text(r.get(mapping.get("address_2", ""), "")),
                "website": website,
                "source_file": uploaded_file.name
            }

            all_rows.append(row)

    return all_rows


# =========================
# STREAMLIT UI
# =========================

st.set_page_config(
    page_title="Almond Intelligence",
    layout="wide"
)

st.title("ALMOND INTELLIGENCE")
st.subheader("Vendor Intelligence • Market Intelligence • AI Matching")

init_db()

uploaded_file = st.file_uploader(
    "Carica vendor list Excel",
    type=["xlsx", "xls"]
)

if uploaded_file:
    if st.button("Aggiorna database fornitori"):
        with st.spinner("Lettura vendor list con OpenAI e aggiornamento DB..."):
            rows = read_vendor_excel(uploaded_file)

            inserted = 0
            skipped = 0

            for row in rows:
                result = save_supplier(row)
                if result == 1:
                    inserted += 1
                else:
                    skipped += 1

        st.success("Database aggiornato.")
        st.write(f"Fornitori letti dal file: {len(rows)}")
        st.write(f"Nuovi fornitori inseriti: {inserted}")
        st.write(f"Fornitori già presenti / duplicati: {skipped}")

        if rows:
            st.dataframe(pd.DataFrame(rows), use_container_width=True)

# =========================
# DB VIEW
# =========================

st.divider()
st.subheader("Database fornitori")

conn = sqlite3.connect(DB_PATH)
df_db = pd.read_sql_query("""
    SELECT 
        supplier_name AS 'Fornitore',
        package_name AS 'Pacchetto / Lavorazione',
        package_code AS 'Codice',
        vat_number AS 'P.IVA',
        email AS 'Email',
        phone AS 'Telefono',
        contact_person AS 'Referente',
        address_1 AS 'Indirizzo 1',
        address_2 AS 'Indirizzo 2',
        website AS 'Sito',
        source_file AS 'File origine',
        created_at AS 'Creato il'
    FROM suppliers
    ORDER BY created_at DESC
""", conn)
conn.close()

st.dataframe(df_db, use_container_width=True)

st.download_button(
    "Scarica database in Excel",
    data=df_db.to_csv(index=False).encode("utf-8"),
    file_name="database_fornitori.csv",
    mime="text/csv"
)
