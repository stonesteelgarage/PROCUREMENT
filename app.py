# =====================================================
# PROCUREMENT INTELLIGENCE
# Almond Intelligence
# =====================================================

import streamlit as st
import pandas as pd
import sqlite3
import os
import re
import json
import requests

from io import BytesIO
from datetime import datetime, timedelta
from rapidfuzz import fuzz
from openai import OpenAI
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from duckduckgo_search import DDGS
from openpyxl.styles import Font, PatternFill, Alignment


# =====================================================
# CONFIG
# =====================================================

APP_TITLE = "PROCUREMENT INTELLIGENCE"
APP_SUBTITLE = "An AI-assisted Vendor Intelligence Platform"
BRAND_NAME = "Almond Intelligence"

DB_NAME = "procurement_intelligence.db"
LOGO_PATH = "logo.png"


def get_secret(key, default=None):
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass

    try:
        import config
        return getattr(config, key, default)
    except Exception:
        return default


OPENAI_API_KEY = get_secret("OPENAI_API_KEY", "")
LOGIN_PASSWORD = get_secret("LOGIN_PASSWORD", "admin")


# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="Procurement Intelligence",
    page_icon="📊",
    layout="wide"
)


# =====================================================
# STYLE
# =====================================================

st.markdown("""
<style>

.stApp {
    background: linear-gradient(135deg, #031B36 0%, #062A4F 45%, #021020 100%);
    color: white;
}

[data-testid="stSidebar"] {
    background-color: #02172E;
}

h1, h2, h3, h4, h5, h6, p, label, span {
    color: white;
}

.main-title {
    font-size: 48px;
    font-weight: 800;
    letter-spacing: 2px;
    color: white;
    margin-bottom: 0px;
}

.subtitle {
    font-size: 20px;
    color: #A9D6FF;
    margin-top: 4px;
    margin-bottom: 6px;
}

.brand {
    font-size: 14px;
    color: #68C7FF;
    letter-spacing: 3px;
    text-transform: uppercase;
    margin-bottom: 20px;
}

.card {
    background: rgba(255,255,255,0.06);
    border: 1px solid rgba(255,255,255,0.12);
    border-radius: 18px;
    padding: 22px;
    margin-bottom: 18px;
}

.kpi-card {
    background: rgba(255,255,255,0.08);
    border: 1px solid rgba(104,199,255,0.30);
    border-radius: 18px;
    padding: 24px;
    text-align: center;
}

.kpi-number {
    font-size: 40px;
    font-weight: 800;
    color: #68C7FF;
}

.kpi-label {
    color: #D7ECFF;
    font-size: 14px;
}

.stButton > button,
.stDownloadButton > button,
div[data-testid="stDownloadButton"] > button,
button[kind="primary"],
button[kind="secondary"] {
    background-color: #02172E !important;
    color: #FFFFFF !important;
    border-radius: 10px !important;
    border: 1px solid #68C7FF !important;
    font-weight: 700 !important;
}

.stButton > button *,
.stDownloadButton > button *,
div[data-testid="stDownloadButton"] > button *,
button[kind="primary"] *,
button[kind="secondary"] * {
    color: #FFFFFF !important;
}

.stButton > button:hover,
.stDownloadButton > button:hover,
div[data-testid="stDownloadButton"] > button:hover,
button[kind="primary"]:hover,
button[kind="secondary"]:hover {
    background-color: #063B70 !important;
    color: #FFFFFF !important;
}

[data-testid="stFileUploader"] button {
    background-color: #02172E !important;
    color: white !important;
    border-radius: 10px !important;
    border: 1px solid #68C7FF !important;
    font-weight: 700 !important;
}

[data-testid="stFileUploader"] button:hover {
    background-color: #063B70 !important;
    color: white !important;
}

</style>
""", unsafe_allow_html=True)


# =====================================================
# DATABASE
# =====================================================

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS suppliers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        package TEXT,
        supplier_name TEXT,
        vat_number TEXT,
        email TEXT,
        phone TEXT,
        address_nl1 TEXT,
        address_nl2 TEXT,
        source_file TEXT,
        created_at TEXT
    )
    """)

    conn.commit()
    conn.close()


init_db()


def insert_supplier(package, supplier_name, vat_number="", email="", phone="", address_nl1="", address_nl2="", source_file=""):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    c.execute("""
    INSERT INTO suppliers (
        package,
        supplier_name,
        vat_number,
        email,
        phone,
        address_nl1,
        address_nl2,
        source_file,
        created_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        package,
        supplier_name,
        vat_number,
        email,
        phone,
        address_nl1,
        address_nl2,
        source_file,
        datetime.now().isoformat()
    ))

    conn.commit()
    conn.close()


def load_suppliers():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT * FROM suppliers", conn)
    conn.close()
    return df


# =====================================================
# LOGIN
# =====================================================

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False


def login_screen():
    st.markdown("<br><br>", unsafe_allow_html=True)

    st.markdown(f"""
    <div class="card" style="max-width:500px;margin:auto;text-align:center;">
        <div class="main-title">{APP_TITLE}</div>
        <div class="subtitle">{APP_SUBTITLE}</div>
        <div class="brand">{BRAND_NAME}</div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns([1.4, 1, 1.4])

    with c2:
        password = st.text_input("Password", type="password")

        if st.button("ACCEDI"):
            if password == LOGIN_PASSWORD:
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("Password errata")


if not st.session_state.logged_in:
    login_screen()
    st.stop()


# =====================================================
# HEADER
# =====================================================

col1, col2 = st.columns([6, 1])

with col1:
    st.markdown(f"""
    <div class="main-title">{APP_TITLE}</div>
    <div class="subtitle">{APP_SUBTITLE}</div>
    <div class="brand">{BRAND_NAME}</div>
    """, unsafe_allow_html=True)

with col2:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=90)


# =====================================================
# SIDEBAR
# =====================================================

st.sidebar.markdown("## Almond Intelligence")
st.sidebar.markdown("Utente: **Amandorla**")
st.sidebar.markdown("---")

if "section" not in st.session_state:
    st.session_state.section = "Dashboard"

if "last_scouting_df" not in st.session_state:
    st.session_state.last_scouting_df = None

if "last_scouting_package" not in st.session_state:
    st.session_state.last_scouting_package = ""


if st.sidebar.button("Dashboard"):
    st.session_state.section = "Dashboard"

if st.sidebar.button("Importa vendor storiche"):
    st.session_state.section = "Importa"

if st.sidebar.button("Genera vendor list"):
    st.session_state.section = "Vendor"

if st.sidebar.button("Scouting fornitori"):
    st.session_state.section = "Scouting"

if st.sidebar.button("Database fornitori"):
    st.session_state.section = "Database"

section = st.session_state.section


# =====================================================
# HELPERS
# =====================================================

def dataframe_to_excel(df, sheet_name="Vendor List"):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        ws = writer.book[sheet_name]

        header_fill = PatternFill("solid", fgColor="02172E")
        header_font = Font(color="FFFFFF", bold=True)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        widths = {
            "A": 28,
            "B": 38,
            "C": 18,
            "D": 20,
            "E": 35,
            "F": 42,
            "G": 50,
            "H": 14
        }

        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    output.seek(0)
    return output


def guess_column(df, names):
    for col in df.columns:
        for n in names:
            if n.lower() in str(col).lower():
                return col
    return None


def supplier_name_from_domain(url):
    domain = urlparse(url).netloc.replace("www.", "")
    name = domain.split(".")[0]
    name = re.sub(r"[^A-Za-z0-9]", " ", name)
    return name.upper().strip()


# =====================================================
# IMPORT VENDOR LIST
# =====================================================

def normalize_vendor_dataframe(df, source_file=""):
    package_col = guess_column(df, ["package", "pacchetto", "categoria", "lavorazione"])
    supplier_col = guess_column(df, ["fornitore", "supplier", "vendor", "ragione sociale", "azienda"])
    vat_col = guess_column(df, ["iva", "vat", "piva", "partita iva"])
    email_col = guess_column(df, ["email", "mail"])
    phone_col = guess_column(df, ["telefono", "phone", "tel"])

    if not package_col or not supplier_col:
        return 0

    inserted = 0

    for _, row in df.iterrows():
        package = str(row.get(package_col, "")).strip()
        supplier = str(row.get(supplier_col, "")).strip()

        if package and supplier:
            insert_supplier(
                package=package,
                supplier_name=supplier,
                vat_number=str(row.get(vat_col, "")) if vat_col else "",
                email=str(row.get(email_col, "")) if email_col else "",
                phone=str(row.get(phone_col, "")) if phone_col else "",
                source_file=source_file
            )
            inserted += 1

    return inserted


# =====================================================
# MATCHING
# =====================================================

def find_matching_suppliers(package_name, threshold=70):
    suppliers_df = load_suppliers()

    if suppliers_df.empty:
        return pd.DataFrame()

    matches = []

    for _, row in suppliers_df.iterrows():
        score = fuzz.token_set_ratio(str(package_name), str(row["package"]))

        if score >= threshold:
            data = row.to_dict()
            data["matching"] = score
            matches.append(data)

    if not matches:
        return pd.DataFrame()

    result = pd.DataFrame(matches)
    result = result.sort_values(by="matching", ascending=False)

    return result


def generate_completed_vendor(template_df, threshold=70):
    package_col = guess_column(template_df, ["package", "pacchetto", "categoria", "lavorazione"])

    if not package_col:
        raise Exception("Colonna pacchetto non trovata.")

    final_rows = []

    for _, row in template_df.iterrows():
        package_name = str(row[package_col]).strip()
        matches = find_matching_suppliers(package_name, threshold)

        if matches.empty:
            final_rows.append(row.to_dict())
            continue

        for _, match in matches.iterrows():
            new_row = row.to_dict()

            for col in template_df.columns:
                col_lower = str(col).lower()

                if "fornitore" in col_lower or "supplier" in col_lower or "vendor" in col_lower:
                    new_row[col] = match["supplier_name"]

                if "iva" in col_lower or "vat" in col_lower or "piva" in col_lower:
                    new_row[col] = match["vat_number"]

                if "email" in col_lower or "mail" in col_lower:
                    new_row[col] = match["email"]

                if "telefono" in col_lower or "phone" in col_lower or "tel" in col_lower:
                    new_row[col] = match["phone"]

            final_rows.append(new_row)

    return pd.DataFrame(final_rows)


# =====================================================
# WEB SCOUTING
# =====================================================

BAD_DOMAINS = [
    "roblox", "wikipedia", "facebook", "linkedin", "instagram", "youtube",
    "amazon", "ebay", "subito", "reddit", "blog", "news", "pinterest",
    "tripadvisor"
]


def is_bad_url(url):
    u = url.lower()
    return any(bad in u for bad in BAD_DOMAINS)


def clean_text(text):
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def extract_contacts(text):
    emails = list(set(re.findall(
        r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
        text
    )))

    phones = list(set(re.findall(
        r"(?:(?:\+39|0039)\s*)?(?:0\d{1,4}[\s.-]?\d{5,8}|3\d{2}[\s.-]?\d{6,7})",
        text
    )))

    vat_numbers = list(set(re.findall(
        r"(?:P\.?\s?IVA|Partita IVA|PIVA|VAT)\s*[:\-]?\s*(\d{11})",
        text,
        flags=re.IGNORECASE
    )))

    return emails, phones, vat_numbers


def build_search_queries(query, area="Italia"):
    return [
        f"{query} azienda fornitore produttore catalogo contatti telefono email {area}",
        f"{query} produttori Italia fornitori catalogo azienda sede contatti",
        f"{query} aziende B2B vendita produzione servizi contatti {area}",
        f"{query} catalogo prodotti azienda industria fornitore {area}",
        f"{query} forniture lavori servizi impresa contatti Italia"
    ]


def discover_supplier_sites(query, area="Italia"):
    search_queries = build_search_queries(query, area)

    results = []
    seen_domains = set()

    with DDGS() as ddgs:
        for search_query in search_queries:
            for r in ddgs.text(search_query, max_results=40):
                url = r.get("href") or r.get("url")

                if not url:
                    continue

                if is_bad_url(url):
                    continue

                domain = urlparse(url).netloc.replace("www.", "")

                if domain in seen_domains:
                    continue

                seen_domains.add(domain)

                results.append({
                    "title": r.get("title", ""),
                    "url": url,
                    "domain": domain,
                    "query_used": search_query
                })

                if len(results) >= 80:
                    return results

    return results


def fetch_page(url):
    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    try:
        r = requests.get(url, headers=headers, timeout=8)

        if r.status_code != 200:
            return "", None

        soup = BeautifulSoup(r.text, "lxml")
        text = soup.get_text(" ")

        return clean_text(text), soup

    except Exception:
        return "", None


def get_internal_links(base_url, soup):
    links = []

    keywords = [
        "contatti", "contact", "azienda", "about", "chi-siamo",
        "servizi", "prodotti", "services", "products", "company",
        "catalogo"
    ]

    base_domain = urlparse(base_url).netloc

    for a in soup.find_all("a", href=True):
        href = urljoin(base_url, a["href"])
        parsed = urlparse(href)

        if parsed.netloc != base_domain:
            continue

        href_lower = href.lower()

        if any(k in href_lower for k in keywords):
            links.append(href)

    return list(set(links))[:8]


def crawl_supplier_site(url):
    all_text = ""

    home_text, soup = fetch_page(url)

    if home_text:
        all_text += home_text + " "

    if soup:
        links = get_internal_links(url, soup)

        for link in links:
            page_text, _ = fetch_page(link)

            if page_text:
                all_text += page_text + " "

    all_text = all_text[:20000]

    emails, phones, vat_numbers = extract_contacts(all_text)

    return {
        "url": url,
        "text": all_text,
        "emails": emails,
        "phones": phones,
        "vat_numbers": vat_numbers
    }


def evaluate_supplier_with_ai(category, crawled):
    if not OPENAI_API_KEY:
        return {}

    client = OpenAI(api_key=OPENAI_API_KEY)

    prompt = f"""
Categoria/lavorazione:
{category}

Testo estratto dal sito:
{crawled["text"][:15000]}

Devi capire se questo sito appartiene a un fornitore potenzialmente utile per procurement.

Restituisci SOLO JSON valido:

{{
"is_supplier": true,
"supplier_name":"",
"physical_address":"",
"products":"",
"matching_percent":0
}}

Regole:
- is_supplier true se il sito sembra di una azienda reale o catalogo aziendale.
- matching_percent da 0 a 100.
- se è poco pertinente ma potrebbe essere utile, dai comunque un matching basso tra 10 e 40.
- se non pertinente, matching_percent 0 e is_supplier false.
- non inventare partita IVA, email o telefoni.
- se il nome azienda non è chiaro, lascia supplier_name vuoto.
- se non trovi indirizzo fisico, lascia physical_address vuoto.
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4.1-mini",
            messages=[
                {"role": "system", "content": "Rispondi solo JSON valido."},
                {"role": "user", "content": prompt}
            ],
            temperature=0
        )

        content = response.choices[0].message.content
        content = content.replace("```json", "")
        content = content.replace("```", "").strip()

        return json.loads(content)

    except Exception:
        return {}


def real_supplier_scouting(category, area="Italia"):
    discovered = discover_supplier_sites(category, area)

    rows = []
    seen = set()

    thresholds = [25, 15, 1]

    for threshold in thresholds:
        rows = []
        seen = set()

        for item in discovered:
            url = item["url"]

            crawled = crawl_supplier_site(url)
            ai_eval = evaluate_supplier_with_ai(category, crawled)

            matching = int(ai_eval.get("matching_percent", 0) or 0)

            supplier_name = str(ai_eval.get("supplier_name", "")).strip()

            if (
                not supplier_name
                or supplier_name.lower() in ["", "azienda", "company", "supplier"]
            ):
                supplier_name = supplier_name_from_domain(url)

            if matching < threshold:
                continue

            if supplier_name.lower() in seen:
                continue

            seen.add(supplier_name.lower())

            rows.append({
                "Nome fornitore": supplier_name,
                "Sito web": url,
                "Partita IVA": ", ".join(crawled.get("vat_numbers", [])),
                "Telefono": ", ".join(crawled.get("phones", [])),
                "Email": ", ".join(crawled.get("emails", [])),
                "Indirizzo fisico": ai_eval.get("physical_address", ""),
                "Prodotti / Servizi": ai_eval.get("products", ""),
                "% Matching": matching
            })

            if len(rows) >= 10:
                break

        if len(rows) >= 2:
            break

    df = pd.DataFrame(rows)

    wanted_cols = [
        "Nome fornitore",
        "Sito web",
        "Partita IVA",
        "Telefono",
        "Email",
        "Indirizzo fisico",
        "Prodotti / Servizi",
        "% Matching"
    ]

    if df.empty:
        return pd.DataFrame(columns=wanted_cols)

    df = df[wanted_cols]
    df = df.sort_values(by="% Matching", ascending=False)

    if len(df) > 10:
        df = df.head(10)

    return df


def save_scouting_to_database(df, package_name):
    inserted = 0

    for _, row in df.iterrows():
        insert_supplier(
            package=package_name,
            supplier_name=row.get("Nome fornitore", ""),
            vat_number=row.get("Partita IVA", ""),
            email=row.get("Email", ""),
            phone=row.get("Telefono", ""),
            address_nl1=row.get("Indirizzo fisico", ""),
            source_file="Scouting Web"
        )
        inserted += 1

    return inserted


# =====================================================
# DASHBOARD
# =====================================================

if section == "Dashboard":
    suppliers_df = load_suppliers()

    total_suppliers = 0
    recent_count = 0

    if not suppliers_df.empty:
        total_suppliers = suppliers_df["supplier_name"].nunique()

        suppliers_df["created_at"] = pd.to_datetime(
            suppliers_df["created_at"],
            errors="coerce"
        )

        limit = datetime.now() - timedelta(days=30)
        recent = suppliers_df[suppliers_df["created_at"] >= limit]
        recent_count = recent["supplier_name"].nunique()

    c1, c2 = st.columns(2)

    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{total_suppliers}</div>
            <div class="kpi-label">Fornitori totali nel database</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{recent_count}</div>
            <div class="kpi-label">Fornitori aggiunti ultimi 30 giorni</div>
        </div>
        """, unsafe_allow_html=True)


# =====================================================
# IMPORT
# =====================================================

elif section == "Importa":
    st.markdown("## Importa vendor storiche")

    uploaded_files = st.file_uploader(
        "Carica file Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:
        if st.button("Importa nel database"):
            total = 0

            for file in uploaded_files:
                try:
                    df = pd.read_excel(file)

                    inserted = normalize_vendor_dataframe(
                        df,
                        source_file=file.name
                    )

                    total += inserted

                    st.success(f"{file.name}: {inserted} righe importate")

                except Exception as e:
                    st.error(f"Errore {file.name}: {e}")

            st.success(f"Import completato. Totale: {total}")


# =====================================================
# GENERA VENDOR LIST
# =====================================================

elif section == "Vendor":
    st.markdown("## Genera vendor list")

    template_file = st.file_uploader(
        "Carica template Excel",
        type=["xlsx", "xls"]
    )

    threshold = st.slider(
        "Matching intelligente",
        50,
        95,
        70
    )

    if template_file:
        template_df = pd.read_excel(template_file)

        st.dataframe(
            template_df.head(20),
            use_container_width=True
        )

        if st.button("Genera vendor list"):
            try:
                result_df = generate_completed_vendor(
                    template_df,
                    threshold
                )

                st.success("Vendor list generata.")

                st.dataframe(
                    result_df,
                    use_container_width=True
                )

                excel_file = dataframe_to_excel(result_df)

                st.download_button(
                    label="Scarica Excel",
                    data=excel_file,
                    file_name="vendor_list_compilata.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(str(e))


# =====================================================
# SCOUTING
# =====================================================

elif section == "Scouting":
    st.markdown("## Scouting fornitori web")

    package_name = st.text_input("Categoria / lavorazione")

    area = st.text_input(
        "Area geografica",
        value="Italia"
    )

    st.info(
        "Il sistema cercherà automaticamente da 2 a 10 fornitori pertinenti."
    )

    if st.button("Avvia scouting web"):
        if not package_name.strip():
            st.warning("Inserisci una categoria.")
        else:
            with st.spinner("Almond Intelligence sta cercando fornitori..."):
                result_df = real_supplier_scouting(
                    package_name,
                    area
                )

            st.session_state.last_scouting_df = result_df
            st.session_state.last_scouting_package = package_name

    result_df = st.session_state.last_scouting_df

    if result_df is not None:
        if result_df.empty:
            st.warning("Nessun fornitore pertinente trovato.")
        else:
            st.success(f"Trovati {len(result_df)} fornitori.")

            st.dataframe(
                result_df,
                use_container_width=True
            )

            excel_file = dataframe_to_excel(
                result_df,
                "Scouting"
            )

            st.download_button(
                label="Scarica scouting Excel",
                data=excel_file,
                file_name=f"scouting_{st.session_state.last_scouting_package}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.markdown("### Approvi il risultato?")

            c1, c2 = st.columns(2)

            with c1:
                if st.button("Sì, aggiorna database"):
                    inserted = save_scouting_to_database(
                        result_df,
                        st.session_state.last_scouting_package
                    )

                    st.success(
                        f"Database aggiornato. Aggiunti: {inserted}"
                    )

            with c2:
                if st.button("No, solo Excel"):
                    st.info("Risultato non salvato.")


# =====================================================
# DATABASE
# =====================================================

elif section == "Database":
    st.markdown("## Database fornitori")

    df = load_suppliers()

    if df.empty:
        st.info("Database vuoto.")

    else:
        search = st.text_input("Cerca fornitore o categoria")

        if search:
            mask = (
                df["supplier_name"]
                .astype(str)
                .str.contains(search, case=False, na=False)
                |
                df["package"]
                .astype(str)
                .str.contains(search, case=False, na=False)
            )

            df_view = df[mask]

        else:
            df_view = df

        st.dataframe(
            df_view,
            use_container_width=True
        )

        excel_file = dataframe_to_excel(df_view)

        st.download_button(
            label="Esporta database Excel",
            data=excel_file,
            file_name="database_fornitori.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
