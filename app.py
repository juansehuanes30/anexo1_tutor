import base64
import io
import re
import unicodedata
from copy import copy
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

APP_TITLE = "Anexo 1 Tutor"
LOGO_PATH = Path(__file__).parent / "assets" / "logo_ptafi_transparente.png"
SHEET_CANDIDATES = ["ACTIVIDADES EN LOS EE", "ACCIONES EN LOS EE"]
MAX_ROWS = 500
DATA_START_ROW = 2
FIRST_COL = 1
LAST_COL = 33  # AG
ACTIVITY_START_COL = 11  # K
ACTIVITY_END_COL = 33  # AG
SI = "Si"
NO = "No"

st.set_page_config(page_title=APP_TITLE, page_icon="📘", layout="wide")


def img_to_base64(path: Path) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def slug(text: str) -> str:
    text = str(text or "").strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    return text


def normalize_yes_no(value: str) -> str:
    s = slug(value)
    return SI if s in {"si", "s", "yes"} else NO


def clean_dane(value) -> str:
    s = re.sub(r"\D", "", str(value or ""))
    return s


def get_worksheet(wb):
    for name in SHEET_CANDIDATES:
        if name in wb.sheetnames:
            return wb[name]
    raise ValueError(f"No se encontró una hoja llamada {SHEET_CANDIDATES}.")


def find_column(df: pd.DataFrame, possible_names):
    normalized = {slug(c): c for c in df.columns}
    for name in possible_names:
        s = slug(name)
        if s in normalized:
            return normalized[s]
    # Fallback: contains all words
    for col in df.columns:
        col_slug = slug(col)
        for name in possible_names:
            words = slug(name).split()
            if words and all(w in col_slug for w in words):
                return col
    return None


COLUMN_MAP_CANDIDATES = {
    "nombre": ["NOMBRE COMPLETO", "NOMBRE DOCENTE", "DOCENTE", "NOMBRES Y APELLIDOS", "NOMBRE"],
    "cedula": ["CEDULA", "CÉDULA", "DOCUMENTO", "NUMERO DE DOCUMENTO", "NÚMERO DE DOCUMENTO", "IDENTIFICACION"],
    "genero": ["GENERO", "GÉNERO", "SEXO"],
    "jornada": ["JORNADA"],
    "cargo": ["CARGO"],
    "grado": ["GRADO"],
    "nivel": ["NIVEL DE ENSEÑANZA", "NIVEL", "NIVEL ENSENANZA"],
}


def read_teacher_raw(uploaded_file):
    """Lee la base de docentes y detecta automáticamente la fila de encabezados.

    Algunas bases institucionales tienen un título en la primera fila y los encabezados
    reales aparecen más abajo. Esta función prueba varias filas hasta encontrar columnas
    como NOMBRE COMPLETO, CEDULA, GENERO, JORNADA, CARGO, GRADO y NIVEL DE ENSEÑANZA.
    """
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file)

    best_df = None
    best_score = -1
    best_header = 0
    for header_row in range(0, 10):
        uploaded_file.seek(0)
        try:
            candidate = pd.read_excel(uploaded_file, header=header_row)
        except Exception:
            continue
        candidate = candidate.dropna(how="all").copy()
        if candidate.empty:
            continue
        score = 0
        for options in COLUMN_MAP_CANDIDATES.values():
            if find_column(candidate, options) is not None:
                score += 1
        if score > best_score:
            best_df = candidate
            best_score = score
            best_header = header_row

    if best_df is None:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file)

    if best_header != 0:
        st.info(f"Se detectó automáticamente que los encabezados de la base están en la fila {best_header + 1}.")
    return best_df


def load_teacher_database(uploaded_file) -> pd.DataFrame:
    df = read_teacher_raw(uploaded_file)
    df = df.dropna(how="all").copy()
    mapping = {key: find_column(df, options) for key, options in COLUMN_MAP_CANDIDATES.items()}
    missing = [k for k, v in mapping.items() if v is None]
    if missing:
        st.warning("No se encontraron algunas columnas esperadas: " + ", ".join(missing))
    out = pd.DataFrame(index=df.index)
    for key, col in mapping.items():
        out[key] = df[col] if col else ""
    out["nombre"] = out["nombre"].astype(str).str.strip()
    out["cedula"] = out["cedula"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    out["genero"] = out["genero"].astype(str).str.strip()
    out["jornada"] = out["jornada"].astype(str).str.strip().replace({"Mañana y Tarde": "Mañana y tarde"})
    out["cargo"] = out["cargo"].astype(str).str.strip()
    out["grado"] = out["grado"].astype(str).str.strip()
    out["nivel"] = out["nivel"].astype(str).str.strip()
    out = out[out["nombre"].str.lower().ne("nan") & out["nombre"].ne("")].drop_duplicates(subset=["cedula", "nombre"])
    out["etiqueta"] = out["nombre"] + " — " + out["cedula"]
    return out.reset_index(drop=True)


def cell_in_range(cell_coord, sqref):
    for part in str(sqref).split():
        min_col, min_row, max_col, max_row = range_boundaries(part)
        # quick parse cell coordinate
        from openpyxl.utils.cell import coordinate_to_tuple
        row, col = coordinate_to_tuple(cell_coord)
        if min_row <= row <= max_row and min_col <= col <= max_col:
            return True
    return False


def extract_list_from_formula(wb, formula):
    if not formula:
        return []
    f = str(formula).strip()
    if f.startswith('"') and f.endswith('"'):
        return [x.strip() for x in f.strip('"').split(',') if x.strip()]
    f = f.replace("'", "")
    if "!" in f and ":" in f:
        sheet_name, rng = f.split("!", 1)
        sheet_name = sheet_name.replace("=", "").strip()
        rng = rng.replace("$", "")
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            vals = []
            for row in ws[rng]:
                for cell in row:
                    if cell.value not in (None, ""):
                        vals.append(str(cell.value).strip())
            return vals
    return []


def validation_options_for_cell(wb, ws, cell_coord):
    for dv in ws.data_validations.dataValidation:
        if dv.type == "list" and cell_in_range(cell_coord, dv.sqref):
            return extract_list_from_formula(wb, dv.formula1)
    return []


def list_values_from_hidden_sheet(wb, header_name):
    """Obtiene opciones desde la hoja oculta 'Listas' usando el encabezado dado."""
    if "Listas" not in wb.sheetnames:
        return []
    ws = wb["Listas"]
    target = slug(header_name)
    col_idx = None
    for cell in ws[1]:
        if slug(cell.value) == target:
            col_idx = cell.column
            break
    if col_idx is None:
        return []
    vals = []
    for row in range(2, ws.max_row + 1):
        value = ws.cell(row=row, column=col_idx).value
        if value not in (None, ""):
            vals.append(str(value).strip())
    return vals


def get_allowed_lists(template_bytes):
    """Lee listas esperadas sin guardar el archivo; así no elimina validaciones extendidas."""
    wb = load_workbook(io.BytesIO(template_bytes), data_only=True)
    return {
        "semana": list_values_from_hidden_sheet(wb, "Semana de acompañamiento"),
        "genero": list_values_from_hidden_sheet(wb, "Género"),
        "jornada": list_values_from_hidden_sheet(wb, "JORNADA"),
        "cargo": list_values_from_hidden_sheet(wb, "CARGO"),
        "grado": list_values_from_hidden_sheet(wb, "GRADO"),
        "nivel": list_values_from_hidden_sheet(wb, "Nivel_Enseñanza"),
        "valor": list_values_from_hidden_sheet(wb, "Valor") or [SI, NO],
    }


def normalize_to_allowed(value, allowed):
    """Ajusta un valor de la base al texto exacto de la lista desplegable de la plantilla."""
    if value in (None, ""):
        return ""
    text = str(value).strip()
    if not allowed:
        return text
    lookup = {slug(x): x for x in allowed}
    return lookup.get(slug(text), text)


def get_template_metadata(template_bytes):
    # Se usa openpyxl solo para leer encabezados y listas; no se guarda con openpyxl.
    wb = load_workbook(io.BytesIO(template_bytes), data_only=True)
    ws = get_worksheet(wb)
    allowed = get_allowed_lists(template_bytes)
    weeks = allowed.get("semana") or validation_options_for_cell(wb, ws, "A2") or [f"Semana {i}" for i in range(1, 9)]
    activities = []
    for col in range(ACTIVITY_START_COL, ACTIVITY_END_COL + 1):
        header = ws.cell(row=1, column=col).value
        activities.append(str(header).strip() if header else f"Actividad columna {col}")
    return ws.title, weeks, activities


def col_to_letter(col):
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def cell_ref(row, col):
    return f"{col_to_letter(col)}{row}"


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    if not m:
        return 0, 0
    letters, row = m.groups()
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - 64)
    return int(row), col


def find_activity_sheet_path(template_bytes):
    """Localiza el XML de la hoja de actividades dentro del .xlsx sin alterar el paquete."""
    with zipfile.ZipFile(io.BytesIO(template_bytes), "r") as zin:
        wb_xml = zin.read("xl/workbook.xml")
        rels_xml = zin.read("xl/_rels/workbook.xml.rels")
    ns_main = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
               "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    wb_root = ET.fromstring(wb_xml)
    rels_root = ET.fromstring(rels_xml)
    rel_map = {}
    for rel in rels_root:
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if rid and target:
            rel_map[rid] = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
    for sheet in wb_root.findall("m:sheets/m:sheet", ns_main):
        name = sheet.attrib.get("name", "")
        if name in SHEET_CANDIDATES:
            rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            return rel_map[rid]
    raise ValueError(f"No se encontró la hoja {SHEET_CANDIDATES} dentro del archivo.")


def set_cell_value(cell, value, numeric=False):
    """Escribe valor en un nodo <c> conservando atributos como estilo. Vacío deja solo formato."""
    # Limpiar hijos existentes y atributos de tipo
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)
    cell.attrib.pop("cm", None)
    cell.attrib.pop("vm", None)
    if value in (None, ""):
        return
    if numeric:
        cleaned = clean_dane(value)
        if cleaned:
            v = ET.SubElement(cell, "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
            v.text = cleaned
        return
    cell.set("t", "inlineStr")
    is_node = ET.SubElement(cell, "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is")
    t_node = ET.SubElement(is_node, "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
    t_node.text = str(value)


def write_records_to_template(template_bytes, records):
    """Escribe únicamente valores en el XML de la hoja, conservando validaciones, listas y estructura.

    A diferencia de guardar con openpyxl, este método no reescribe el libro completo ni elimina
    validaciones extendidas de Excel. Por eso se mantienen las listas desplegables de A, D-H y K-AG.
    """
    if len(records) > MAX_ROWS:
        raise ValueError(f"La plantilla solo permite {MAX_ROWS} registros.")

    allowed = get_allowed_lists(template_bytes)
    sheet_path = find_activity_sheet_path(template_bytes)

    # Registrar namespaces para conservar prefijos comunes.
    ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    ET.register_namespace('xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')
    ET.register_namespace('x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main')
    ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    ET.register_namespace('x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')
    ET.register_namespace('xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision')
    ET.register_namespace('xr2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2')
    ET.register_namespace('xr3', 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3')

    with zipfile.ZipFile(io.BytesIO(template_bytes), "r") as zin:
        sheet_xml = zin.read(sheet_path)
        root = ET.fromstring(sheet_xml)
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        sheet_data = root.find("m:sheetData", ns)
        if sheet_data is None:
            raise ValueError("La hoja de actividades no contiene sheetData.")

        # Índices de filas y celdas existentes.
        rows = {int(row.attrib.get("r")): row for row in sheet_data.findall("m:row", ns)}
        row2 = rows.get(DATA_START_ROW)
        base_attrs = {}
        if row2 is not None:
            for c in row2.findall("m:c", ns):
                _, col_idx = split_ref(c.attrib.get("r", ""))
                if FIRST_COL <= col_idx <= LAST_COL:
                    base_attrs[col_idx] = {k: v for k, v in c.attrib.items() if k != "r"}

        def get_or_create_row(row_num):
            row = rows.get(row_num)
            if row is not None:
                return row
            row = ET.Element("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row", {"r": str(row_num), "spans": "1:33"})
            # Insertar en orden.
            inserted = False
            for i, existing in enumerate(list(sheet_data)):
                if int(existing.attrib.get("r", 0)) > row_num:
                    sheet_data.insert(i, row)
                    inserted = True
                    break
            if not inserted:
                sheet_data.append(row)
            rows[row_num] = row
            return row

        def get_or_create_cell(row, row_num, col_num):
            ref = cell_ref(row_num, col_num)
            existing_cells = row.findall("m:c", ns)
            for c in existing_cells:
                if c.attrib.get("r") == ref:
                    return c
            attrs = {"r": ref}
            attrs.update(base_attrs.get(col_num, {}))
            c = ET.Element("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c", attrs)
            inserted = False
            for i, existing in enumerate(existing_cells):
                _, ec = split_ref(existing.attrib.get("r", ""))
                if ec > col_num:
                    row.insert(i, c)
                    inserted = True
                    break
            if not inserted:
                row.append(c)
            return c

        # Limpiar y reescribir A:AG manteniendo estilos. Las filas no usadas quedan vacías.
        for offset in range(MAX_ROWS):
            row_num = DATA_START_ROW + offset
            row = get_or_create_row(row_num)
            record = records[offset] if offset < len(records) else None
            if record:
                activity_values = record.get("actividades", {})
                base_values = {
                    1: normalize_to_allowed(record.get("semana", ""), allowed.get("semana", [])),
                    2: record.get("nombre", ""),
                    3: record.get("cedula", ""),
                    4: normalize_to_allowed(record.get("genero", ""), allowed.get("genero", [])),
                    5: normalize_to_allowed(record.get("jornada", ""), allowed.get("jornada", [])),
                    6: normalize_to_allowed(record.get("cargo", ""), allowed.get("cargo", [])),
                    7: normalize_to_allowed(record.get("grado", ""), allowed.get("grado", [])),
                    8: normalize_to_allowed(record.get("nivel", ""), allowed.get("nivel", [])),
                    9: clean_dane(record.get("dane_ee", "")),
                    10: clean_dane(record.get("dane_sede", "")),
                }
            else:
                activity_values = {}
                base_values = {}

            for col in range(FIRST_COL, LAST_COL + 1):
                c = get_or_create_cell(row, row_num, col)
                if not record:
                    set_cell_value(c, "")
                    continue
                if col <= 10:
                    # C, I y J quedan como número para evitar el aviso verde de Excel.
                    set_cell_value(c, base_values.get(col, ""), numeric=(col in {3, 9, 10}))
                else:
                    # Usar el texto exacto de la lista desplegable: normalmente Si/No.
                    header = ""
                    # Los encabezados se leen desde la fila 1 del mismo XML.
                    header_cell = None
                    header_row = rows.get(1)
                    if header_row is not None:
                        for hc in header_row.findall("m:c", ns):
                            if hc.attrib.get("r") == cell_ref(1, col):
                                header_cell = hc
                                break
                    # Para encabezados con sharedStrings no es necesario aquí: el diccionario ya viene de openpyxl por nombre.
                    # Se identifica por posición dentro de activities.
                    idx = col - ACTIVITY_START_COL
                    header = st.session_state.activities[idx] if 0 <= idx < len(st.session_state.activities) else f"Actividad columna {col}"
                    val = activity_values.get(header, NO)
                    val = normalize_to_allowed(val, allowed.get("valor", [SI, NO]))
                    set_cell_value(c, val, numeric=False)

        new_sheet_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new_sheet_xml if item.filename == sheet_path else zin.read(item.filename)
                zout.writestr(item, data)
    output.seek(0)
    return output.getvalue()


def css():
    logo_b64 = img_to_base64(LOGO_PATH)
    st.markdown(f"""
    <style>
    :root {{
        --morado-oscuro: #250034;
        --morado: #5b107a;
        --fucsia: #a227b5;
        --texto-oscuro: #21122f;
        --tarjeta: rgba(255,255,255,.96);
    }}

    .stApp {{
        background: linear-gradient(135deg, #250034 0%, #5b107a 45%, #a227b5 100%);
        color: #ffffff;
    }}

    [data-testid="stHeader"] {{
        background: rgba(0,0,0,0);
    }}

    .main .block-container {{
        padding-top: 1rem;
        max-width: 1250px;
    }}

    .hero {{
        background: radial-gradient(circle at top, rgba(255, 126, 221, .38), rgba(32, 0, 55, .72)), linear-gradient(120deg, #2d004d, #7a1b98);
        border: 1px solid rgba(255,255,255,.22);
        border-radius: 28px;
        padding: 30px 28px 24px;
        text-align: center;
        box-shadow: 0 18px 45px rgba(0,0,0,.28);
        margin-bottom: 22px;
    }}

    .hero img {{
        max-height: 150px;
        margin-bottom: 8px;
    }}

    .hero h1 {{
        font-size: 3.1rem;
        margin: 0;
        color: #ffffff !important;
        text-shadow: 0 3px 8px rgba(0,0,0,.25);
        font-weight: 900;
    }}

    .hero p {{
        font-size: 1.15rem;
        color: #f8ecff !important;
        margin-top: 8px;
        font-weight: 600;
    }}

    .card {{
        background: var(--tarjeta);
        color: var(--texto-oscuro) !important;
        border-radius: 22px;
        padding: 22px;
        border: 1px solid rgba(255,255,255,.55);
        box-shadow: 0 14px 35px rgba(0,0,0,.20);
        margin-bottom: 18px;
    }}

    .card * {{
        color: var(--texto-oscuro) !important;
    }}

    .step-title {{
        font-size: 1.35rem;
        font-weight: 900;
        color: #3f0758 !important;
        margin-bottom: 6px;
    }}

    .small-note {{
        color: #4d365f !important;
        font-size: .98rem;
        font-weight: 500;
    }}

    /* Texto general sobre el fondo morado */
    h1, h2, h3, h4, h5, h6,
    .stMarkdown, .stMarkdown p,
    [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] li {{
        color: #ffffff;
    }}

    /* Títulos de Streamlit */
    h2, h3 {{
        color: #ffffff !important;
        font-weight: 900 !important;
        text-shadow: 0 2px 6px rgba(0,0,0,.25);
    }}

    /* Etiquetas de inputs, selectbox, radio, multiselect y uploader */
    .stTextInput label,
    .stSelectbox label,
    .stMultiSelect label,
    .stRadio label,
    .stFileUploader label,
    .stTextArea label,
    [data-testid="stWidgetLabel"] p {{
        color: #ffffff !important;
        font-weight: 800 !important;
    }}

    /* Opciones de radio y textos auxiliares sobre fondo morado */
    .stRadio [role="radiogroup"] label,
    .stRadio [role="radiogroup"] span,
    .stRadio [role="radiogroup"] p {{
        color: #ffffff !important;
        font-weight: 700 !important;
    }}

    /* Campos de entrada: texto oscuro sobre caja clara */
    input, textarea,
    [data-baseweb="input"] input,
    [data-baseweb="select"] div,
    [data-baseweb="textarea"] textarea {{
        color: #21122f !important;
        caret-color: #21122f !important;
    }}

    /* Selectbox y multiselect */
    [data-baseweb="select"] {{
        background-color: #ffffff !important;
        border-radius: 12px !important;
    }}

    [data-baseweb="select"] * {{
        color: #21122f !important;
    }}

    /* Uploader */
    div[data-testid="stFileUploader"] section {{
        background: rgba(255,255,255,.94) !important;
        border-radius: 16px !important;
        border: 1px solid rgba(255,255,255,.70) !important;
    }}

    div[data-testid="stFileUploader"] section * {{
        color: #21122f !important;
    }}

    /* Alertas */
    [data-testid="stAlert"] {{
        background-color: #f3e8ff !important;
        border: 1px solid rgba(255,255,255,.65) !important;
        border-radius: 14px !important;
    }}

    [data-testid="stAlert"] * {{
        color: #21122f !important;
        font-weight: 600 !important;
    }}

    /* Métricas y tablas */
    [data-testid="stMetric"] {{
        background: rgba(255,255,255,.94);
        border-radius: 16px;
        padding: 14px;
    }}

    [data-testid="stMetric"] * {{
        color: #21122f !important;
    }}

    /* Botones */
    .stButton button, .stDownloadButton button {{
        border-radius: 14px;
        font-weight: 800;
        min-height: 44px;
    }}

    /* Captions */
    .stCaptionContainer, .stCaptionContainer p {{
        color: #f8ecff !important;
        font-weight: 600;
    }}
    </style>
    <div class="hero">
        <img src="data:image/png;base64,{logo_b64}" alt="PTAFI 3.0" />
        <h1>Anexo 1 Tutor</h1>
        <p>Legalización de actividades PTAFI 3.0 sobre la plantilla oficial del Ministerio.</p>
    </div>
    """, unsafe_allow_html=True)


css()

if "template_bytes" not in st.session_state:
    st.session_state.template_bytes = None
if "teacher_df" not in st.session_state:
    st.session_state.teacher_df = None
if "records" not in st.session_state:
    st.session_state.records = []
if "activities" not in st.session_state:
    st.session_state.activities = []
if "weeks" not in st.session_state:
    st.session_state.weeks = []

st.markdown('<div class="card"><div class="step-title">PASO 0 — Configuración inicial</div><div class="small-note">Carga la plantilla oficial y la base docente. Si ya cargaste archivos en esta sesión, puedes conservarlos.</div></div>', unsafe_allow_html=True)

with st.container():
    col_name, col_delivery, col_dane, col_sede = st.columns([1.4, .8, 1, 1])
    tutor_name = col_name.text_input("Nombre del tutor", placeholder="Ejemplo: David")
    entrega = col_delivery.selectbox("Número de entrega", [f"E{i}" for i in range(1, 11)], index=0)
    dane_ee = col_dane.text_input("Código DANE del EE", placeholder="218150000578", max_chars=12)
    dane_sede = col_sede.text_input("Código DANE de la sede", placeholder="218150000578", max_chars=12)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📄 Anexo 1 oficial")
        template_mode = st.radio("¿Cómo desea trabajar el Anexo 1?", ["Cargar nuevo archivo", "Conservar archivo actual"], horizontal=True, key="template_mode")
        if template_mode == "Cargar nuevo archivo" or st.session_state.template_bytes is None:
            template_file = st.file_uploader("Cargar Anexo 1 del PTAFI (.xlsx)", type=["xlsx"], key="template_uploader")
            if template_file is not None:
                st.session_state.template_bytes = template_file.getvalue()
                try:
                    sheet_name, weeks, activities = get_template_metadata(st.session_state.template_bytes)
                    st.session_state.weeks = weeks
                    st.session_state.activities = activities
                    st.success(f"Plantilla cargada. Hoja activa: {sheet_name}. Actividades detectadas: {len(activities)}.")
                except Exception as e:
                    st.error(f"No fue posible leer la plantilla: {e}")
                    st.session_state.template_bytes = None
        else:
            st.info("Se conservará el Anexo 1 cargado en esta sesión.")

    with c2:
        st.subheader("👩‍🏫 Base de datos de docentes")
        base_mode = st.radio("¿Cómo desea trabajar la base de docentes?", ["Cargar nueva base de datos", "Conservar base actual"], horizontal=True, key="base_mode")
        if base_mode == "Cargar nueva base de datos" or st.session_state.teacher_df is None:
            base_file = st.file_uploader("Cargar base de docentes (.xlsx o .csv)", type=["xlsx", "csv"], key="base_uploader")
            if base_file is not None:
                try:
                    st.session_state.teacher_df = load_teacher_database(base_file)
                    st.success(f"Base cargada. Docentes detectados: {len(st.session_state.teacher_df)}.")
                except Exception as e:
                    st.error(f"No fue posible leer la base docente: {e}")
                    st.session_state.teacher_df = None
        else:
            st.info("Se conservará la base de docentes cargada en esta sesión.")

ready = st.session_state.template_bytes is not None and st.session_state.teacher_df is not None
if not ready:
    st.info("Carga el Anexo 1 oficial y la base de datos para activar el flujo de diligenciamiento.")
    st.stop()

if len(clean_dane(dane_ee)) != 12 or len(clean_dane(dane_sede)) != 12:
    st.warning("Los códigos DANE del EE y de la sede deben tener 12 dígitos. Puedes continuar ajustándolos antes de descargar.")

df = st.session_state.teacher_df.copy()
activities = st.session_state.activities
weeks = st.session_state.weeks or [f"Semana {i}" for i in range(1, 9)]

st.markdown('<div class="card"><div class="step-title">PASO 1 — Seleccionar semana</div></div>', unsafe_allow_html=True)
semana = st.selectbox("Semana de acompañamiento", weeks)

st.markdown('<div class="card"><div class="step-title">PASO 2 — Seleccionar docentes</div><div class="small-note">Usa modo individual o grupal. Puedes filtrar para encontrar docentes más rápido.</div></div>', unsafe_allow_html=True)

filter_cols = st.columns(3)
for label, field, col in [("Filtrar por jornada", "jornada", filter_cols[0]), ("Filtrar por grado", "grado", filter_cols[1]), ("Filtrar por nivel", "nivel", filter_cols[2])]:
    vals = sorted([v for v in df[field].dropna().unique().tolist() if str(v).strip() and str(v).lower() != "nan"])
    selected_filter = col.multiselect(label, vals, key=f"filter_{field}")
    if selected_filter:
        df = df[df[field].isin(selected_filter)]

modo = st.radio("Modo de registro", ["Individual", "Grupal"], horizontal=True)
if modo == "Individual":
    selected = st.selectbox("Seleccionar docente", df["etiqueta"].tolist())
    selected_labels = [selected] if selected else []
else:
    selected_labels = st.multiselect("Seleccionar varios docentes", df["etiqueta"].tolist())

st.markdown('<div class="card"><div class="step-title">PASO 3 — Marcar actividades realizadas</div><div class="small-note">Marca solo las actividades realizadas. Las no seleccionadas se guardarán automáticamente como “No”.</div></div>', unsafe_allow_html=True)
selected_activities = st.multiselect("Actividades con respuesta Sí", activities)

col_add, col_reset = st.columns([1, 1])
with col_add:
    add_clicked = st.button("➕ Agregar registro(s)", use_container_width=True, type="primary")
with col_reset:
    if st.button("🧹 Limpiar registros agregados", use_container_width=True):
        st.session_state.records = []
        st.rerun()

if add_clicked:
    if not selected_labels:
        st.error("Selecciona al menos un docente.")
    else:
        activity_dict = {activity: (SI if activity in selected_activities else NO) for activity in activities}
        current_df = st.session_state.teacher_df
        added = 0
        for label in selected_labels:
            row = current_df[current_df["etiqueta"] == label]
            if row.empty:
                continue
            r = row.iloc[0].to_dict()
            st.session_state.records.append({
                "semana": semana,
                "nombre": r.get("nombre", ""),
                "cedula": r.get("cedula", ""),
                "genero": r.get("genero", ""),
                "jornada": r.get("jornada", ""),
                "cargo": r.get("cargo", ""),
                "grado": r.get("grado", ""),
                "nivel": r.get("nivel", ""),
                "dane_ee": clean_dane(dane_ee),
                "dane_sede": clean_dane(dane_sede),
                "actividades": activity_dict.copy(),
            })
            added += 1
        st.success(f"Se agregaron {added} registro(s). Total acumulado: {len(st.session_state.records)}.")

st.markdown('<div class="card"><div class="step-title">PASO 4 — ¿Agregar más o finalizar?</div><div class="small-note">Si necesitas más docentes o actividades, repite los pasos 1 a 3. Cuando termines, descarga el archivo final.</div></div>', unsafe_allow_html=True)

records = st.session_state.records
st.metric("Registros agregados", len(records))
if records:
    preview = []
    for r in records:
        yes_count = sum(1 for v in r["actividades"].values() if v == SI)
        preview.append({"Semana": r["semana"], "Docente": r["nombre"], "Cédula": r["cedula"], "Actividades Sí": yes_count})
    st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)

filename = f"{entrega}_{clean_dane(dane_ee) or 'DANE'}.xlsx"
if records:
    try:
        output_bytes = write_records_to_template(st.session_state.template_bytes, records)
        st.download_button(
            label=f"⬇️ Descargar archivo diligenciado: {filename}",
            data=output_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        st.caption("El archivo descargado conserva la plantilla original y solo escribe en las columnas A a AG de la hoja de actividades.")
    except Exception as e:
        st.error(f"No fue posible generar el archivo final: {e}")
else:
    st.info("Agrega al menos un registro para habilitar la descarga final.")
