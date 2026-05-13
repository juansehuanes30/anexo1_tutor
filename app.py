import base64
import io
import re
import unicodedata
from copy import copy
from pathlib import Path

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
SI = "Sí"
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


def load_teacher_database(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    df = df.dropna(how="all").copy()
    mapping = {key: find_column(df, options) for key, options in COLUMN_MAP_CANDIDATES.items()}
    missing = [k for k, v in mapping.items() if v is None]
    if missing:
        st.warning("No se encontraron algunas columnas esperadas: " + ", ".join(missing))
    out = pd.DataFrame()
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


def get_template_metadata(template_bytes):
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = get_worksheet(wb)
    weeks = validation_options_for_cell(wb, ws, "A2") or [f"Semana {i}" for i in range(1, 9)]
    activities = []
    for col in range(ACTIVITY_START_COL, ACTIVITY_END_COL + 1):
        header = ws.cell(row=1, column=col).value
        activities.append(str(header).strip() if header else f"Actividad columna {col}")
    return ws.title, weeks, activities


def clear_writable_area(ws):
    for row in range(DATA_START_ROW, DATA_START_ROW + MAX_ROWS):
        for col in range(FIRST_COL, LAST_COL + 1):
            ws.cell(row=row, column=col).value = None


def write_records_to_template(template_bytes, records):
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = get_worksheet(wb)
    clear_writable_area(ws)
    if len(records) > MAX_ROWS:
        raise ValueError(f"La plantilla solo permite {MAX_ROWS} registros.")
    for offset, record in enumerate(records):
        row = DATA_START_ROW + offset
        values = [
            record.get("semana", ""),
            record.get("nombre", ""),
            record.get("cedula", ""),
            record.get("genero", ""),
            record.get("jornada", ""),
            record.get("cargo", ""),
            record.get("grado", ""),
            record.get("nivel", ""),
            record.get("dane_ee", ""),
            record.get("dane_sede", ""),
        ]
        for idx, value in enumerate(values, start=1):
            ws.cell(row=row, column=idx).value = value
        activity_values = record.get("actividades", {})
        for col in range(ACTIVITY_START_COL, ACTIVITY_END_COL + 1):
            header = str(ws.cell(row=1, column=col).value or f"Actividad columna {col}").strip()
            ws.cell(row=row, column=col).value = activity_values.get(header, NO)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def css():
    logo_b64 = img_to_base64(LOGO_PATH)
    st.markdown(f"""
    <style>
    .stApp {{
        background: linear-gradient(135deg, #290040 0%, #5b107a 45%, #a227b5 100%);
        color: #ffffff;
    }}
    [data-testid="stHeader"] {{ background: rgba(0,0,0,0); }}
    .main .block-container {{ padding-top: 1rem; max-width: 1250px; }}
    .hero {{
        background: radial-gradient(circle at top, rgba(255, 126, 221, .38), rgba(32, 0, 55, .72)), linear-gradient(120deg, #2d004d, #7a1b98);
        border: 1px solid rgba(255,255,255,.20);
        border-radius: 28px;
        padding: 28px 28px 22px;
        text-align: center;
        box-shadow: 0 18px 45px rgba(0,0,0,.28);
        margin-bottom: 20px;
    }}
    .hero img {{ max-height: 150px; margin-bottom: 8px; }}
    .hero h1 {{ font-size: 3.1rem; margin: 0; color: #fff; text-shadow: 0 3px 8px rgba(0,0,0,.25); }}
    .hero p {{ font-size: 1.15rem; color: #f6e9ff; margin-top: 8px; }}
    .card {{
        background: rgba(255,255,255,.93);
        color: #21122f;
        border-radius: 22px;
        padding: 22px;
        border: 1px solid rgba(255,255,255,.45);
        box-shadow: 0 14px 35px rgba(0,0,0,.20);
        margin-bottom: 16px;
    }}
    .step-title {{ font-size: 1.35rem; font-weight: 800; color: #3f0758; margin-bottom: 6px; }}
    .small-note {{ color: #5a4a68; font-size: .95rem; }}
    div[data-testid="stFileUploader"] section {{ background: rgba(255,255,255,.88); border-radius: 16px; }}
    .stButton button, .stDownloadButton button {{ border-radius: 14px; font-weight: 700; }}
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
