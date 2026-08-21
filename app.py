# -*- coding: utf-8 -*-
"""
Análisis Integral de la Flota Vehicular – PESV
(Seguridad Vial, Eficiencia y Desempeño)

App de Streamlit — carga un archivo Excel/CSV con la bitácora de flota
y muestra el dashboard completo: KPIs, aspectos destacados, recorrido,
combustible/eficiencia, seguridad, estacionamiento y huella de carbono.

Estructura esperada del archivo (mismos nombres de columnas que el
dashboard original en React):
    Placa ID | Fecha | Kilometraje(km) | Exceso de velocidad |
    Estacionamiento | Combustible(gal (us))

Cómo correrlo:
    streamlit run app.py

Cómo publicarlo como página web (gratis):
    1. Sube esta carpeta a un repo de GitHub (app.py + requirements.txt
       + informe_junio_demo.xlsx).
    2. Entra a https://share.streamlit.io , conecta el repo y despliega.
       Streamlit Community Cloud te da una URL pública.

Cómo correrlo en Google Colab:
    !pip install streamlit pandas plotly openpyxl pyngrok -q
    !streamlit run app.py &>/content/log.txt &
    from pyngrok import ngrok
    print(ngrok.connect(8501))
"""

import io
import re
from datetime import date
from xml.sax.saxutils import escape as _xml_escape

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURACIÓN DE PÁGINA
# ============================================================
st.set_page_config(
    page_title="Análisis Integral de la Flota Vehicular – PESV",
    page_icon="🚚",
    layout="wide",
)

# ============================================================
# PALETA DE COLORES — "tecnológica": azul, verde, morado
# Fondo más claro que un dashboard negro puro, pero se mantiene
# en modo oscuro para que los colores de acento resalten.
# ============================================================
C = {
    "bg": "#1B2436",         # fondo general (antes casi negro, ahora más claro)
    "bg_grad_end": "#141C2C",
    "panel": "#26314D",      # tarjetas / paneles
    "panel2": "#2E3B5C",     # inputs / elementos internos
    "border": "#3B4A6B",
    "border_light": "#4C5D82",
    "text": "#EAF0FF",
    "muted": "#9FB0D1",
    "muted2": "#6E7FA3",

    "blue": "#3B82F6",
    "blue_dim": "#16233F",
    "green": "#22C55E",
    "green_dim": "#12291D",
    "purple": "#A78BFA",
    "purple_dim": "#241F45",
    "cyan": "#22D3EE",
    "cyan_dim": "#12303A",
    "red": "#F87171",
    "red_dim": "#3A1E20",
}

FONT = "'Poppins', 'Nunito', sans-serif"

# ============================================================
# CSS — tipografía redondeada + fondo claro + estilo de filtros
# ============================================================
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&family=Nunito:wght@500;600;700;800&display=swap');

html, body, [class*="css"] {{
    font-family: {FONT} !important;
}}

.stApp {{
    background: linear-gradient(180deg, {C['bg']} 0%, {C['bg_grad_end']} 100%);
}}

h1, h2, h3, h4 {{
    font-family: {FONT} !important;
    color: {C['text']} !important;
    font-weight: 700 !important;
}}

p, span, label, div {{
    color: {C['text']};
}}

/* Tarjeta de encabezado */
.header-card {{
    background: {C['panel']};
    border: 1px solid {C['border']};
    border-radius: 16px;
    padding: 20px 26px;
    margin-bottom: 18px;
}}
.header-eyebrow {{
    text-transform: uppercase;
    letter-spacing: 0.12em;
    font-size: 11px;
    font-weight: 600;
    color: {C['blue']};
}}
.header-title {{
    font-size: 26px;
    font-weight: 800;
    color: {C['text']};
    line-height: 1.25;
    margin-top: 2px;
}}

/* KPI cards */
.kpi-card {{
    background: {C['panel']};
    border: 1px solid {C['border']};
    border-radius: 14px;
    padding: 16px 18px;
    height: 100%;
}}
.kpi-label {{
    text-transform: uppercase;
    font-size: 10.5px;
    letter-spacing: 0.07em;
    font-weight: 600;
    color: {C['muted']};
    margin-bottom: 8px;
}}
.kpi-value {{
    font-size: 24px;
    font-weight: 800;
    color: {C['text']};
}}
.kpi-unit {{
    font-size: 12px;
    color: {C['muted']};
    margin-left: 4px;
    font-weight: 500;
}}

/* Highlight cards */
.hl-card {{
    background: {C['panel']};
    border: 1px solid {C['border']};
    border-radius: 14px;
    padding: 14px 16px;
    height: 100%;
}}
.hl-label {{
    text-transform: uppercase;
    font-size: 10px;
    letter-spacing: 0.06em;
    font-weight: 600;
    color: {C['muted']};
    margin-bottom: 6px;
}}
.hl-primary {{
    font-weight: 700;
    font-size: 13px;
    color: {C['text']};
}}
.hl-value {{
    font-weight: 800;
    font-size: 15px;
}}

/* Section header */
.section-header {{
    display: flex;
    align-items: center;
    gap: 8px;
    margin: 6px 0 10px 0;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.10em;
    font-size: 13px;
}}

/* Chart panel */
.chart-panel {{
    background: {C['panel']};
    border: 1px solid {C['border']};
    border-radius: 14px;
    padding: 16px 18px 6px 18px;
    margin-bottom: 16px;
}}
.chart-title {{
    font-weight: 700;
    font-size: 14.5px;
    color: {C['text']};
}}
.chart-subtitle {{
    font-size: 11.5px;
    color: {C['muted']};
    margin-bottom: 6px;
}}

/* Inputs, multiselect, date input, buttons */
.stMultiSelect [data-baseweb="select"] > div {{
    background-color: {C['panel2']} !important;
    border-color: {C['border']} !important;
    border-radius: 10px !important;
    font-family: {FONT} !important;
    font-size: 13.5px !important;
}}
.stMultiSelect span[data-baseweb="tag"] {{
    background-color: {C['blue']} !important;
    border-radius: 8px !important;
    font-family: {FONT} !important;
    font-weight: 600 !important;
}}
.stDateInput input {{
    background-color: {C['panel2']} !important;
    color: {C['text']} !important;
    border-color: {C['border']} !important;
    border-radius: 10px !important;
    font-family: {FONT} !important;
}}
.stButton button {{
    background-color: {C['panel2']};
    color: {C['text']};
    border: 1px solid {C['border']};
    border-radius: 10px;
    font-family: {FONT};
    font-weight: 600;
    font-size: 13px;
}}
.stButton button:hover {{
    border-color: {C['blue']};
    color: {C['blue']};
}}
label, .stMultiSelect label, .stDateInput label {{
    font-family: {FONT} !important;
    font-weight: 600 !important;
    color: {C['muted']} !important;
    font-size: 12.5px !important;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}}
[data-testid="stFileUploader"] section {{
    background-color: {C['panel2']} !important;
    border-color: {C['border']} !important;
    border-radius: 12px !important;
}}
footer {{ visibility: hidden; }}
</style>
""", unsafe_allow_html=True)

# ============================================================
# FACTOR DE HUELLA DE CARBONO
# Fuente: EPA (Registro Federal 2010, base IPCC 2006):
#   8,887 g CO2 por galón de gasolina consumida.
# Se aplica sobre el combustible realmente consumido (columna
# Combustible(gal)), que a su vez está directamente asociado al
# kilometraje recorrido por cada vehículo — es más preciso que
# asumir un factor genérico de g/km para un "vehículo promedio",
# porque usa el consumo real registrado en la bitácora.
# ============================================================
CO2_KG_POR_GALON = 8.887

# ============================================================
# CONFIGURACIÓN DE IA (GROQ)
# El token se lee primero de st.secrets / variable de entorno
# GROQ_API_KEY (recomendado, sobre todo si el repo va a ser
# público en GitHub) y solo si no existe ninguna de las dos
# se usa el valor por defecto embebido aquí.
# ============================================================
import os

GROQ_MODEL = "openai/gpt-oss-120b"
try:
    _groq_key_from_secrets = st.secrets.get("GROQ_API_KEY")
except Exception:
    _groq_key_from_secrets = None
GROQ_API_KEY = (
    _groq_key_from_secrets
    or os.environ.get("GROQ_API_KEY")
    or "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"
)


# ============================================================
# CARGA Y LIMPIEZA DE DATOS
# ============================================================
REQUIRED_COLS = {
    "placa": ["placa"],
    "fecha": ["fecha"],
    "km": ["kilometraje"],
    "exceso": ["exceso"],
    "estac": ["estacionamiento"],
    "comb": ["combustible"],
}


def normalize_header(s: str) -> str:
    import unicodedata
    s = str(s).lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return "".join(ch for ch in s if ch.isalnum())


@st.cache_data(show_spinner=False)
def load_dataframe(file_bytes: bytes, filename: str) -> pd.DataFrame:
    if filename.lower().endswith(".csv"):
        raw = pd.read_csv(io.BytesIO(file_bytes))
    else:
        raw = pd.read_excel(io.BytesIO(file_bytes))

    col_map = {}
    for col in raw.columns:
        norm = normalize_header(col)
        for key, needles in REQUIRED_COLS.items():
            if key not in col_map and any(n in norm for n in needles):
                col_map[key] = col

    missing = [k for k in REQUIRED_COLS if k not in col_map]
    if missing:
        raise ValueError(f"Faltan columnas: {', '.join(missing)}")

    df = pd.DataFrame({
        "placa": raw[col_map["placa"]].astype(str).str.strip(),
        "fecha": pd.to_datetime(raw[col_map["fecha"]], errors="coerce"),
        "km": pd.to_numeric(raw[col_map["km"]], errors="coerce").fillna(0),
        "exceso": pd.to_numeric(raw[col_map["exceso"]], errors="coerce").fillna(0),
        "estac": pd.to_numeric(raw[col_map["estac"]], errors="coerce").fillna(0),
        "comb": pd.to_numeric(raw[col_map["comb"]], errors="coerce").fillna(0),
    })
    df = df.dropna(subset=["placa", "fecha"])
    df["fecha"] = df["fecha"].dt.date
    return df


# ============================================================
# ENCABEZADO + CARGA DE ARCHIVO
# ============================================================
left, right = st.columns([3, 1.3])
with left:
    st.markdown(f"""
    <div class="header-card">
        <div class="header-eyebrow">Panel de análisis vehicular · PESV</div>
        <div class="header-title">Análisis Integral de la Flota Vehicular – PESV<br/>
        <span style="font-size:15px; font-weight:500; color:{C['muted']};">
        (Seguridad Vial, Eficiencia y Desempeño)</span></div>
    </div>
    """, unsafe_allow_html=True)

with right:
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Cargar dataset (.xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"])

# Fuente de datos: archivo cargado o demo embebido
try:
    if uploaded is not None:
        df = load_dataframe(uploaded.getvalue(), uploaded.name)
        source_label = uploaded.name
    else:
        with open("informe_junio_demo.xlsx", "rb") as f:
            demo_bytes = f.read()
        df = load_dataframe(demo_bytes, "informe_junio_demo.xlsx")
        source_label = "informe_junio_demo.xlsx (precargado)"
except Exception as e:
    st.error(
        f"No se pudo interpretar el archivo ({e}). Verifica que tenga las columnas: "
        "Placa ID · Fecha · Kilometraje(km) · Exceso de velocidad · Estacionamiento · Combustible(gal (us))."
    )
    st.stop()

st.caption(f"Fuente de datos: **{source_label}**")

# ============================================================
# ESTADO DE FILTROS
# ============================================================
all_placas = sorted(df["placa"].unique().tolist())
min_date, max_date = df["fecha"].min(), df["fecha"].max()

if "selected_placas" not in st.session_state or st.session_state.get("_last_source") != source_label:
    st.session_state.selected_placas = all_placas.copy()
    st.session_state.date_range = (min_date, max_date)
    st.session_state["_last_source"] = source_label

# ============================================================
# PANEL DE FILTROS (se mantienen igual, solo cambia tipografía)
# ============================================================
st.markdown('<div class="header-card">', unsafe_allow_html=True)

fc1, fc2 = st.columns([2.2, 1])
with fc1:
    st.markdown(f"**🚚 Vehículos ({len(st.session_state.selected_placas)}/{len(all_placas)})**")
    selected_placas = st.multiselect(
        "Selecciona o busca por placa",
        options=all_placas,
        default=st.session_state.selected_placas,
        key="ms_placas",
        label_visibility="collapsed",
    )
with fc2:
    st.markdown("**Acciones rápidas**")
    b1, b2, b3 = st.columns(3)
    if b1.button("Todas", use_container_width=True):
        selected_placas = all_placas.copy()
    if b2.button("Ninguna", use_container_width=True):
        selected_placas = []
    if b3.button("↺ Reset", use_container_width=True):
        selected_placas = all_placas.copy()
        st.session_state.date_range = (min_date, max_date)

st.session_state.selected_placas = selected_placas

st.markdown("<div style='height:1px; background:#3B4A6B; margin:10px 0;'></div>", unsafe_allow_html=True)

dc1, dc2 = st.columns([1, 3])
with dc1:
    st.markdown("**📅 Periodo**")
with dc2:
    date_range = st.date_input(
        "Rango de fechas",
        value=st.session_state.date_range,
        min_value=min_date,
        max_value=max_date,
        label_visibility="collapsed",
    )
    if isinstance(date_range, tuple) and len(date_range) == 2:
        st.session_state.date_range = date_range

st.markdown('</div>', unsafe_allow_html=True)

date_start, date_end = st.session_state.date_range
if isinstance(date_start, tuple):
    date_start, date_end = date_start

# ============================================================
# DATOS FILTRADOS
# ============================================================
filtered = df[
    df["placa"].isin(selected_placas)
    & (df["fecha"] >= date_start)
    & (df["fecha"] <= date_end)
].copy()

if filtered.empty:
    st.markdown(f"""
    <div class="header-card" style="text-align:center; padding:50px;">
        <div style="font-size:15px; color:{C['muted']};">
        No hay datos para la combinación de filtros actual.<br/>
        Selecciona al menos un vehículo y un rango de fechas válido.</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ============================================================
# AGREGACIONES (misma lógica que el dashboard original)
# ============================================================
total_km = filtered["km"].sum()
total_comb = filtered["comb"].sum()
total_exceso = filtered["exceso"].sum()
total_estac = filtered["estac"].sum()
efficiency = (total_km / total_comb) if total_comb > 0 else 0
n_vehicles = len(selected_placas)

daily = (
    filtered.groupby("fecha", as_index=False)
    .agg(km=("km", "sum"), comb=("comb", "sum"), exceso=("exceso", "sum"), estac=("estac", "sum"))
    .sort_values("fecha")
)
daily["eficiencia"] = np.where(daily["comb"] > 0, daily["km"] / daily["comb"], 0)
daily["label"] = pd.to_datetime(daily["fecha"]).dt.strftime("%d/%m")
daily["co2_kg"] = daily["comb"] * CO2_KG_POR_GALON

per_vehicle = (
    filtered.groupby("placa", as_index=False)
    .agg(km=("km", "sum"), comb=("comb", "sum"), exceso=("exceso", "sum"),
         estac=("estac", "sum"), days=("fecha", "count"))
)
dias_con_exceso = filtered[filtered["exceso"] > 0].groupby("placa").size()
per_vehicle["diasConExceso"] = per_vehicle["placa"].map(dias_con_exceso).fillna(0).astype(int)
per_vehicle["avgKm"] = np.where(per_vehicle["days"] > 0, per_vehicle["km"] / per_vehicle["days"], 0)
per_vehicle["eficiencia"] = np.where(per_vehicle["comb"] > 0, per_vehicle["km"] / per_vehicle["comb"], 0)
per_vehicle["co2_kg"] = per_vehicle["comb"] * CO2_KG_POR_GALON
per_vehicle["co2_g_por_km"] = np.where(per_vehicle["km"] > 0, (per_vehicle["co2_kg"] * 1000) / per_vehicle["km"], 0)

avg_km_sorted = per_vehicle.sort_values("avgKm", ascending=False)
exceso_sorted = per_vehicle.sort_values("diasConExceso", ascending=False)
co2_sorted = per_vehicle.sort_values("co2_kg", ascending=False)

start_ts = pd.Timestamp(date_start)
filtered["_diff_days"] = (pd.to_datetime(filtered["fecha"]) - start_ts).dt.days
filtered["_week"] = (filtered["_diff_days"] // 7).clip(lower=0)
weekly = filtered.groupby("_week", as_index=False)["exceso"].sum()
weekly["label"] = "Sem. " + (weekly["_week"] + 1).astype(str)

total_co2_kg = total_comb * CO2_KG_POR_GALON
co2_intensity = (total_co2_kg * 1000 / total_km) if total_km > 0 else 0

highlight_km = per_vehicle.loc[per_vehicle["km"].idxmax()]
highlight_exceso = per_vehicle.loc[per_vehicle["diasConExceso"].idxmax()]
eff_candidates = per_vehicle[per_vehicle["comb"] > 0]
highlight_eff = eff_candidates.loc[eff_candidates["eficiencia"].idxmax()] if not eff_candidates.empty else None
highlight_day = daily.loc[daily["km"].idxmax()]

# ============================================================
# HELPERS DE FORMATO
# ============================================================
def nf0(v): return f"{v:,.0f}".replace(",", ".")
def nf1(v): return f"{v:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
def nf2(v): return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def kpi_card(label, value, unit=""):
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}<span class="kpi-unit">{unit}</span></div>
    </div>
    """, unsafe_allow_html=True)

def section_header(title, color):
    st.markdown(f"""
    <div class="section-header" style="color:{color};">
        <span>{title}</span>
        <span style="flex:1; height:1px; background:{C['border']};"></span>
    </div>
    """, unsafe_allow_html=True)

def chart_panel_open(title, subtitle):
    st.markdown(f"""
    <div class="chart-panel">
        <div class="chart-title">{title}</div>
        <div class="chart-subtitle">{subtitle}</div>
    """, unsafe_allow_html=True)

def chart_panel_close():
    st.markdown("</div>", unsafe_allow_html=True)

def build_fleet_summary_text() -> str:
    """Arma un resumen en texto plano de los datos filtrados actuales,
    que se envía como contexto al modelo de IA."""
    lines = []
    lines.append(f"Periodo analizado: {date_start} a {date_end}")
    lines.append(f"Vehículos incluidos ({n_vehicles}): {', '.join(selected_placas)}")
    lines.append("")
    lines.append("KPIs generales del periodo filtrado:")
    lines.append(f"- Kilometraje total: {nf0(total_km)} km")
    lines.append(f"- Combustible total consumido: {nf1(total_comb)} gal")
    lines.append(f"- Eficiencia promedio: {nf1(efficiency)} km/gal")
    lines.append(f"- Excesos de velocidad totales: {nf0(total_exceso)}")
    lines.append(f"- Estacionamientos totales: {nf0(total_estac)}")
    lines.append(f"- CO2 emitido total: {nf1(total_co2_kg)} kg")
    lines.append(f"- Intensidad de emisiones promedio: {nf0(co2_intensity)} g CO2/km")
    lines.append("")
    lines.append("Aspectos destacados:")
    lines.append(f"- Mayor recorrido: {highlight_km['placa']} con {nf0(highlight_km['km'])} km")
    lines.append(f"- Más excesos de velocidad: {highlight_exceso['placa']} con "
                 f"{nf0(highlight_exceso['diasConExceso'])} días")
    if highlight_eff is not None:
        lines.append(f"- Mejor eficiencia: {highlight_eff['placa']} con "
                     f"{nf1(highlight_eff['eficiencia'])} km/gal")
    lines.append(f"- Día de mayor actividad: "
                 f"{pd.to_datetime(highlight_day['fecha']).strftime('%d/%m/%Y')} "
                 f"con {nf0(highlight_day['km'])} km")
    lines.append("")
    lines.append("Detalle por vehículo (placa | km | combustible gal | eficiencia km/gal | "
                 "días con exceso | estacionamientos | CO2 kg):")
    for _, row in per_vehicle.sort_values("km", ascending=False).iterrows():
        lines.append(
            f"- {row['placa']} | {nf0(row['km'])} km | {nf1(row['comb'])} gal | "
            f"{nf1(row['eficiencia'])} km/gal | {int(row['diasConExceso'])} días | "
            f"{nf0(row['estac'])} paradas | {nf1(row['co2_kg'])} kg CO2"
        )
    return "\n".join(lines)


def call_groq_analysis(data_summary: str) -> str:
    """Envía el resumen de datos a Groq y devuelve el análisis en texto/markdown."""
    from groq import Groq

    client = Groq(api_key=GROQ_API_KEY)
    system_prompt = (
        "Eres un analista experto en gestión de flotas vehiculares y en el Plan Estratégico "
        "de Seguridad Vial (PESV) en Colombia. Recibes datos agregados y reales de una bitácora "
        "de flota (no debes inventar cifras adicionales) y debes producir un análisis profesional "
        "en español, claro y accionable, en formato markdown, organizado en estas secciones:\n"
        "1. Resumen ejecutivo\n"
        "2. Hallazgos sobre recorrido y eficiencia de combustible\n"
        "3. Riesgos de seguridad vial (excesos de velocidad)\n"
        "4. Huella de carbono\n"
        "5. Recomendaciones concretas y priorizadas\n"
        "Usa únicamente las cifras que se te entregan a continuación."
    )
    resp = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"Datos agregados de la flota:\n\n{data_summary}\n\n"
                                        "Genera el análisis completo."},
        ],
        temperature=0.4,
        max_tokens=2200,
    )
    return resp.choices[0].message.content


def _md_line_to_reportlab(line: str) -> str:
    """Escapa XML y convierte negrita markdown (**texto**) y viñetas (- ) a
    marcado simple compatible con reportlab Paragraph."""
    escaped = _xml_escape(line)
    escaped = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", escaped)
    escaped = re.sub(r"^-\s+", "• ", escaped)
    escaped = re.sub(r"^\d+\.\s+", lambda m: m.group(0), escaped)
    return escaped


def build_analysis_pdf(analysis_text: str, source_label: str, d_start, d_end) -> bytes:
    """Genera un PDF descargable con el análisis producido por la IA."""
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        leftMargin=2 * cm, rightMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm,
        title="Análisis de Flota Vehicular - IA",
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TituloPersonalizado", parent=styles["Title"], fontSize=18, spaceAfter=6)
    meta_style = ParagraphStyle("Meta", parent=styles["Normal"], textColor=colors.grey,
                                 fontSize=9, spaceAfter=16, leading=13)
    h_style = ParagraphStyle("Encabezado", parent=styles["Heading2"], spaceBefore=14, spaceAfter=6)
    body_style = ParagraphStyle("Cuerpo", parent=styles["Normal"], fontSize=10.5, leading=15, spaceAfter=7)

    story = [
        Paragraph("Análisis Integral de la Flota Vehicular – Generado por IA", title_style),
        Paragraph(
            f"Fuente de datos: {_xml_escape(source_label)}  ·  Periodo: {d_start} a {d_end}  ·  "
            f"Modelo: {GROQ_MODEL}  ·  Generado el: {date.today().isoformat()}",
            meta_style,
        ),
    ]

    for raw_line in analysis_text.split("\n"):
        block = raw_line.strip()
        if not block:
            story.append(Spacer(1, 6))
            continue
        if block.startswith("### "):
            story.append(Paragraph(_xml_escape(block[4:]), h_style))
        elif block.startswith("## "):
            story.append(Paragraph(_xml_escape(block[3:]), h_style))
        elif block.startswith("# "):
            story.append(Paragraph(_xml_escape(block[2:]), h_style))
        else:
            story.append(Paragraph(_md_line_to_reportlab(block), body_style))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def base_layout(fig, height=300):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT, color=C["muted"], size=12),
        hoverlabel=dict(bgcolor=C["panel2"], font=dict(family=FONT, color=C["text"])),
        showlegend=False,
        xaxis=dict(gridcolor=C["border"], zeroline=False),
        yaxis=dict(gridcolor=C["border"], zeroline=False),
    )
    return fig

# ============================================================
# KPIs
# ============================================================
k1, k2, k3, k4, k5, k6 = st.columns(6)
with k1: kpi_card("Kilometraje total", nf0(total_km), "km")
with k2: kpi_card("Combustible total", nf1(total_comb), "gal")
with k3: kpi_card("Excesos de velocidad", nf0(total_exceso), "")
with k4: kpi_card("Estacionamientos", nf0(total_estac), "")
with k5: kpi_card("Eficiencia promedio", nf1(efficiency), "km/gal")
with k6: kpi_card("Vehículos activos", nf0(n_vehicles), "")

st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

# ============================================================
# ASPECTOS DESTACADOS
# ============================================================
st.markdown(f"""<div class="section-header" style="color:{C['blue']};">🏆 Aspectos destacados</div>""",
            unsafe_allow_html=True)
h1, h2, h3, h4 = st.columns(4)
with h1:
    st.markdown(f"""<div class="hl-card"><div class="hl-label">Mayor recorrido</div>
    <div class="hl-primary">{highlight_km['placa']}</div>
    <div class="hl-value" style="color:{C['blue']};">{nf0(highlight_km['km'])} km</div></div>""",
    unsafe_allow_html=True)
with h2:
    st.markdown(f"""<div class="hl-card"><div class="hl-label">Más excesos de velocidad</div>
    <div class="hl-primary">{highlight_exceso['placa']}</div>
    <div class="hl-value" style="color:{C['red']};">{nf0(highlight_exceso['diasConExceso'])} días</div></div>""",
    unsafe_allow_html=True)
with h3:
    if highlight_eff is not None:
        st.markdown(f"""<div class="hl-card"><div class="hl-label">Mejor eficiencia</div>
        <div class="hl-primary">{highlight_eff['placa']}</div>
        <div class="hl-value" style="color:{C['green']};">{nf1(highlight_eff['eficiencia'])} km/gal</div></div>""",
        unsafe_allow_html=True)
with h4:
    day_label = pd.to_datetime(highlight_day["fecha"]).strftime("%d %b %Y")
    st.markdown(f"""<div class="hl-card"><div class="hl-label">Día de mayor actividad</div>
    <div class="hl-primary">{day_label}</div>
    <div class="hl-value" style="color:{C['blue']};">{nf0(highlight_day['km'])} km</div></div>""",
    unsafe_allow_html=True)

st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

# ============================================================
# SECCIÓN: RECORRIDO
# ============================================================
section_header("📈 Recorrido", C["blue"])

chart_panel_open("Kilometraje diario", "Suma de km recorridos por día — vehículos seleccionados")
fig = go.Figure()
fig.add_trace(go.Scatter(
    x=daily["label"], y=daily["km"], mode="lines", fill="tozeroy",
    line=dict(color=C["blue"], width=2.5),
    fillcolor="rgba(59,130,246,0.25)",
    hovertemplate="%{y:,.0f} km<extra></extra>",
))
base_layout(fig, 280)
st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
chart_panel_close()

col_a, col_b = st.columns(2)
with col_a:
    chart_panel_open("Promedio de km diarios por vehículo", "Intensidad de uso relativa entre vehículos")
    fig = go.Figure(go.Bar(
        x=avg_km_sorted["placa"], y=avg_km_sorted["avgKm"],
        marker_color=C["blue"], marker=dict(cornerradius=4),
        hovertemplate="%{y:,.0f} km/día<extra></extra>",
    ))
    base_layout(fig, 280)
    fig.update_xaxes(tickangle=-45)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    chart_panel_close()

with col_b:
    chart_panel_open("Kilometraje vs. combustible", "Relación diaria por vehículo — cada punto es un registro")
    fig = go.Figure(go.Scatter(
        x=filtered["comb"], y=filtered["km"], mode="markers",
        marker=dict(color=C["blue"], size=7, opacity=0.55),
        hovertemplate="Combustible: %{x:.2f} gal<br>Km: %{y:,.0f}<extra></extra>",
    ))
    base_layout(fig, 280)
    fig.update_xaxes(title="Combustible (gal)")
    fig.update_yaxes(title="Kilometraje")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    chart_panel_close()

# ============================================================
# SECCIÓN: COMBUSTIBLE Y EFICIENCIA
# ============================================================
section_header("⛽ Combustible y eficiencia", C["green"])
col_c, col_d = st.columns(2)
with col_c:
    chart_panel_open("Combustible diario", "Consumo total por día (gal)")
    fig = go.Figure(go.Bar(
        x=daily["label"], y=daily["comb"], marker_color=C["green"], marker=dict(cornerradius=4),
        hovertemplate="%{y:.2f} gal<extra></extra>",
    ))
    base_layout(fig, 260)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    chart_panel_close()

with col_d:
    chart_panel_open("Eficiencia diaria", "Kilómetros por galón de combustible")
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=daily["label"], y=daily["eficiencia"], mode="lines",
        line=dict(color=C["cyan"], width=2.5),
        hovertemplate="%{y:.1f} km/gal<extra></extra>",
    ))
    fig.add_hline(y=efficiency, line_dash="dash", line_color=C["muted"],
                   annotation_text="Promedio", annotation_font_color=C["muted"])
    base_layout(fig, 260)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    chart_panel_close()

# ============================================================
# SECCIÓN: SEGURIDAD
# ============================================================
section_header("⚠️ Seguridad", C["red"])
col_e, col_f = st.columns(2)
with col_e:
    chart_panel_open("Días con exceso de velocidad", "Cantidad de días con al menos un exceso, por vehículo")
    fig = go.Figure(go.Bar(
        x=exceso_sorted["placa"], y=exceso_sorted["diasConExceso"],
        marker_color=C["red"], marker=dict(cornerradius=4),
        hovertemplate="%{y} días<extra></extra>",
    ))
    base_layout(fig, 280)
    fig.update_xaxes(tickangle=-45)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    chart_panel_close()

with col_f:
    chart_panel_open("Excesos de velocidad semanales", "Total de incidentes agrupados por semana del periodo filtrado")
    fig = go.Figure(go.Bar(
        x=weekly["label"], y=weekly["exceso"], marker_color=C["red"], marker=dict(cornerradius=4),
        hovertemplate="%{y} excesos<extra></extra>",
    ))
    base_layout(fig, 280)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    chart_panel_close()

# ============================================================
# SECCIÓN: ESTACIONAMIENTO
# ============================================================
section_header("📍 Estacionamiento", C["purple"])
chart_panel_open("Estacionamientos diarios", "Cantidad total de paradas/estacionamientos registrados por día")
fig = go.Figure(go.Bar(
    x=daily["label"], y=daily["estac"], marker_color=C["purple"], marker=dict(cornerradius=4),
    hovertemplate="%{y} estacionamientos<extra></extra>",
))
base_layout(fig, 260)
st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
chart_panel_close()

# ============================================================
# SECCIÓN NUEVA: HUELLA DE CARBONO
# ============================================================
section_header("🌱 Huella de carbono", C["green"])

hc1, hc2 = st.columns(2)
with hc1:
    kpi_card("CO₂ emitido (periodo filtrado)", nf1(total_co2_kg), "kg")
with hc2:
    kpi_card("Intensidad promedio", nf0(co2_intensity), "g CO₂/km")

st.markdown(f"""
<div style="font-size:11.5px; color:{C['muted']}; margin:10px 0 16px 0; line-height:1.5;">
Metodología: se calcula multiplicando el combustible realmente consumido (gal) por el factor
oficial de la EPA/IPCC de <b>8,887 g CO₂ por galón de gasolina quemada</b>. Se usa el consumo real
en vez de un factor genérico por kilómetro, porque el consumo está directamente ligado al
kilometraje recorrido por cada vehículo en esta bitácora — es la forma más precisa de estimar la
huella con los datos disponibles.
</div>
""", unsafe_allow_html=True)

col_g, col_h = st.columns(2)
with col_g:
    chart_panel_open("CO₂ emitido por vehículo", "Kilogramos de CO₂ estimados, por vehículo (ordenado)")
    fig = go.Figure(go.Bar(
        x=co2_sorted["placa"], y=co2_sorted["co2_kg"],
        marker_color=C["green"], marker=dict(cornerradius=4),
        hovertemplate="%{y:.1f} kg CO₂<extra></extra>",
    ))
    base_layout(fig, 280)
    fig.update_xaxes(tickangle=-45)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    chart_panel_close()

with col_h:
    chart_panel_open("CO₂ diario emitido", "Kilogramos de CO₂ estimados por día, según combustible consumido")
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=daily["label"], y=daily["co2_kg"], mode="lines", fill="tozeroy",
        line=dict(color=C["green"], width=2.5),
        fillcolor="rgba(34,197,94,0.25)",
        hovertemplate="%{y:.1f} kg CO₂<extra></extra>",
    ))
    base_layout(fig, 280)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    chart_panel_close()

chart_panel_open("Intensidad de emisiones por vehículo", "Gramos de CO₂ por kilómetro recorrido — a menor valor, más eficiente/limpio")
fig = go.Figure(go.Bar(
    x=per_vehicle.sort_values("co2_g_por_km")["placa"],
    y=per_vehicle.sort_values("co2_g_por_km")["co2_g_por_km"],
    marker_color=C["purple"], marker=dict(cornerradius=4),
    hovertemplate="%{y:.0f} g CO₂/km<extra></extra>",
))
base_layout(fig, 260)
fig.update_xaxes(tickangle=-45)
st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
chart_panel_close()

# ============================================================
# SECCIÓN NUEVA: ANÁLISIS CON IA (GROQ)
# ============================================================
section_header("🤖 Análisis con IA", C["cyan"])

st.markdown(f"""
<div style="font-size:11.5px; color:{C['muted']}; margin:0 0 12px 0; line-height:1.5;">
Genera un análisis narrativo de los datos filtrados (KPIs, seguridad vial, eficiencia y huella de
carbono) usando IA — modelo <b>{GROQ_MODEL}</b> vía Groq. Luego puedes descargarlo como PDF.
</div>
""", unsafe_allow_html=True)

ai_col1, ai_col2 = st.columns([1, 3])
with ai_col1:
    generar_analisis = st.button("✨ Generar análisis con IA", use_container_width=True)

if generar_analisis:
    with st.spinner("Analizando datos de la flota con IA..."):
        try:
            resumen_datos = build_fleet_summary_text()
            texto_analisis = call_groq_analysis(resumen_datos)
            st.session_state["ai_analysis_text"] = texto_analisis
            st.session_state["ai_analysis_source"] = source_label
            st.session_state["ai_analysis_period"] = (date_start, date_end)
        except Exception as e:
            st.error(f"No se pudo generar el análisis con IA: {e}")

if st.session_state.get("ai_analysis_text"):
    chart_panel_open(
        "Análisis generado por IA",
        f"Modelo: {GROQ_MODEL} · Fuente: {st.session_state.get('ai_analysis_source', source_label)}",
    )
    st.markdown(st.session_state["ai_analysis_text"])
    chart_panel_close()

    try:
        _period = st.session_state.get("ai_analysis_period", (date_start, date_end))
        pdf_bytes = build_analysis_pdf(
            st.session_state["ai_analysis_text"],
            st.session_state.get("ai_analysis_source", source_label),
            _period[0], _period[1],
        )
        st.download_button(
            "⬇️ Descargar análisis en PDF",
            data=pdf_bytes,
            file_name=f"analisis_ia_flota_{date.today().isoformat()}.pdf",
            mime="application/pdf",
            use_container_width=False,
        )
    except Exception as e:
        st.error(f"No se pudo generar el PDF: {e}")

st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

# ============================================================
# FOOTER
# ============================================================
st.markdown(f"""
<div style="text-align:center; color:{C['muted2']}; font-size:11px; padding:16px 0;">
Estructura esperada del archivo: Placa ID · Fecha · Kilometraje(km) · Exceso de velocidad ·
Estacionamiento · Combustible(gal (us))
</div>
""", unsafe_allow_html=True)
