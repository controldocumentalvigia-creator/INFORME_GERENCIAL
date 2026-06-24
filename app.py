# -*- coding: utf-8 -*-
"""
Informe Gerencial Ejecutivo Operativo y Financiero
Streamlit Cloud - versión robusta con cargador dinámico de base.

Características:
- Carga archivos Excel/CSV desde la barra lateral.
- Soporta columnas con nombres variables.
- No se cae si Plotly o AgGrid no están instalados: usa visuales nativos de Streamlit.
- Excluye ESTADO OP = ANULADO.
- Calcula facturación, costos, margen, rentabilidad y variaciones.
- Incluye resumen gerencial, mapas de calor, alertas, clientes, operación, financiero,
  rentabilidad, responsables, flota/conductores y mapa estratégico cuando existan coordenadas.
"""

from __future__ import annotations

import io
import re
import hashlib
import itertools
import uuid
import unicodedata
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# Plotly es opcional para evitar ModuleNotFoundError si Streamlit no instaló dependencias.
try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_OK = True
except Exception:
    PLOTLY_OK = False
    px = None
    go = None

# AgGrid es opcional.
try:
    from st_aggrid import AgGrid, GridOptionsBuilder
    AGGRID_OK = True
except Exception:
    AGGRID_OK = False

st.set_page_config(
    page_title="Informe Gerencial Ejecutivo",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================
# ESTILO CORPORATIVO
# =========================
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1500px;}
    html, body, [class*="css"] {color:#0B1220 !important;}
    [data-testid="stAppViewContainer"] {background:#F6F8FB;}
    [data-testid="stSidebar"] {background: linear-gradient(180deg,#08324C 0%,#0B4F78 100%);} 
    [data-testid="stSidebar"] * {color: white !important; font-weight:700;}
    h1, h2, h3 {color:#0B1220 !important; font-weight: 900;}
    p, span, label, div {color:#0B1220;}
    .main-title {font-size: 2.35rem; font-weight: 900; color:#0B1220; line-height:1.12; margin-bottom:.2rem;}
    .subtitle {font-size: 1.05rem; color:#1F2937; margin-bottom:1.2rem; font-weight:600;}
    .metric-card {
        background:#FFFFFF; border:1px solid #E5E7EB; border-radius:18px; padding:16px 18px;
        box-shadow:0 6px 18px rgba(15,23,42,.08); min-height:112px; border-left:6px solid #0B4F78;
    }
    .metric-title {font-size:.82rem; color:#111827 !important; font-weight:900; text-transform:uppercase; letter-spacing:.02rem;}
    .metric-value {font-size:1.45rem; color:#000000 !important; font-weight:950; margin-top:6px; overflow-wrap:anywhere;}
    .metric-note {font-size:.82rem; color:#1F2937 !important; margin-top:6px; font-weight:800;}
    [data-testid="stMetricValue"] {color:#000000 !important; font-weight:900 !important;}
    [data-testid="stMetricLabel"] {color:#111827 !important; font-weight:800 !important;}
    [data-testid="stMetricDelta"] {font-weight:900 !important;}
    .stDataFrame, .stTable {color:#000000 !important;}
    .card {background:#FFFFFF; border:1px solid #E5E7EB; border-radius:18px; padding:18px; box-shadow:0 6px 18px rgba(15,23,42,.06); margin-bottom:16px;}
    .alert-red {background:#FEE2E2; color:#7F1D1D; padding:11px 14px; border-radius:12px; margin:8px 0; font-weight:800;}
    .alert-yellow {background:#FEF3C7; color:#78350F; padding:11px 14px; border-radius:12px; margin:8px 0; font-weight:800;}
    .alert-green {background:#DCFCE7; color:#14532D; padding:11px 14px; border-radius:12px; margin:8px 0; font-weight:800;}
    .info-box {background:#EFF6FF; color:#1E3A8A; padding:13px 15px; border-radius:12px; margin:8px 0; font-weight:700;}
    .warn-box {background:#FFFBEB; color:#92400E; padding:13px 15px; border-radius:12px; margin:8px 0; font-weight:700;}
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# TEMA VISUAL OSCURO PARA GRÁFICOS
# =========================
PLOTLY_LAYOUT = dict(
    font=dict(family="Arial, sans-serif", size=14, color="#000000"),
    title_font=dict(family="Arial, sans-serif", size=20, color="#000000"),
    xaxis=dict(
        tickfont=dict(size=13, color="#000000"),
        title_font=dict(size=14, color="#000000"),
        showgrid=True,
        gridcolor="#D1D5DB",
        linecolor="#111827",
        zerolinecolor="#9CA3AF",
    ),
    yaxis=dict(
        tickfont=dict(size=13, color="#000000"),
        title_font=dict(size=14, color="#000000"),
        showgrid=True,
        gridcolor="#D1D5DB",
        linecolor="#111827",
        zerolinecolor="#9CA3AF",
    ),
    legend=dict(font=dict(size=13, color="#000000")),
)

def apply_dark_plotly_theme(fig):
    """Aplica letras y números oscuros a cualquier gráfico Plotly."""
    if fig is None:
        return fig
    fig.update_layout(**PLOTLY_LAYOUT)
    fig.update_coloraxes(colorbar=dict(tickfont=dict(color="#000000", size=12), title_font=dict(color="#000000", size=13)))
    return fig

MESES = {1:"Ene", 2:"Feb", 3:"Mar", 4:"Abr", 5:"May", 6:"Jun", 7:"Jul", 8:"Ago", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dic", 0:"Sin fecha"}
MESES_ORDEN = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic", "Sin fecha"]

# Coordenadas aproximadas para fallback por ciudad, si no existen lat/lon.
COORD_CIUDADES = {
    "BOGOTA": (4.7110, -74.0721), "BOGOTÁ": (4.7110, -74.0721), "IBAGUE": (4.4389, -75.2322), "IBAGUÉ": (4.4389, -75.2322),
    "MEDELLIN": (6.2442, -75.5812), "MEDELLÍN": (6.2442, -75.5812), "CALI": (3.4516, -76.5320), "BARRANQUILLA": (10.9685, -74.7813),
    "CARTAGENA": (10.3910, -75.4794), "BUCARAMANGA": (7.1193, -73.1227), "VILLAVICENCIO": (4.1420, -73.6266), "YOPAL": (5.3378, -72.3959),
    "TUNJA": (5.5353, -73.3678), "PEREIRA": (4.8087, -75.6906), "MANIZALES": (5.0703, -75.5138), "ARMENIA": (4.5339, -75.6811),
    "NEIVA": (2.9345, -75.2809), "PASTO": (1.2136, -77.2811), "CUCUTA": (7.8939, -72.5078), "CÚCUTA": (7.8939, -72.5078),
    "MONTERIA": (8.7500, -75.8833), "MONTERÍA": (8.7500, -75.8833), "SANTA MARTA": (11.2408, -74.1990), "RIOHACHA": (11.5444, -72.9072),
    "VALLEDUPAR": (10.4631, -73.2532), "SINCELEJO": (9.3047, -75.3978), "POPAYAN": (2.4448, -76.6147), "POPAYÁN": (2.4448, -76.6147),
    "SOACHA": (4.5794, -74.2168), "SIBATE": (4.4915, -74.2590), "SIBATÉ": (4.4915, -74.2590), "CHIA": (4.8610, -74.0587), "CHÍA": (4.8610, -74.0587),
}

COLUMN_CANDIDATES: Dict[str, List[str]] = {
    "fecha": ["CARGA", "FECHA", "FECHA SERVICIO", "FECHA DE SERVICIO", "FECHA PRESTACION", "FECHA PRESTACIÓN", "F SERVICIO", "DIA"],
    "cliente": ["CLIENTE", "NOMBRE CLIENTE", "RAZON SOCIAL", "RAZÓN SOCIAL", "EMPRESA", "CUENTA"],
    "v_cliente": ["V.CLIENTE", "V CLIENTE", "VALOR CLIENTE", "FACTURACION", "FACTURACIÓN", "VALOR FACTURADO", "VALOR VENTA", "VENTA"],
    "v_conduct": ["V.CONDUCT", "V CONDUCT", "VALOR CONDUCTOR", "VALOR CONDUCT", "COSTO", "VALOR COSTO", "FLETE", "V FLETE", "COSTO CONDUCTOR"],
    "estado_op": ["ESTADO OP", "ESTADO OPERATIVO", "ESTADO", "ESTADO SERVICIO", "ESTADO DEL SERVICIO"],
    "tray_prop": ["TRAY. PROP", "TRAY PROP", "TRAYECTO PROP", "TRAYECTO PROPIO", "FLOTA", "PROPIO TERCERO", "PROPIO/TERCERO"],
    "linea_neg": ["LINEA NEG", "LÍNEA NEG", "LINEA NEGOCIO", "LÍNEA NEGOCIO", "LINEA DE NEGOCIO"],
    "tipo_negocio": ["T.NEGOCIO", "T NEGOCIO", "TIPO NEGOCIO", "TIPO DE NEGOCIO", "NEGOCIO"],
    "placa": ["PLACA", "VEHICULO", "VEHÍCULO", "PLACA VEHICULO", "PLACA VEHÍCULO"],
    "conductor": ["CONDUCTO", "CONDUCTOR", "NOMBRE CONDUCTOR", "NOM CONDUCTOR"],
    "origen": ["ORIGEN", "CIUDAD ORIGEN", "MUNICIPIO ORIGEN", "PUNTO ORIGEN"],
    "destino": ["DESTINO", "CIUDAD DESTINO", "MUNICIPIO DESTINO", "PUNTO DESTINO"],
    "tipo_vehiculo": ["T. VEHICULO", "T VEHICULO", "TIPO VEHICULO", "TIPO VEHÍCULO", "CLASE VEHICULO"],
    "coordinador": ["COORDINADOR", "COORD", "RESPONSABLE", "EJECUTIVO", "ANALISTA"],
    "usuario": ["USUARIO", "CREADO POR", "CREADOR", "QUIEN CREO", "QUIÉN CREÓ", "CREADO", "USER", "USUARIO CREA", "CREADO_POR"],
    "lat": ["LAT", "LATITUD", "LATITUDE", "LAT_ORIGEN"],
    "lon": ["LON", "LONG", "LONGITUD", "LONGITUDE", "LNG", "LON_ORIGEN"],
}

# =========================
# FUNCIONES BASE
# =========================
def normalize_text(value: object) -> str:
    if value is None:
        return ""
    txt = str(value).strip().upper()
    txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
    txt = re.sub(r"\s+", " ", txt)
    return txt


def normalize_col(col: object) -> str:
    txt = normalize_text(col)
    txt = txt.replace("_", " ").replace("-", " ").replace("/", " ")
    txt = re.sub(r"[^A-Z0-9 ]", "", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def first_existing(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Busca columnas candidatas.

    Primero intenta coincidencia exacta normalizada; luego coincidencia parcial.
    La validación fina de columnas sensibles, como conductor, se hace aparte para
    evitar confundir VALOR CONDUCTOR / V.CONDUCT con el nombre del conductor.
    """
    norm_map = {normalize_col(c): c for c in df.columns}
    for cand in candidates:
        key = normalize_col(cand)
        if key in norm_map:
            return norm_map[key]
    for c in df.columns:
        n = normalize_col(c)
        for cand in candidates:
            key = normalize_col(cand)
            if key and key in n:
                return c
    return None


STRICT_CONDUCTOR_NAMES = {
    "CONDUCTOR",
    "CONDUCTO",
    "NOMBRE CONDUCTOR",
    "NOM CONDUCTOR",
    "NOMBRE DEL CONDUCTOR",
    "NOMBRE CONDUCTO",
}

BAD_CONDUCTOR_TOKENS = {
    "VALOR", "V CLIENTE", "V CONDUCT", "VCONDUCT", "COSTO",
    "FLETE", "TARIFA", "MARGEN", "RENTABILIDAD", "FACTURACION", "FACTURACIÓN"
}


def is_bad_conductor_column_name(col: object) -> bool:
    n = normalize_col(col)
    return any(tok in n for tok in BAD_CONDUCTOR_TOKENS)


def numeric_ratio(series: pd.Series) -> float:
    if series.empty:
        return 0.0
    sample = series.dropna().astype(str).str.strip()
    sample = sample[sample.ne("")].head(1000)
    if sample.empty:
        return 0.0
    nums = pd.to_numeric(sample.str.replace(r"[^0-9,.-]", "", regex=True).str.replace(",", ".", regex=False), errors="coerce")
    return float(nums.notna().mean())


def resolve_conductor_column(df: pd.DataFrame, current: Optional[str]) -> Optional[str]:
    """Identifica la columna real del conductor sin confundirla con costos.

    REGLA DE NEGOCIO DEL ARCHIVO TRAYECTOS TODOS:
    La base puede traer varias columnas con nombres parecidos a CONDUCTO/CONDUCTOR.
    Para este informe se debe tomar como nombre real del conductor la columna W
    del Excel, que equivale a la posición 22 en pandas.

    Esta prioridad evita que el dashboard use por error V.CONDUCT, VALOR CONDUCTOR
    o cualquier columna monetaria como nombre del conductor.
    """
    # 1) Prioridad absoluta: columna W del Excel = índice 22 en pandas.
    if len(df.columns) > 22:
        col_w = df.columns[22]
        if col_w in df.columns and not is_bad_conductor_column_name(col_w):
            # Si la columna W no es principalmente numérica, se asume como nombre.
            if numeric_ratio(df[col_w]) < 0.80:
                return col_w

    # 2) Coincidencia exacta fuerte por nombre normalizado.
    norm_map = {normalize_col(c): c for c in df.columns}
    for name in STRICT_CONDUCTOR_NAMES:
        if name in norm_map and not is_bad_conductor_column_name(norm_map[name]):
            if numeric_ratio(df[norm_map[name]]) < 0.80:
                return norm_map[name]

    # 3) Si el actual no parece financiero ni numérico, se acepta.
    if current and current in df.columns:
        if not is_bad_conductor_column_name(current) and numeric_ratio(df[current]) < 0.80:
            return current

    # 4) Búsqueda flexible: columnas que mencionen conductor, excluyendo valores/costos.
    candidates = []
    for c in df.columns:
        n = normalize_col(c)
        if "CONDUCT" in n and not is_bad_conductor_column_name(c):
            ratio = numeric_ratio(df[c])
            score = 100 if n in STRICT_CONDUCTOR_NAMES else 50
            score -= 40 if ratio > 0.80 else 0
            # Preferimos texto con más valores diferentes.
            unique_count = df[c].astype(str).str.strip().replace("", np.nan).nunique(dropna=True)
            score += min(unique_count, 100) / 100
            candidates.append((score, c))
    if candidates:
        candidates.sort(reverse=True, key=lambda x: x[0])
        return candidates[0][1]

    return None


def money_to_number(s: pd.Series) -> pd.Series:
    """Convierte valores monetarios a número.

    Regla principal para Colombia: punto = miles, coma = decimal.
    También soporta números ya numéricos y textos simples.
    """
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0)

    def parse_one(v):
        if pd.isna(v):
            return 0.0
        txt = str(v).strip()
        if txt == "":
            return 0.0
        txt = txt.replace("COP", "").replace("$", "").replace(" ", "")
        txt = re.sub(r"[^0-9,.-]", "", txt)
        if txt in ["", "-", ",", "."]:
            return 0.0
        # Si tiene coma, se interpreta como formato colombiano: 1.234.567,89
        if "," in txt:
            txt = txt.replace(".", "").replace(",", ".")
        else:
            # Si tiene más de un punto, son separadores de miles: 1.234.567
            if txt.count(".") > 1:
                txt = txt.replace(".", "")
            # Si tiene un punto y exactamente 3 dígitos después, probablemente es miles: 123.456
            elif txt.count(".") == 1:
                left, right = txt.split(".")
                if len(right) == 3 and len(left) <= 3:
                    txt = left + right
        try:
            return float(txt)
        except Exception:
            return 0.0

    return s.apply(parse_one).astype(float).fillna(0)


def fmt_cop(value: float) -> str:
    try:
        return "$ " + f"{float(value):,.0f}".replace(",", ".")
    except Exception:
        return "$ 0"


def fmt_pct(value: float) -> str:
    try:
        return f"{float(value)*100:,.2f} %".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00 %"


def fmt_num(value: float) -> str:
    try:
        return f"{float(value):,.0f}".replace(",", ".")
    except Exception:
        return "0"




def safe_key(*parts: object) -> str:
    raw = "_".join(str(p) for p in parts)
    raw = re.sub(r"[^A-Za-z0-9_]+", "_", raw)[:80]
    digest = hashlib.md5("|".join(str(p) for p in parts).encode("utf-8")).hexdigest()[:8]
    return f"{raw}_{digest}"



_PLOT_COUNTER = itertools.count(1)

def unique_plot_key(*parts: object) -> str:
    """Genera keys realmente únicos para evitar StreamlitDuplicateElementKey/Id.
    Streamlit puede repetir IDs si dos gráficos tienen la misma estructura; por eso
    agregamos contador y un sufijo corto uuid en cada renderizado.
    """
    return safe_key(*parts, next(_PLOT_COUNTER), uuid.uuid4().hex[:8])


def value_label(value: float, metric: str) -> str:
    if metric in ["Facturación", "Costos", "Margen", "Variación $", "Valor Actual", "Valor Anterior"] or "__V_CLIENTE__" in metric or "__MARGEN__" in metric:
        return fmt_cop(value)
    if metric in ["Rentabilidad", "Participación", "Variación %", "Crecimiento"]:
        return fmt_pct(value)
    return fmt_num(value)


def variation_arrow(value: float) -> str:
    if not np.isfinite(value):
        return "—"
    if value > 0:
        return "▲ " + fmt_pct(value)
    if value < 0:
        return "▼ " + fmt_pct(value)
    return "■ 0,00 %"


def make_kpi(title: str, value: str, note: str = "") -> None:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">{title}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def get_excel_sheets(file_bytes: bytes, filename: str) -> List[str]:
    ext = filename.lower().split(".")[-1]
    engine = "xlrd" if ext == "xls" else "openpyxl"
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
    return xls.sheet_names


@st.cache_data(show_spinner=False)
def load_file(file_bytes: bytes, filename: str, sheet_name: Optional[str]) -> pd.DataFrame:
    lower = filename.lower()
    if lower.endswith(".csv"):
        for enc in ["utf-8", "latin1", "cp1252"]:
            try:
                return pd.read_csv(io.BytesIO(file_bytes), sep=None, engine="python", encoding=enc)
            except Exception:
                continue
        return pd.read_csv(io.BytesIO(file_bytes), encoding="latin1")

    ext = lower.split(".")[-1]
    engine = "xlrd" if ext == "xls" else "openpyxl"
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, engine=engine)


@st.cache_data(show_spinner=False)
def prepare_data(df_raw: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Optional[str]]]:
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").copy()
    colmap = {k: first_existing(df, v) for k, v in COLUMN_CANDIDATES.items()}
    # Corrección crítica: no permitir que VALOR CONDUCTOR / V.CONDUCT sea tomado
    # como nombre del conductor.
    colmap["conductor"] = resolve_conductor_column(df, colmap.get("conductor"))

    fecha = colmap.get("fecha")
    if fecha:
        df["__FECHA__"] = pd.to_datetime(df[fecha], errors="coerce", dayfirst=True)
    else:
        df["__FECHA__"] = pd.NaT

    if df["__FECHA__"].isna().all():
        df["__ANIO__"] = "Sin fecha"
        df["__MES_NUM__"] = 0
        df["__MES__"] = "Sin fecha"
        df["__PERIODO_MES__"] = "Sin fecha"
        df["__BIMESTRE__"] = "Sin fecha"
        df["__TRIMESTRE__"] = "Sin fecha"
        df["__SEMESTRE__"] = "Sin fecha"
    else:
        df["__ANIO__"] = df["__FECHA__"].dt.year.astype("Int64").astype(str).replace("<NA>", "Sin fecha")
        df["__MES_NUM__"] = df["__FECHA__"].dt.month.fillna(0).astype(int)
        df["__MES__"] = df["__MES_NUM__"].map(MESES).fillna("Sin fecha")
        df["__PERIODO_MES__"] = df["__ANIO__"] + "-" + df["__MES_NUM__"].astype(str).str.zfill(2)
        df["__BIMESTRE__"] = df["__ANIO__"] + "-B" + (((df["__MES_NUM__"] - 1) // 2) + 1).astype(str)
        df["__TRIMESTRE__"] = df["__ANIO__"] + "-T" + (((df["__MES_NUM__"] - 1) // 3) + 1).astype(str)
        df["__SEMESTRE__"] = df["__ANIO__"] + "-S" + (((df["__MES_NUM__"] - 1) // 6) + 1).astype(str)
        df.loc[df["__MES_NUM__"].eq(0), ["__PERIODO_MES__", "__BIMESTRE__", "__TRIMESTRE__", "__SEMESTRE__"]] = "Sin fecha"

    vc = colmap.get("v_cliente")
    vd = colmap.get("v_conduct")
    df["__V_CLIENTE__"] = money_to_number(df[vc]) if vc else 0.0
    df["__V_CONDUCT__"] = money_to_number(df[vd]) if vd else 0.0
    df["__MARGEN__"] = df["__V_CLIENTE__"] - df["__V_CONDUCT__"]
    df["__RENTABILIDAD__"] = np.where(df["__V_CLIENTE__"].ne(0), df["__MARGEN__"] / df["__V_CLIENTE__"], 0)
    df["__SERVICIOS__"] = 1

    for key in ["cliente", "coordinador", "usuario", "estado_op", "tray_prop", "linea_neg", "tipo_negocio", "placa", "conductor", "origen", "destino", "tipo_vehiculo"]:
        col = colmap.get(key)
        out = f"__{key.upper()}__"
        if col:
            df[out] = df[col].fillna("Sin dato").astype(str).str.strip()
            df[out] = df[out].replace(["", "nan", "None", "NaN"], "Sin dato")
        else:
            df[out] = "Sin dato"
        if key == "tray_prop":
            df[out] = df[out].apply(classify_tray_prop)

    estado = colmap.get("estado_op")
    if estado:
        norm_estado = df[estado].apply(normalize_text)
        df = df[~norm_estado.eq("ANULADO")].copy()

    return df, colmap


def classify_tray_prop(value: object) -> str:
    """Clasifica TRAY. PROP para evitar ejes 0/1 y hacer lectura gerencial.

    Convención usada cuando la base trae binario:
    1 = FLOTA PROPIA, 0 = TERCEROS.
    Si el valor ya viene en texto, respeta propio/tercero.
    """
    txt = normalize_text(value)
    if txt in ["", "SIN DATO", "NAN", "NONE"]:
        return "Sin dato"
    if txt in ["1", "1.0", "SI", "S", "TRUE", "VERDADERO"]:
        return "FLOTA PROPIA"
    if txt in ["0", "0.0", "NO", "N", "FALSE", "FALSO"]:
        return "TERCEROS"
    if "TERCER" in txt:
        return "TERCEROS"
    if "PROPI" in txt or "FLOTA" in txt:
        return "FLOTA PROPIA"
    return str(value).strip() if str(value).strip() else "Sin dato"



def period_col(periodo: str) -> str:
    return {
        "Mensual": "__PERIODO_MES__",
        "Bimestral": "__BIMESTRE__",
        "Trimestral": "__TRIMESTRE__",
        "Semestral": "__SEMESTRE__",
        "Anual": "__ANIO__",
    }.get(periodo, "__PERIODO_MES__")


def aggregate(df: pd.DataFrame, group_cols: List[str]) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=group_cols + ["Servicios", "Facturación", "Costos", "Margen", "Rentabilidad"])
    g = df.groupby(group_cols, dropna=False, as_index=False).agg(
        Servicios=("__SERVICIOS__", "sum"),
        Facturación=("__V_CLIENTE__", "sum"),
        Costos=("__V_CONDUCT__", "sum"),
        Margen=("__MARGEN__", "sum"),
    )
    g["Rentabilidad"] = np.where(g["Facturación"].ne(0), g["Margen"] / g["Facturación"], 0)
    return g


def apply_filter(df: pd.DataFrame, label: str, col: str, key: str, max_opts: int = 10000) -> Tuple[pd.DataFrame, List[str]]:
    """Filtro encadenado. Las opciones se recalculan con la base ya filtrada y limpia selecciones inválidas."""
    if col not in df.columns:
        return df, []
    vals = df[col].dropna().astype(str).unique().tolist()
    if col == "__MES__":
        vals = [m for m in MESES_ORDEN if m in vals]
    else:
        vals = sorted(vals)
    if len(vals) > max_opts:
        vals = vals[:max_opts]

    prev = st.session_state.get(key, [])
    valid_prev = [v for v in prev if v in vals]
    if prev != valid_prev:
        st.session_state[key] = valid_prev

    selected = st.sidebar.multiselect(label, vals, default=valid_prev, key=key)
    if selected:
        return df[df[col].astype(str).isin(selected)].copy(), selected
    return df, []




def apply_selected_filters_for_comparison(df_source: pd.DataFrame, include_month: bool = False) -> pd.DataFrame:
    """Reconstruye una base comparativa con los filtros seleccionados.

    Se usa especialmente para variaciones vs periodo anterior.
    Cuando el usuario selecciona un mes específico, el KPI usa ese mes,
    pero la comparación necesita consultar el mes anterior con los mismos filtros.
    """
    d = df_source.copy()
    filter_map = [
        ("flt_anio", "__ANIO__"),
        ("flt_mes_revisar", "__MES__"),
        ("flt_cliente", "__CLIENTE__"),
        ("flt_coordinador", "__COORDINADOR__"),
        ("flt_usuario", "__USUARIO__"),
        ("flt_linea", "__LINEA_NEG__"),
        ("flt_tipo_negocio", "__TIPO_NEGOCIO__"),
        ("flt_tray_prop", "__TRAY_PROP__"),
        ("flt_estado", "__ESTADO_OP__"),
        ("flt_tipo_vehiculo", "__TIPO_VEHICULO__"),
        ("flt_placa", "__PLACA__"),
        ("flt_conductor", "__CONDUCTOR__"),
        ("flt_origen", "__ORIGEN__"),
        ("flt_destino", "__DESTINO__"),
    ]
    for key_name, col_name in filter_map:
        if col_name not in d.columns:
            continue
        val = st.session_state.get(key_name, [])
        if key_name == "flt_mes_revisar":
            if not include_month and val != "Todos":
                continue
            if val and val != "Todos":
                d = d[d[col_name].astype(str).eq(str(val))].copy()
        else:
            if isinstance(val, list) and val:
                d = d[d[col_name].astype(str).isin([str(x) for x in val])].copy()
    return d

def chart_bar(data: pd.DataFrame, x: str, y: str, title: str, top: int = 20, key: Optional[str] = None, ascending: bool = False) -> None:
    if data.empty or x not in data.columns or y not in data.columns:
        st.info("Sin datos para graficar.")
        return

    d = data.sort_values(y, ascending=ascending).head(top).copy()
    d[x] = d[x].astype(str).replace(["", "nan", "None", "NaN"], "Sin dato")
    d["__LABEL__"] = d[y].apply(lambda v: value_label(v, y))
    key = unique_plot_key("bar", key or "auto", title, x, y, top, "asc" if ascending else "desc", len(d))

    # Para nombres largos, como clientes o conductores, se usa barra horizontal.
    long_labels = d[x].astype(str).str.len().mean() > 12 if not d.empty else False
    horizontal = x in ["__CONDUCTOR__", "__CLIENTE__", "__ORIGEN__", "__DESTINO__"] or long_labels

    if PLOTLY_OK:
        if horizontal:
            # Orden visual correcto: el primer registro queda arriba.
            plot_df = d.iloc[::-1].copy()
            fig = px.bar(plot_df, x=y, y=x, orientation="h", text="__LABEL__", title=title)
            fig.update_yaxes(title_text="", tickfont=dict(size=11, color="#000000"))
            fig.update_xaxes(title_text=y)
            hovertemplate = f"%{{y}}<br>{y}: %{{text}}<extra></extra>"
        else:
            fig = px.bar(d, x=x, y=y, text="__LABEL__", title=title)
            fig.update_xaxes(title_text=x, tickfont=dict(size=11, color="#000000"))
            fig.update_yaxes(title_text=y)
            hovertemplate = f"%{{x}}<br>{y}: %{{text}}<extra></extra>"

        fig.update_layout(
            height=470 if horizontal else 430,
            plot_bgcolor="white",
            paper_bgcolor="white",
            title_font_size=18,
            margin=dict(l=20, r=20, t=70, b=40),
        )
        fig.update_traces(
            marker_color="#0B4F78",
            textposition="outside",
            cliponaxis=False,
            hovertemplate=hovertemplate,
        )
        fig = apply_dark_plotly_theme(fig)

        axis_to_format = fig.update_xaxes if horizontal else fig.update_yaxes
        if y in ["Facturación", "Costos", "Margen", "Variación $", "Valor Actual", "Valor Anterior"]:
            axis_to_format(title_text=y, tickprefix="$ ", tickformat=",.0f")
        elif y == "Rentabilidad" or "%" in y:
            axis_to_format(title_text=y, tickformat=".1%")
        else:
            axis_to_format(title_text=y, tickformat=",.0f")

        st.plotly_chart(fig, use_container_width=True, key=key)
    else:
        st.subheader(title)
        st.bar_chart(d.set_index(x)[y])


def chart_line(data: pd.DataFrame, x: str, y: str, title: str, color: Optional[str] = None) -> None:
    if data.empty or x not in data.columns or y not in data.columns:
        st.info("Sin datos para graficar.")
        return
    d = data.sort_values(x).copy()
    if PLOTLY_OK:
        fig = px.line(d, x=x, y=y, color=color, markers=True, title=title)
        fig.update_layout(height=430, plot_bgcolor="white", paper_bgcolor="white", title_font_size=18)
        fig = apply_dark_plotly_theme(fig)
        st.plotly_chart(fig, use_container_width=True, key=unique_plot_key("line", title, x, y, color or ""))
    else:
        st.subheader(title)
        if color and color in d.columns:
            pivot = d.pivot_table(index=x, columns=color, values=y, aggfunc="sum", fill_value=0)
            st.line_chart(pivot)
        else:
            st.line_chart(d.set_index(x)[y])


def chart_heatmap(data: pd.DataFrame, index: str, columns: str, values: str, title: str, aggfunc="sum", top_rows: int = 25) -> None:
    if data.empty or index not in data.columns or columns not in data.columns or values not in data.columns:
        st.info("Sin datos suficientes para este mapa de calor.")
        return
    base = data.copy()
    top_idx = base.groupby(index)[values].sum().sort_values(ascending=False).head(top_rows).index
    base = base[base[index].isin(top_idx)]
    pivot = pd.pivot_table(base, index=index, columns=columns, values=values, aggfunc=aggfunc, fill_value=0)
    pivot = pivot.sort_index(axis=1)
    if PLOTLY_OK:
        fig = px.imshow(
            pivot,
            aspect="auto",
            text_auto=".2s",
            color_continuous_scale=["#DC2626", "#FBBF24", "#16A34A"],
            title=title,
        )
        fig.update_layout(height=max(420, min(900, 90 + 26 * len(pivot))), title_font_size=18)
        fig = apply_dark_plotly_theme(fig)
        st.plotly_chart(fig, use_container_width=True, key=unique_plot_key("heatmap", title, index, columns, values))
    else:
        st.subheader(title)
        st.dataframe(pivot.style.background_gradient(cmap="RdYlGn", axis=None), use_container_width=True, height=520)


def show_table(df: pd.DataFrame, title: str, height: int = 420) -> None:
    st.subheader(title)
    if df.empty:
        st.info("Sin datos para mostrar.")
        return
    show = df.copy()
    for col in show.columns:
        if col in ["Facturación", "Costos", "Margen", "Variación $", "Valor Actual", "Valor Anterior"]:
            show[col] = show[col].apply(fmt_cop)
        elif col in ["Rentabilidad", "Participación", "Variación %", "Crecimiento", "Var Facturación %", "Var Margen %", "Var Servicios %", "OTIF Cierre", "% Cerrado", "% Pendiente"]:
            show[col] = show[col].apply(fmt_pct)
    if AGGRID_OK:
        gb = GridOptionsBuilder.from_dataframe(show)
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=20)
        gb.configure_default_column(filter=True, sortable=True, resizable=True)
        AgGrid(show, gridOptions=gb.build(), height=height, theme="streamlit", fit_columns_on_grid_load=False)
    else:
        st.dataframe(show, use_container_width=True, height=height)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Datos")
    return output.getvalue()


def ordered_periods(values: List[str], periodo: str = "Mensual") -> List[str]:
    def period_sort_key(v: str):
        v = str(v)
        if v == "Sin fecha":
            return (9999, 99)
        m = re.match(r"^(\d{4})-(\d{2})$", v)
        if m:
            return (int(m.group(1)), int(m.group(2)))
        m = re.match(r"^(\d{4})-B(\d+)$", v)
        if m:
            return (int(m.group(1)), int(m.group(2)) * 2)
        m = re.match(r"^(\d{4})-T(\d+)$", v)
        if m:
            return (int(m.group(1)), int(m.group(2)) * 3)
        m = re.match(r"^(\d{4})-S(\d+)$", v)
        if m:
            return (int(m.group(1)), int(m.group(2)) * 6)
        if re.match(r"^\d{4}$", v):
            return (int(v), 12)
        return (9998, v)
    return sorted([str(x) for x in values], key=period_sort_key)



def variation_table(df: pd.DataFrame, group_col: str, pcol: str, value_col: str) -> pd.DataFrame:
    periods = ordered_periods([p for p in df[pcol].dropna().astype(str).unique().tolist() if p != "Sin fecha"])
    if len(periods) < 2:
        return pd.DataFrame()
    prev_p, curr_p = periods[-2], periods[-1]
    piv = df[df[pcol].astype(str).isin([prev_p, curr_p])].pivot_table(index=group_col, columns=pcol, values=value_col, aggfunc="sum", fill_value=0).reset_index()
    if prev_p not in piv.columns:
        piv[prev_p] = 0
    if curr_p not in piv.columns:
        piv[curr_p] = 0
    piv["Valor Anterior"] = piv[prev_p]
    piv["Valor Actual"] = piv[curr_p]
    piv["Variación $"] = piv["Valor Actual"] - piv["Valor Anterior"]
    piv["Variación %"] = np.where(piv["Valor Anterior"].ne(0), piv["Variación $"] / piv["Valor Anterior"], np.where(piv["Valor Actual"].ne(0), 1, 0))
    return piv[[group_col, "Valor Anterior", "Valor Actual", "Variación $", "Variación %"]].sort_values("Variación %")




def period_summary(df: pd.DataFrame, pcol: str) -> pd.DataFrame:
    fin = aggregate(df, [pcol]).copy()
    if not fin.empty:
        order = ordered_periods(fin[pcol].dropna().astype(str).unique().tolist())
        fin["__ORDER__"] = pd.Categorical(fin[pcol].astype(str), categories=order, ordered=True)
        fin = fin.sort_values("__ORDER__").drop(columns=["__ORDER__"])
    if fin.empty:
        return fin
    fin["Var Facturación %"] = fin["Facturación"].pct_change().replace([np.inf, -np.inf], np.nan).fillna(0)
    fin["Var Margen %"] = fin["Margen"].pct_change().replace([np.inf, -np.inf], np.nan).fillna(0)
    fin["Var Servicios %"] = fin["Servicios"].pct_change().replace([np.inf, -np.inf], np.nan).fillna(0)
    fin["Variación Facturación"] = fin["Var Facturación %"].apply(variation_arrow)
    fin["Variación Margen"] = fin["Var Margen %"].apply(variation_arrow)
    fin["Variación Servicios"] = fin["Var Servicios %"].apply(variation_arrow)
    return fin


def chart_financial_combo(fin: pd.DataFrame, pcol: str, title: str, key: str) -> None:
    if fin.empty:
        st.info("Sin datos financieros para graficar.")
        return
    if PLOTLY_OK:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=fin[pcol], y=fin["Facturación"], name="Facturación",
            text=fin["Facturación"].apply(fmt_cop), textposition="outside", marker_color="#0B4F78",
            hovertemplate="Periodo: %{x}<br>Facturación: %{text}<extra></extra>"
        ))
        fig.add_trace(go.Scatter(
            x=fin[pcol], y=fin["Margen"], name="Margen", mode="lines+markers+text",
            text=fin["Variación Facturación"], textposition="top center", yaxis="y2", line=dict(color="#60A5FA", width=3),
            hovertemplate="Periodo: %{x}<br>Margen: %{customdata}<br>Variación facturación vs anterior: %{text}<extra></extra>",
            customdata=fin["Margen"].apply(fmt_cop)
        ))
        fig.update_layout(
            title=title, height=470, plot_bgcolor="white", paper_bgcolor="white",
            yaxis=dict(title="Facturación", tickprefix="$ ", tickformat=",.0f"),
            yaxis2=dict(title="Margen", overlaying="y", side="right", tickprefix="$ ", tickformat=",.0f"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(t=80, b=40, l=40, r=40),
        )
        fig = apply_dark_plotly_theme(fig)
        st.plotly_chart(fig, use_container_width=True, key=key)
    else:
        st.subheader(title)
        st.line_chart(fin.set_index(pcol)[["Facturación", "Margen"]])




def is_cierre_total_estado(value: object) -> bool:
    """Regla oficial de cierre mensual.

    ESTADO OP = CUMPLIDO OPERATIVO significa que la remesa/servicio quedó cerrado
    totalmente a corte de fin de mes. ANULADO ya fue excluido en prepare_data.
    """
    return normalize_text(value) == "CUMPLIDO OPERATIVO"


def closure_summary(df_base: pd.DataFrame) -> Dict[str, float]:
    """KPIs de cierre operativo/OTIF.

    Denominador: total de servicios no anulados dentro del filtro.
    Numerador: servicios con ESTADO OP = CUMPLIDO OPERATIVO.
    Pendiente: todo estado diferente de CUMPLIDO OPERATIVO.
    """
    total = int(len(df_base))
    cerrados = int(df_base["__ESTADO_OP__"].apply(is_cierre_total_estado).sum()) if "__ESTADO_OP__" in df_base.columns else 0
    pendientes = max(total - cerrados, 0)
    otif = cerrados / total if total else 0.0
    corte = df_base["__FECHA__"].max() if "__FECHA__" in df_base.columns and not df_base["__FECHA__"].isna().all() else pd.NaT
    return {"total": total, "cerrados": cerrados, "pendientes": pendientes, "otif": otif, "corte": corte}


def closure_person_table(df_base: pd.DataFrame, person_col: str = "__USUARIO__") -> pd.DataFrame:
    """Tabla tipo Excel: Quien creó x Estado OP + Total + OTIF."""
    if df_base.empty or person_col not in df_base.columns or "__ESTADO_OP__" not in df_base.columns:
        return pd.DataFrame()
    d = df_base.copy()
    d[person_col] = d[person_col].fillna("Sin dato").astype(str).replace("", "Sin dato")
    d["__ESTADO_OP__"] = d["__ESTADO_OP__"].fillna("Sin dato").astype(str).replace("", "Sin dato")
    pivot = pd.crosstab(d[person_col], d["__ESTADO_OP__"])
    pivot["Total general"] = pivot.sum(axis=1)
    pivot["Cerradas"] = d.groupby(person_col)["__ESTADO_OP__"].apply(lambda s: s.apply(is_cierre_total_estado).sum())
    pivot["Pendientes cierre"] = pivot["Total general"] - pivot["Cerradas"]
    pivot["OTIF Cierre"] = np.where(pivot["Total general"].ne(0), pivot["Cerradas"] / pivot["Total general"], 0)
    pivot = pivot.reset_index().rename(columns={person_col: "Quien creó"})
    estado_cols = [c for c in pivot.columns if c not in ["Quien creó", "Total general", "Cerradas", "Pendientes cierre", "OTIF Cierre"]]
    ordered = ["Quien creó"] + sorted(estado_cols) + ["Total general", "Cerradas", "Pendientes cierre", "OTIF Cierre"]
    return pivot[ordered].sort_values(["OTIF Cierre", "Total general"], ascending=[False, False])

def closure_matrix_by_dimension(df_base: pd.DataFrame, dim_col: str, dim_label: str) -> pd.DataFrame:
    """Matriz tipo tabla dinámica: dimensión vs ESTADO OP + métricas de cierre.

    Regla de negocio:
    - Denominador: servicios válidos del filtro actual, sin ANULADO.
    - Numerador: ESTADO OP = CUMPLIDO OPERATIVO.
    - Pendiente: todo estado diferente de CUMPLIDO OPERATIVO.
    """
    if df_base.empty or dim_col not in df_base.columns or "__ESTADO_OP__" not in df_base.columns:
        return pd.DataFrame()
    d = df_base.copy()
    d[dim_col] = d[dim_col].fillna("Sin dato").astype(str).replace("", "Sin dato")
    d["__ESTADO_OP__"] = d["__ESTADO_OP__"].fillna("Sin dato").astype(str).replace("", "Sin dato")

    pivot = pd.crosstab(d[dim_col], d["__ESTADO_OP__"])
    pivot["Total general"] = pivot.sum(axis=1)
    pivot["Cerradas"] = d.groupby(dim_col)["__ESTADO_OP__"].apply(lambda s: s.apply(is_cierre_total_estado).sum())
    pivot["Pendientes cierre"] = pivot["Total general"] - pivot["Cerradas"]
    pivot["OTIF Cierre"] = np.where(pivot["Total general"].ne(0), pivot["Cerradas"] / pivot["Total general"], 0)
    pivot["% Pendiente"] = np.where(pivot["Total general"].ne(0), pivot["Pendientes cierre"] / pivot["Total general"], 0)

    pivot = pivot.reset_index().rename(columns={dim_col: dim_label})
    estado_cols = [c for c in pivot.columns if c not in [dim_label, "Total general", "Cerradas", "Pendientes cierre", "OTIF Cierre", "% Pendiente"]]
    ordered = [dim_label] + sorted(estado_cols) + ["Total general", "Cerradas", "Pendientes cierre", "OTIF Cierre", "% Pendiente"]
    return pivot[ordered].sort_values(["Pendientes cierre", "Total general"], ascending=[False, False])


def closure_matrix_client_person(df_base: pd.DataFrame, person_col: str = "__USUARIO__") -> pd.DataFrame:
    """Matriz cliente + responsable vs ESTADO OP para identificar pendientes por cuenta y persona."""
    if df_base.empty or "__CLIENTE__" not in df_base.columns or person_col not in df_base.columns or "__ESTADO_OP__" not in df_base.columns:
        return pd.DataFrame()
    d = df_base.copy()
    d["Cliente"] = d["__CLIENTE__"].fillna("Sin dato").astype(str).replace("", "Sin dato")
    d["Quien creó"] = d[person_col].fillna("Sin dato").astype(str).replace("", "Sin dato")
    d["Estado OP"] = d["__ESTADO_OP__"].fillna("Sin dato").astype(str).replace("", "Sin dato")

    pivot = pd.crosstab([d["Cliente"], d["Quien creó"]], d["Estado OP"])
    pivot["Total general"] = pivot.sum(axis=1)
    cierre_group = d.groupby(["Cliente", "Quien creó"])["Estado OP"].apply(lambda s: s.apply(is_cierre_total_estado).sum())
    pivot["Cerradas"] = cierre_group
    pivot["Pendientes cierre"] = pivot["Total general"] - pivot["Cerradas"]
    pivot["OTIF Cierre"] = np.where(pivot["Total general"].ne(0), pivot["Cerradas"] / pivot["Total general"], 0)
    pivot["% Pendiente"] = np.where(pivot["Total general"].ne(0), pivot["Pendientes cierre"] / pivot["Total general"], 0)
    pivot = pivot.reset_index()
    estado_cols = [c for c in pivot.columns if c not in ["Cliente", "Quien creó", "Total general", "Cerradas", "Pendientes cierre", "OTIF Cierre", "% Pendiente"]]
    ordered = ["Cliente", "Quien creó"] + sorted(estado_cols) + ["Total general", "Cerradas", "Pendientes cierre", "OTIF Cierre", "% Pendiente"]
    return pivot[ordered].sort_values(["Pendientes cierre", "Total general"], ascending=[False, False])


def add_total_row_closure(tbl: pd.DataFrame, label_col: str) -> pd.DataFrame:
    """Agrega fila Total general a matrices de cierre."""
    if tbl.empty or label_col not in tbl.columns:
        return tbl
    total_row = {col: 0 for col in tbl.columns}
    total_row[label_col] = "Total general"
    for col in tbl.columns:
        if col == label_col or col in ["OTIF Cierre", "% Pendiente"]:
            continue
        total_row[col] = pd.to_numeric(tbl[col], errors="coerce").fillna(0).sum()
    total = total_row.get("Total general", 0)
    cerradas = total_row.get("Cerradas", 0)
    pendientes = total_row.get("Pendientes cierre", 0)
    total_row["OTIF Cierre"] = cerradas / total if total else 0
    total_row["% Pendiente"] = pendientes / total if total else 0
    return pd.concat([tbl, pd.DataFrame([total_row])], ignore_index=True)


def closure_by_period(df_base: pd.DataFrame, pcol: str) -> pd.DataFrame:
    if df_base.empty or pcol not in df_base.columns:
        return pd.DataFrame()
    d = df_base.copy()
    d["__CERRADO__"] = d["__ESTADO_OP__"].apply(is_cierre_total_estado).astype(int)
    g = d.groupby(pcol, as_index=False).agg(
        Servicios=("__SERVICIOS__", "sum"),
        Cerradas=("__CERRADO__", "sum"),
    )
    g["Pendientes cierre"] = g["Servicios"] - g["Cerradas"]
    g["OTIF Cierre"] = np.where(g["Servicios"].ne(0), g["Cerradas"] / g["Servicios"], 0)
    g["__ORD__"] = g[pcol].apply(lambda v: ordered_periods([v])[0] if str(v) != "Sin fecha" else "9999")
    return g.sort_values(pcol).drop(columns=["__ORD__"], errors="ignore")


def closure_status_matrix(df_base: pd.DataFrame, person_col: str = "__USUARIO__") -> pd.DataFrame:
    """Matriz pura para vista similar a tabla dinámica de Excel."""
    tbl = closure_person_table(df_base, person_col)
    if tbl.empty:
        return tbl
    total_row = {col: 0 for col in tbl.columns}
    total_row["Quien creó"] = "Total general"
    numeric_cols = [c for c in tbl.columns if c not in ["Quien creó", "OTIF Cierre"]]
    for c in numeric_cols:
        total_row[c] = pd.to_numeric(tbl[c], errors="coerce").fillna(0).sum()
    total_row["OTIF Cierre"] = total_row.get("Cerradas", 0) / total_row.get("Total general", 1) if total_row.get("Total general", 0) else 0
    return pd.concat([tbl, pd.DataFrame([total_row])], ignore_index=True)

def executive_summary(df: pd.DataFrame, pcol: str) -> str:
    fact = df["__V_CLIENTE__"].sum()
    cost = df["__V_CONDUCT__"].sum()
    margin = df["__MARGEN__"].sum()
    rent = margin / fact if fact else 0
    clientes = df["__CLIENTE__"].nunique()
    servicios = len(df)
    by_cliente = aggregate(df, ["__CLIENTE__"]).sort_values("Facturación", ascending=False)
    top_client = by_cliente.iloc[0]["__CLIENTE__"] if not by_cliente.empty else "Sin dato"
    top_part = by_cliente.iloc[0]["Facturación"] / fact if fact and not by_cliente.empty else 0
    periods = ordered_periods([p for p in df[pcol].dropna().astype(str).unique() if p != "Sin fecha"])
    trend = "No hay periodos suficientes para comparar tendencia."
    if len(periods) >= 2:
        prev_p, curr_p = periods[-2], periods[-1]
        a = df[df[pcol].astype(str).eq(curr_p)]["__V_CLIENTE__"].sum()
        b = df[df[pcol].astype(str).eq(prev_p)]["__V_CLIENTE__"].sum()
        var = (a - b) / b if b else 0
        direction = "crecimiento" if var >= 0 else "disminución"
        trend = f"Frente al periodo anterior, la facturación presenta {direction} de {fmt_pct(var)}."
    return (
        f"Con los filtros seleccionados se analizan {fmt_num(servicios)} servicios, {fmt_num(clientes)} clientes, "
        f"facturación de {fmt_cop(fact)}, costos de {fmt_cop(cost)}, margen de {fmt_cop(margin)} y rentabilidad global de {fmt_pct(rent)}. "
        f"El cliente con mayor participación es {top_client}, con {fmt_pct(top_part)} de la facturación. {trend}"
    )


def find_city_coords(text: object) -> Optional[Tuple[float, float]]:
    norm = normalize_text(text)
    if not norm:
        return None
    for city, coord in COORD_CIUDADES.items():
        if normalize_text(city) in norm:
            return coord
    return None

# =========================
# APP
# =========================
st.markdown("<div class='main-title'>📊 Informe Gerencial Ejecutivo Operativo y Financiero</div>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>Centro de Control Empresarial para análisis operativo, financiero, rentabilidad, clientes, responsables y alertas ejecutivas.</div>", unsafe_allow_html=True)

if not PLOTLY_OK:
    st.markdown("<div class='warn-box'>Plotly no está instalado. La app funciona con gráficos básicos. Para visuales premium agrega <b>plotly==5.24.1</b> al archivo requirements.txt y reinicia la app.</div>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## 📁 Cargar base actualizada")
    uploaded = st.file_uploader(
        "Sube TRAYECTOS TODOS / CONSOLIDADO",
        type=["xlsx", "xlsm", "xls", "csv"],
        help="Carga aquí la base actualizada cada vez que necesites presentar el informe.",
    )

if uploaded is None:
    st.markdown("<div class='info-box'>Carga un archivo Excel o CSV desde la barra lateral para iniciar el informe.</div>", unsafe_allow_html=True)
    st.stop()

file_bytes = uploaded.getvalue()
sheet_name = None
if not uploaded.name.lower().endswith(".csv"):
    try:
        sheets = get_excel_sheets(file_bytes, uploaded.name)
        with st.sidebar:
            sheet_name = st.selectbox("Hoja a analizar", sheets, index=0)
    except ImportError as e:
        st.error("No fue posible leer el archivo .xls porque falta xlrd. Agrega xlrd==2.0.1 a requirements.txt, reinicia Streamlit o guarda el archivo como .xlsx.")
        st.exception(e)
        st.stop()
    except Exception as e:
        st.error("No fue posible leer las hojas. Si el archivo es .xls, intenta guardarlo como .xlsx y vuelve a cargarlo.")
        st.exception(e)
        st.stop()

try:
    with st.spinner("Cargando y preparando la base..."):
        raw_df = load_file(file_bytes, uploaded.name, sheet_name)
        df_base, colmap = prepare_data(raw_df)
except ImportError as e:
    st.error("Falta una librería para leer el archivo. Para .xls instala xlrd==2.0.1. Para .xlsx instala openpyxl==3.1.5.")
    st.exception(e)
    st.stop()
except Exception as e:
    st.error("No fue posible cargar la base. Revisa que no esté protegida, vacía o dañada.")
    st.exception(e)
    st.stop()

with st.sidebar:
    st.markdown("---")
    st.markdown("## 🎛️ Filtros")

    df = df_base.copy()
    if st.button("🔄 Limpiar filtros", use_container_width=True):
        for k in list(st.session_state.keys()):
            if k.startswith("flt_"):
                del st.session_state[k]
        st.rerun()

    # Primero se filtra por año; luego se muestran únicamente los meses disponibles
    # para ese año. Esto hace que el filtro de mes sí afecte todos los demás filtros,
    # KPIs, tablas y gráficos.
    df, anios_sel = apply_filter(df, "Año", "__ANIO__", "flt_anio")

    meses_disponibles = [m for m in MESES_ORDEN if m in df["__MES__"].dropna().astype(str).unique().tolist()]
    if not meses_disponibles:
        meses_disponibles = ["Sin fecha"]
    mes_actual = st.session_state.get("flt_mes_revisar", "Todos")
    if mes_actual not in ["Todos"] + meses_disponibles:
        mes_actual = "Todos"
    mes_revisar = st.selectbox(
        "Mes a revisar",
        ["Todos"] + meses_disponibles,
        index=(["Todos"] + meses_disponibles).index(mes_actual),
        key="flt_mes_revisar",
        help="Selecciona el mes exacto que quieres presentar. Este filtro afecta todo el informe."
    )
    if mes_revisar != "Todos":
        df = df[df["__MES__"].astype(str).eq(mes_revisar)].copy()

    periodo = st.selectbox(
        "Periodo de análisis",
        ["Mensual", "Bimestral", "Trimestral", "Semestral", "Anual"],
        index=0,
        key="flt_periodo_analisis",
        help="Define cómo se agrupan las tendencias y mapas de calor."
    )
    pcol = period_col(periodo)

    df, _ = apply_filter(df, "Cliente", "__CLIENTE__", "flt_cliente")
    df, _ = apply_filter(df, "Coordinador", "__COORDINADOR__", "flt_coordinador")
    df, _ = apply_filter(df, "Usuario / Creado Por", "__USUARIO__", "flt_usuario")
    df, _ = apply_filter(df, "Línea de Negocio", "__LINEA_NEG__", "flt_linea")
    df, _ = apply_filter(df, "Tipo de Negocio", "__TIPO_NEGOCIO__", "flt_tipo_negocio")
    df, _ = apply_filter(df, "Trayecto Propio / Tercero", "__TRAY_PROP__", "flt_tray_prop")
    df, _ = apply_filter(df, "Estado Operativo", "__ESTADO_OP__", "flt_estado")
    df, _ = apply_filter(df, "Tipo Vehículo", "__TIPO_VEHICULO__", "flt_tipo_vehiculo")
    df, _ = apply_filter(df, "Placa", "__PLACA__", "flt_placa")
    df, _ = apply_filter(df, "Conductor", "__CONDUCTOR__", "flt_conductor")
    df, _ = apply_filter(df, "Origen", "__ORIGEN__", "flt_origen")
    df, _ = apply_filter(df, "Destino", "__DESTINO__", "flt_destino")
    st.markdown("---")
    st.caption(f"Registros cargados: {len(df_base):,}".replace(",", "."))
    st.caption(f"Registros filtrados: {len(df):,}".replace(",", "."))

# Base comparativa: conserva los mismos filtros, pero sin limitar al mes seleccionado.
# Así las variaciones vs periodo anterior siguen funcionando cuando se elige un mes específico.
df_compare = apply_selected_filters_for_comparison(df_base, include_month=False)

if df.empty:
    st.warning("No hay registros con los filtros seleccionados.")
    st.stop()

# Contexto visible del filtro principal para presentación gerencial
mes_visible = st.session_state.get("flt_mes_revisar", "Todos")
periodo_visible = st.session_state.get("flt_periodo_analisis", "Mensual")
st.markdown(
    f"""
    <div class='section-card' style='padding:12px 16px; margin-top:8px;'>
        <b>📅 Mes a revisar:</b> {mes_visible} &nbsp;&nbsp; | &nbsp;&nbsp;
        <b>📊 Agrupación:</b> {periodo_visible} &nbsp;&nbsp; | &nbsp;&nbsp;
        <b>Registros filtrados:</b> {len(df):,}
    </div>
    """.replace(",", "."),
    unsafe_allow_html=True,
)

# KPIs principales
fact = df["__V_CLIENTE__"].sum()
cost = df["__V_CONDUCT__"].sum()
margin = df["__MARGEN__"].sum()
rent = margin / fact if fact else 0
services = len(df)
clientes = df["__CLIENTE__"].replace("Sin dato", np.nan).nunique()
vehiculos = df["__PLACA__"].replace("Sin dato", np.nan).nunique()
conductores = df["__CONDUCTOR__"].replace("Sin dato", np.nan).nunique()

cols = st.columns(4)
with cols[0]: make_kpi("Producción Total", fmt_num(services), "Servicios")
with cols[1]: make_kpi("Facturación Total", fmt_cop(fact), "V.CLIENTE")
with cols[2]: make_kpi("Margen Total", fmt_cop(margin), "V.CLIENTE - V.CONDUCT")
with cols[3]: make_kpi("Rentabilidad Global", fmt_pct(rent), "Margen / Facturación")
cols = st.columns(4)
with cols[0]: make_kpi("Costos Totales", fmt_cop(cost), "V.CONDUCT")
with cols[1]: make_kpi("Clientes Activos", fmt_num(clientes), "Clientes con registros")
with cols[2]: make_kpi("Vehículos Activos", fmt_num(vehiculos), "Placas únicas")
with cols[3]: make_kpi("Conductores Activos", fmt_num(conductores), "Conductores únicos")

# Información de columnas detectadas
with st.expander("Columnas detectadas y reglas matemáticas aplicadas"):
    detect = pd.DataFrame([{"Campo estándar": k, "Columna encontrada": v or "No encontrada"} for k, v in colmap.items()])
    st.dataframe(detect, use_container_width=True)
    st.markdown("""
    **Reglas validadas:**  
    - Producción = conteo de registros no anulados.  
    - Facturación = suma de V.CLIENTE.  
    - Costos = suma de V.CONDUCT.  
    - Margen = Facturación - Costos.  
    - Rentabilidad % = Margen / Facturación.  
    - La rentabilidad agregada se calcula sobre totales, no como promedio simple de filas.  
    - ESTADO OP = ANULADO se excluye de KPIs, producción, facturación y rentabilidad.
    """)

# Tabs
tabs = st.tabs([
    "1. Resumen Gerencial",
    "2. Clientes",
    "3. Operación",
    "4. Financiero",
    "5. Rentabilidad",
    "6. Responsables",
    "7. Flota y Conductores",
    "8. Mapa Estratégico",
    "9. Alertas Ejecutivas",
    "10. Cierre / OTIF",
])

with tabs[0]:
    st.markdown("### Resumen Ejecutivo Automático")
    st.markdown(f"<div class='card'>{executive_summary(df_compare if not df_compare.empty else df, pcol)}</div>", unsafe_allow_html=True)
    g_period = period_summary(df_compare if not df_compare.empty else df, pcol)
    c1, c2 = st.columns(2)
    with c1:
        chart_financial_combo(g_period, pcol, f"Facturación y Margen - {periodo}", key=safe_key("combo_resumen", periodo, len(df)))
    with c2:
        chart_heatmap(df, "__CLIENTE__", pcol, "__V_CLIENTE__", "Mapa de Calor: Cliente vs Periodo - Facturación")
    c3, c4 = st.columns(2)
    with c3:
        chart_heatmap(df, "__CLIENTE__", pcol, "__MARGEN__", "Mapa de Calor: Cliente vs Periodo - Margen")
    with c4:
        tmp_r = aggregate(df, ["__CLIENTE__", pcol])
        chart_heatmap(tmp_r, "__CLIENTE__", pcol, "Rentabilidad", "Mapa de Calor: Cliente vs Periodo - Rentabilidad", aggfunc="mean")

with tabs[1]:
    clientes_tbl = aggregate(df, ["__CLIENTE__"]).sort_values("Facturación", ascending=False)
    total_fact = clientes_tbl["Facturación"].sum() or 1
    clientes_tbl["Participación"] = clientes_tbl["Facturación"] / total_fact
    c1, c2 = st.columns(2)
    with c1:
        chart_bar(clientes_tbl, "__CLIENTE__", "Facturación", "Top 10 Clientes por Facturación", 10)
    with c2:
        chart_bar(clientes_tbl[clientes_tbl["Facturación"] > 0].sort_values("Rentabilidad", ascending=False), "__CLIENTE__", "Rentabilidad", "Top 10 Clientes por Rentabilidad", 10)
    evolucion = aggregate(df, [pcol, "__CLIENTE__"])
    top_clients = clientes_tbl.head(8)["__CLIENTE__"].tolist()
    chart_line(evolucion[evolucion["__CLIENTE__"].isin(top_clients)], pcol, "Facturación", "Evolución de Facturación por Cliente", "__CLIENTE__")
    show_table(clientes_tbl.rename(columns={"__CLIENTE__": "Cliente"}), "Tabla Ejecutiva por Cliente")
    st.download_button("Descargar clientes en Excel", to_excel_bytes(clientes_tbl.rename(columns={"__CLIENTE__": "Cliente"})), "clientes_filtrados.xlsx")

with tabs[2]:
    c1, c2 = st.columns(2)
    with c1:
        estado_tbl = aggregate(df, ["__ESTADO_OP__"]).sort_values("Servicios", ascending=False)
        chart_bar(estado_tbl, "__ESTADO_OP__", "Servicios", "Pareto Estado Operativo", 20)
    with c2:
        tray_tbl = aggregate(df, ["__TRAY_PROP__"]).sort_values("Servicios", ascending=False)
        chart_bar(tray_tbl, "__TRAY_PROP__", "Servicios", "Producción Flota Propia / Terceros", 20)
    c3, c4 = st.columns(2)
    with c3:
        linea_tbl = aggregate(df, ["__LINEA_NEG__"]).sort_values("Facturación", ascending=False)
        chart_bar(linea_tbl, "__LINEA_NEG__", "Facturación", "Facturación por Línea de Negocio", 20)
    with c4:
        tipo_tbl = aggregate(df, ["__TIPO_NEGOCIO__"]).sort_values("Facturación", ascending=False)
        chart_bar(tipo_tbl, "__TIPO_NEGOCIO__", "Facturación", "Facturación por Tipo de Negocio", 20)
    c5, c6 = st.columns(2)
    with c5:
        origen_tbl = aggregate(df, ["__ORIGEN__"]).sort_values("Servicios", ascending=False)
        chart_bar(origen_tbl, "__ORIGEN__", "Servicios", "Top Orígenes", 20)
    with c6:
        destino_tbl = aggregate(df, ["__DESTINO__"]).sort_values("Servicios", ascending=False)
        chart_bar(destino_tbl, "__DESTINO__", "Servicios", "Top Destinos", 20)

with tabs[3]:
    fin = period_summary(df_compare if not df_compare.empty else df, pcol)
    chart_financial_combo(fin, pcol, f"Análisis Financiero {periodo} con Variación vs Periodo Anterior", key=safe_key("combo_financiero", periodo, len(df)))
    if PLOTLY_OK and not fin.empty:
        fig_var = px.bar(fin, x=pcol, y="Var Facturación %", text="Variación Facturación", title="Subidas y Bajadas de Facturación vs Periodo Anterior")
        fig_var.update_layout(height=390, plot_bgcolor="white", paper_bgcolor="white", yaxis=dict(title="Variación %", tickformat=".1%"))
        fig_var.update_traces(marker_color=np.where(fin["Var Facturación %"] >= 0, "#16A34A", "#DC2626"), textposition="outside", hovertemplate="Periodo: %{x}<br>Variación: %{text}<extra></extra>")
        fig_var = apply_dark_plotly_theme(fig_var)
        st.plotly_chart(fig_var, use_container_width=True, key=unique_plot_key("var_facturacion", periodo, len(df)))
    elif not fin.empty:
        st.bar_chart(fin.set_index(pcol)["Var Facturación %"])
    show_table(fin.rename(columns={pcol: "Periodo"}), "Tabla Financiera por Periodo")
    var_fact = variation_table(df_compare if not df_compare.empty else df, "__CLIENTE__", pcol, "__V_CLIENTE__")
    if not var_fact.empty:
        show_table(var_fact.rename(columns={"__CLIENTE__": "Cliente"}), "Variación de Facturación por Cliente")

with tabs[4]:
    st.markdown("### Análisis de Rentabilidad: eficiencia (%) vs aporte económico ($)")
    st.info(
        "La rentabilidad porcentual muestra eficiencia; el margen en pesos muestra la utilidad real. "
        "Un cliente pequeño puede tener alta rentabilidad %, pero un cliente grande puede aportar mucho más margen en COP."
    )

    rent_cliente = aggregate(df, ["__CLIENTE__"])
    rent_cliente = rent_cliente[rent_cliente["Facturación"] > 0].copy()

    if rent_cliente.empty:
        st.info("Sin datos de clientes para analizar rentabilidad.")
    else:
        total_clientes_rent = len(rent_cliente)
        top_n_rent = min(20, max(1, total_clientes_rent // 2)) if total_clientes_rent < 40 else 20

        clientes_mas_rentables = rent_cliente.sort_values("Rentabilidad", ascending=False).head(top_n_rent).copy()
        clientes_menos_rentables = rent_cliente.sort_values("Rentabilidad", ascending=True).head(top_n_rent).copy()
        clientes_mayor_margen = rent_cliente.sort_values("Margen", ascending=False).head(20).copy()
        clientes_mayor_facturacion = rent_cliente.sort_values("Facturación", ascending=False).head(20).copy()

        c1, c2 = st.columns(2)
        with c1:
            chart_bar(
                clientes_mas_rentables, "__CLIENTE__", "Rentabilidad",
                f"Top {top_n_rent} Más Rentables (%)", top_n_rent,
                key="rent_clientes_mas_porcentaje", ascending=False
            )
        with c2:
            chart_bar(
                clientes_menos_rentables, "__CLIENTE__", "Rentabilidad",
                f"Top {top_n_rent} Menor Rentabilidad (%)", top_n_rent,
                key="rent_clientes_menor_porcentaje", ascending=True
            )

        c3, c4 = st.columns(2)
        with c3:
            chart_bar(
                clientes_mayor_margen, "__CLIENTE__", "Margen",
                "Top 20 Mayor Margen ($ COP)", 20,
                key="rent_clientes_mayor_margen", ascending=False
            )
        with c4:
            chart_bar(
                clientes_mayor_facturacion, "__CLIENTE__", "Facturación",
                "Top 20 Mayor Facturación ($ COP)", 20,
                key="rent_clientes_mayor_facturacion", ascending=False
            )

        st.markdown("### Matriz Gerencial Cliente: Facturación, Margen y Rentabilidad")
        matriz = rent_cliente.copy()
        total_fact_matriz = matriz["Facturación"].sum() or 1
        total_margen_matriz = matriz["Margen"].sum() or 1
        mediana_fact = matriz["Facturación"].median() if not matriz.empty else 0
        mediana_rent = matriz["Rentabilidad"].median() if not matriz.empty else 0
        matriz["Participación Facturación"] = matriz["Facturación"] / total_fact_matriz
        matriz["Participación Margen"] = matriz["Margen"] / total_margen_matriz

        def clasificar_cliente(row):
            alta_fact = row["Facturación"] >= mediana_fact
            alta_rent = row["Rentabilidad"] >= mediana_rent
            if alta_fact and alta_rent:
                return "⭐ Estratégico: alta facturación y buena rentabilidad"
            if alta_fact and not alta_rent:
                return "⚠️ Alto volumen con rentabilidad baja: revisar tarifas/costos"
            if not alta_fact and alta_rent:
                return "📈 Oportunidad: rentable, pero con bajo volumen"
            return "🔴 Riesgo: bajo volumen y baja rentabilidad"

        matriz["Clasificación Gerencial"] = matriz.apply(clasificar_cliente, axis=1)
        matriz = matriz.sort_values(["Margen", "Facturación"], ascending=False)

        matriz_show = matriz.rename(columns={
            "__CLIENTE__": "Cliente",
            "Participación Facturación": "Part. Facturación",
            "Participación Margen": "Part. Margen",
        })[[
            "Cliente", "Servicios", "Facturación", "Costos", "Margen", "Rentabilidad",
            "Part. Facturación", "Part. Margen", "Clasificación Gerencial"
        ]]
        show_table(matriz_show, "Matriz Gerencial de Rentabilidad por Cliente", height=520)

        st.download_button(
            "Descargar matriz gerencial de rentabilidad",
            data=to_excel_bytes(matriz_show),
            file_name="matriz_gerencial_rentabilidad_clientes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_matriz_rentabilidad_clientes"
        )

        st.markdown("### Rentabilidad por vehículo y tipo de vehículo")
        c5, c6 = st.columns(2)
        with c5:
            placa_margin = aggregate(df, ["__PLACA__"]).query("__PLACA__ != 'Sin dato'").sort_values("Margen", ascending=False)
            chart_bar(placa_margin, "__PLACA__", "Margen", "Top 20 Mayor Margen por Placa", 20, key="rent_placa_mayor_margen", ascending=False)
        with c6:
            tipo_rent = aggregate(df, ["__TIPO_VEHICULO__"]).sort_values("Rentabilidad", ascending=False)
            chart_bar(tipo_rent, "__TIPO_VEHICULO__", "Rentabilidad", "Rentabilidad por Tipo Vehículo", 20, key="rent_tipo_vehiculo", ascending=False)

with tabs[5]:
    resp_col = "__COORDINADOR__" if df["__COORDINADOR__"].nunique() > 1 else "__USUARIO__"
    resp_name = "Coordinador" if resp_col == "__COORDINADOR__" else "Usuario / Creado Por"
    resp_tbl = aggregate(df, [resp_col]).sort_values("Facturación", ascending=False)
    chart_bar(resp_tbl, resp_col, "Facturación", f"Desempeño por {resp_name}", 20)
    chart_heatmap(df, resp_col, "__CLIENTE__", "__V_CLIENTE__", f"Mapa de Calor: {resp_name} vs Cliente - Facturación")
    tmp_resp = aggregate(df, [resp_col, "__CLIENTE__"])
    chart_heatmap(tmp_resp, resp_col, "__CLIENTE__", "Rentabilidad", f"Mapa de Calor: {resp_name} vs Cliente - Rentabilidad", aggfunc="mean")
    show_table(resp_tbl.rename(columns={resp_col: resp_name}), f"Tabla {resp_name}")

with tabs[6]:
    c1, c2 = st.columns(2)
    placa_tbl = aggregate(df, ["__PLACA__"]).query("__PLACA__ != 'Sin dato'").sort_values("Facturación", ascending=False)
    cond_tbl = aggregate(df, ["__CONDUCTOR__"]).query("__CONDUCTOR__ != 'Sin dato'").sort_values("Servicios", ascending=False)
    with c1:
        chart_bar(placa_tbl, "__PLACA__", "Facturación", "Top 20 Vehículos por Facturación", 20)
    with c2:
        chart_bar(cond_tbl, "__CONDUCTOR__", "Servicios", "Top 20 Conductores por Producción", 20)
    tipo_veh = aggregate(df, ["__TIPO_VEHICULO__"]).sort_values("Facturación", ascending=False)
    show_table(tipo_veh.rename(columns={"__TIPO_VEHICULO__": "Tipo Vehículo"}), "Tipo de Vehículo")

with tabs[7]:
    lat_col, lon_col = colmap.get("lat"), colmap.get("lon")
    if lat_col and lon_col:
        mdf = df.copy()
        mdf["lat"] = pd.to_numeric(mdf[lat_col], errors="coerce")
        mdf["lon"] = pd.to_numeric(mdf[lon_col], errors="coerce")
        mdf = mdf.dropna(subset=["lat", "lon"])
    else:
        mdf = df.copy()
        coords = mdf["__ORIGEN__"].apply(find_city_coords)
        mdf["lat"] = coords.apply(lambda x: x[0] if x else np.nan)
        mdf["lon"] = coords.apply(lambda x: x[1] if x else np.nan)
        mdf = mdf.dropna(subset=["lat", "lon"])
    if not mdf.empty and PLOTLY_OK:
        map_agg = mdf.groupby(["__ORIGEN__", "lat", "lon"], as_index=False).agg(Servicios=("__SERVICIOS__", "sum"), Facturación=("__V_CLIENTE__", "sum"))
        fig = px.scatter_mapbox(map_agg, lat="lat", lon="lon", size="Servicios", color="Facturación", hover_name="__ORIGEN__", zoom=4, mapbox_style="carto-positron", title="Mapa Estratégico por Origen")
        fig.update_layout(height=620)
        fig = apply_dark_plotly_theme(fig)
        st.plotly_chart(fig, use_container_width=True, key=unique_plot_key("mapa_estrategico", len(mdf)))
    elif not mdf.empty:
        st.map(mdf[["lat", "lon"]])
    else:
        st.info("No se encontraron coordenadas ni ciudades reconocidas para generar mapa. Se muestra análisis Origen/Destino.")
    c1, c2 = st.columns(2)
    with c1:
        origen_tbl = aggregate(df, ["__ORIGEN__"]).sort_values("Servicios", ascending=False)
        chart_bar(origen_tbl, "__ORIGEN__", "Servicios", "Top Orígenes", 20)
    with c2:
        destino_tbl = aggregate(df, ["__DESTINO__"]).sort_values("Servicios", ascending=False)
        chart_bar(destino_tbl, "__DESTINO__", "Servicios", "Top Destinos", 20)

with tabs[8]:
    st.markdown("### Centro de Alertas Ejecutivas")
    alerts: List[Tuple[str, str]] = []
    var_tbl = variation_table(df_compare if not df_compare.empty else df, "__CLIENTE__", pcol, "__V_CLIENTE__")
    if not var_tbl.empty:
        for _, r in var_tbl[var_tbl["Variación %"] <= -0.20].head(10).iterrows():
            alerts.append(("red", f"🔴 Cliente {r['__CLIENTE__']} presenta caída de {fmt_pct(r['Variación %'])} en facturación."))
        for _, r in var_tbl[var_tbl["Variación %"] >= 0.15].sort_values("Variación %", ascending=False).head(10).iterrows():
            alerts.append(("green", f"🟢 Cliente {r['__CLIENTE__']} presenta crecimiento de {fmt_pct(r['Variación %'])} en facturación."))
    rent_tbl = aggregate(df, ["__CLIENTE__"])
    for _, r in rent_tbl[rent_tbl["Rentabilidad"] < 0.15].sort_values("Rentabilidad").head(10).iterrows():
        alerts.append(("red", f"🔴 Cliente {r['__CLIENTE__']} tiene rentabilidad inferior al 15%: {fmt_pct(r['Rentabilidad'])}."))
    total_fact = rent_tbl["Facturación"].sum()
    if total_fact:
        rent_tbl["Participación"] = rent_tbl["Facturación"] / total_fact
        for _, r in rent_tbl[rent_tbl["Participación"] > 0.30].sort_values("Participación", ascending=False).head(10).iterrows():
            alerts.append(("yellow", f"🟡 Alta concentración: {r['__CLIENTE__']} representa {fmt_pct(r['Participación'])} de la facturación."))
    if alerts:
        for level, txt in alerts[:30]:
            st.markdown(f"<div class='alert-{level}'>{txt}</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='alert-green'>🟢 No se identifican alertas críticas con los filtros actuales.</div>", unsafe_allow_html=True)
    if not var_tbl.empty:
        show_table(var_tbl.rename(columns={"__CLIENTE__": "Cliente"}), "Análisis de Variaciones por Cliente")
    show_table(rent_tbl.rename(columns={"__CLIENTE__": "Cliente"}), "Concentración y Rentabilidad por Cliente")



with tabs[9]:
    st.markdown("### Seguimiento operativo de personal y cierre mensual")
    st.markdown(
        """
        <div class='info-box'>
        Regla de negocio: <b>OTIF / % de cierre</b> = servicios con <b>ESTADO OP = CUMPLIDO OPERATIVO</b> / total de servicios no anulados. 
        Los estados CUMPLIDO, EN PROGRAMACIÓN, EN TRÁNSITO y demás estados diferentes de CUMPLIDO OPERATIVO quedan como pendientes de cierre.
        </div>
        """,
        unsafe_allow_html=True,
    )

    cierre = closure_summary(df)
    corte_txt = "Sin fecha"
    if pd.notna(cierre.get("corte")):
        corte_txt = pd.to_datetime(cierre["corte"]).strftime("%d/%m/%Y")

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        make_kpi("OTIF Cierre", fmt_pct(cierre["otif"]), f"Corte: {corte_txt}")
    with k2:
        make_kpi("Servicios a cerrar", fmt_num(cierre["total"]), "Total sin anulados")
    with k3:
        make_kpi("Remesas cerradas", fmt_num(cierre["cerrados"]), "Estado: Cumplido Operativo")
    with k4:
        make_kpi("Pendientes de cierre", fmt_num(cierre["pendientes"]), "No están en Cumplido Operativo")

    person_col = "__USUARIO__"
    person_label = "Quien creó"
    if df["__USUARIO__"].nunique() <= 1 and df["__COORDINADOR__"].nunique() > 1:
        person_col = "__COORDINADOR__"
        person_label = "Coordinador"

    st.markdown(f"### Tabla dinámica de cierre por {person_label}")
    cierre_personal = closure_status_matrix(df, person_col)
    if not cierre_personal.empty:
        show_table(cierre_personal, f"Seguimiento operativo: {person_label} vs Estado OP", height=520)
        st.download_button(
            "Descargar seguimiento de cierre por personal",
            data=to_excel_bytes(cierre_personal),
            file_name="seguimiento_cierre_por_personal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_seguimiento_cierre_personal"
        )
    else:
        st.info("No hay datos suficientes para construir el seguimiento por personal.")

    st.markdown("### Seguimiento de cierre por cliente")
    st.markdown(
        """
        <div class='info-box'>
        Esta matriz permite ver <b>Cliente vs Estado OP</b>. Ayuda a identificar qué cliente tiene más remesas pendientes de cierre y cuál está afectando el cierre mensual.
        </div>
        """,
        unsafe_allow_html=True,
    )
    cierre_cliente = closure_matrix_by_dimension(df, "__CLIENTE__", "Cliente")
    if not cierre_cliente.empty:
        cierre_cliente_total = add_total_row_closure(cierre_cliente, "Cliente")
        show_table(cierre_cliente_total, "Cliente vs Estado OP - Seguimiento de cierre", height=560)
        st.download_button(
            "Descargar Cliente vs Estado OP",
            data=to_excel_bytes(cierre_cliente_total),
            file_name="cliente_vs_estado_op_cierre.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_cliente_estado_op_cierre"
        )

        ccli1, ccli2 = st.columns(2)
        with ccli1:
            chart_bar(
                cierre_cliente.sort_values("Pendientes cierre", ascending=False),
                "Cliente", "Pendientes cierre", "Clientes con más pendientes de cierre", 20,
                key="clientes_mas_pendientes_cierre", ascending=False
            )
        with ccli2:
            chart_bar(
                cierre_cliente.sort_values("OTIF Cierre", ascending=False),
                "Cliente", "OTIF Cierre", "% Cierre por Cliente", 20,
                key="otif_cierre_cliente", ascending=False
            )
    else:
        st.info("No hay datos suficientes para construir Cliente vs Estado OP.")

    st.markdown("### Cliente + Quien creó vs Estado OP")
    cierre_cliente_persona = closure_matrix_client_person(df, person_col)
    if not cierre_cliente_persona.empty:
        show_table(cierre_cliente_persona, "Cliente + Quien creó vs Estado OP", height=560)
        st.download_button(
            "Descargar Cliente + Quien creó vs Estado OP",
            data=to_excel_bytes(cierre_cliente_persona),
            file_name="cliente_quien_creo_vs_estado_op.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_cliente_persona_estado_op"
        )
    else:
        st.info("No hay datos suficientes para construir Cliente + Quien creó vs Estado OP.")

    c1, c2 = st.columns(2)
    with c1:
        ranking_cierre = closure_person_table(df, person_col)
        if not ranking_cierre.empty:
            chart_bar(
                ranking_cierre[ranking_cierre[person_label if person_label in ranking_cierre.columns else "Quien creó"] != "Total general"] if person_label in ranking_cierre.columns else ranking_cierre,
                "Quien creó", "OTIF Cierre", "Ranking % Cierre por Quien Creó", 20,
                key="otif_ranking_personal", ascending=False
            )
    with c2:
        pendientes_personal = closure_person_table(df, person_col)
        if not pendientes_personal.empty:
            chart_bar(
                pendientes_personal.sort_values("Pendientes cierre", ascending=False),
                "Quien creó", "Pendientes cierre", "Pendientes de Cierre por Quien Creó", 20,
                key="otif_pendientes_personal", ascending=False
            )

    st.markdown("### Evolución del cierre por periodo")
    cierre_periodo = closure_by_period(df_compare if not df_compare.empty else df, pcol)
    if not cierre_periodo.empty:
        c3, c4 = st.columns(2)
        with c3:
            chart_bar(cierre_periodo, pcol, "OTIF Cierre", f"% Cierre - {periodo}", 20, key="otif_periodo", ascending=False)
        with c4:
            cierre_stack = cierre_periodo.rename(columns={pcol: "Periodo"})
            show_table(cierre_stack, "Detalle cierre por periodo", height=360)
    else:
        st.info("No hay suficientes periodos para mostrar evolución de cierre.")

    st.markdown("### Recomendación de uso para remuneración mensual")
    st.markdown(
        """
        <div class='warn-box'>
        Para remuneración, usa como indicador principal <b>OTIF Cierre</b> por <b>Quien creó</b>, con una regla mínima de volumen. 
        Ejemplo gerencial: pagar incentivo si el responsable tiene mínimo 50 servicios en el mes y alcanza 95% o más de cierre a corte de fin de mes.
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("---")
st.download_button("⬇️ Descargar base filtrada en Excel", to_excel_bytes(df), "base_filtrada_informe_gerencial.xlsx")
st.caption("Informe Gerencial Ejecutivo | Streamlit Cloud | Base cargada dinámicamente por el usuario")
