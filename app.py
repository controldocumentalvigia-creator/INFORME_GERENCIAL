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


def money_to_number(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0)
    cleaned = (
        s.astype(str)
        .str.replace("COP", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(r"[^0-9,\.\-]", "", regex=True)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


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

    estado = colmap.get("estado_op")
    if estado:
        norm_estado = df[estado].apply(normalize_text)
        df = df[~norm_estado.eq("ANULADO")].copy()

    return df, colmap


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


def chart_bar(data: pd.DataFrame, x: str, y: str, title: str, top: int = 20, key: Optional[str] = None) -> None:
    if data.empty or x not in data.columns or y not in data.columns:
        st.info("Sin datos para graficar.")
        return
    d = data.sort_values(y, ascending=False).head(top).copy()
    d["__LABEL__"] = d[y].apply(lambda v: value_label(v, y))
    if key is None:
        key = unique_plot_key("bar", title, x, y, top)
    else:
        key = unique_plot_key("bar", key, title, x, y, top)
    if PLOTLY_OK:
        fig = px.bar(d, x=x, y=y, text="__LABEL__", title=title)
        fig.update_layout(
            height=430, plot_bgcolor="white", paper_bgcolor="white", title_font_size=18,
            yaxis=dict(title=y, tickprefix="$ " if y in ["Facturación", "Costos", "Margen"] else "", tickformat=",.0f" if y in ["Facturación", "Costos", "Margen", "Servicios"] else ".0%" if y == "Rentabilidad" else None),
        )
        fig.update_traces(marker_color="#0B4F78", textposition="outside", cliponaxis=False, hovertemplate=f"%{{x}}<br>{y}: %{{text}}<extra></extra>")
        fig = apply_dark_plotly_theme(fig)
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
        elif col in ["Rentabilidad", "Participación", "Variación %", "Crecimiento", "Var Facturación %", "Var Margen %", "Var Servicios %"]:
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


def variation_table(df: pd.DataFrame, group_col: str, pcol: str, value_col: str) -> pd.DataFrame:
    periods = sorted([p for p in df[pcol].dropna().astype(str).unique().tolist() if p != "Sin fecha"])
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
    fin = aggregate(df, [pcol]).sort_values(pcol).copy()
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
    periods = sorted([p for p in df[pcol].dropna().astype(str).unique() if p != "Sin fecha"])
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
    periodo = st.selectbox("Periodo de análisis", ["Mensual", "Bimestral", "Trimestral", "Semestral", "Anual"], index=0)
    pcol = period_col(periodo)

    df = df_base.copy()
    if st.button("🔄 Limpiar filtros", use_container_width=True):
        for k in list(st.session_state.keys()):
            if k.startswith("flt_"):
                del st.session_state[k]
        st.rerun()
    df, _ = apply_filter(df, "Año", "__ANIO__", "flt_anio")
    df, _ = apply_filter(df, "Mes", "__MES__", "flt_mes")
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

if df.empty:
    st.warning("No hay registros con los filtros seleccionados.")
    st.stop()

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
with st.expander("Columnas detectadas en la base"):
    detect = pd.DataFrame([{"Campo estándar": k, "Columna encontrada": v or "No encontrada"} for k, v in colmap.items()])
    st.dataframe(detect, use_container_width=True)

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
])

with tabs[0]:
    st.markdown("### Resumen Ejecutivo Automático")
    st.markdown(f"<div class='card'>{executive_summary(df, pcol)}</div>", unsafe_allow_html=True)
    g_period = period_summary(df, pcol)
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
        chart_bar(clientes_tbl.sort_values("Rentabilidad", ascending=False), "__CLIENTE__", "Rentabilidad", "Top 10 Clientes por Rentabilidad", 10)
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
    fin = period_summary(df, pcol)
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
    var_fact = variation_table(df, "__CLIENTE__", pcol, "__V_CLIENTE__")
    if not var_fact.empty:
        show_table(var_fact.rename(columns={"__CLIENTE__": "Cliente"}), "Variación de Facturación por Cliente")

with tabs[4]:
    rent_cliente = aggregate(df, ["__CLIENTE__"]).sort_values("Rentabilidad", ascending=False)
    c1, c2 = st.columns(2)
    with c1:
        chart_bar(rent_cliente.head(20), "__CLIENTE__", "Rentabilidad", "Top 20 Más Rentables", 20)
    with c2:
        chart_bar(rent_cliente.sort_values("Rentabilidad", ascending=True).head(20), "__CLIENTE__", "Rentabilidad", "Top 20 Menos Rentables", 20)
    c3, c4 = st.columns(2)
    with c3:
        placa_margin = aggregate(df, ["__PLACA__"]).query("__PLACA__ != 'Sin dato'").sort_values("Margen", ascending=False)
        chart_bar(placa_margin, "__PLACA__", "Margen", "Top 20 Mayor Margen por Placa", 20)
    with c4:
        tipo_rent = aggregate(df, ["__TIPO_VEHICULO__"]).sort_values("Rentabilidad", ascending=False)
        chart_bar(tipo_rent, "__TIPO_VEHICULO__", "Rentabilidad", "Rentabilidad por Tipo Vehículo", 20)
    show_table(rent_cliente.rename(columns={"__CLIENTE__": "Cliente"}), "Rentabilidad por Cliente")

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
    var_tbl = variation_table(df, "__CLIENTE__", pcol, "__V_CLIENTE__")
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

st.markdown("---")
st.download_button("⬇️ Descargar base filtrada en Excel", to_excel_bytes(df), "base_filtrada_informe_gerencial.xlsx")
st.caption("Informe Gerencial Ejecutivo | Streamlit Cloud | Base cargada dinámicamente por el usuario")

