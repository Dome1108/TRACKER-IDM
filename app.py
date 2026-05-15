import streamlit as st
import pandas as pd
import plotly.express as px
import requests
from io import BytesIO
from datetime import date
from pathlib import Path
from html import escape

# =====================================================
# CONFIGURACIÓN GENERAL
# =====================================================

st.set_page_config(
    page_title="Tracker IDM",
    page_icon="📌",
    layout="wide"
)

# Archivo local opcional
EXCEL_PATH = None

# URL RAW del Excel en GitHub
EXCEL_URL = "https://raw.githubusercontent.com/Dome1108/TRACKER-IDM/main/TRACKER%20IDM.xlsx"

# Pestaña que debe leer del Excel
EXCEL_SHEET = "BASE"


# =====================================================
# ESTILOS
# =====================================================

st.markdown(
    """
    <style>
    .stApp {
        background-color: #0f1117;
        color: #f2f2f2;
    }

    section[data-testid="stSidebar"] {
        background-color: #171923;
    }

    .main-title {
        font-size: 36px;
        font-weight: 800;
        margin-bottom: 4px;
        color: #ffffff;
    }

    .subtitle {
        color: #b8b8b8;
        font-size: 15px;
        margin-bottom: 25px;
    }

    .section-title {
        font-size: 25px;
        font-weight: 800;
        margin-top: 30px;
        margin-bottom: 15px;
        color: #ffffff;
    }

    .project-card {
        background-color: #1f2430;
        border-radius: 16px;
        padding: 18px;
        margin-bottom: 14px;
        border: 1px solid #30384a;
        box-shadow: 0 4px 14px rgba(0,0,0,0.25);
    }

    .project-title {
        font-size: 19px;
        font-weight: 800;
        color: #ffffff;
        margin-bottom: 8px;
        line-height: 1.35;
    }

    .project-meta {
        font-size: 14px;
        color: #d1d5db;
        margin-bottom: 5px;
        line-height: 1.4;
    }

    .tag {
        display: inline-block;
        padding: 5px 10px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
        margin-right: 6px;
        margin-top: 9px;
    }

    .tag-ok {
        background-color: #123d2a;
        color: #67e8a3;
    }

    .tag-progress {
        background-color: #3d3312;
        color: #ffd166;
    }

    .tag-plan {
        background-color: #13294b;
        color: #83b9ff;
    }

    .tag-danger {
        background-color: #4a1515;
        color: #ff8a8a;
    }

    .tag-warning {
        background-color: #4a3715;
        color: #ffd27a;
    }

    .tag-neutral {
        background-color: #3a3a3a;
        color: #d0d0d0;
    }

    .progress-label {
        font-size: 12px;
        color: #bdbdbd;
        margin-top: 12px;
        margin-bottom: 4px;
    }

    .small-note {
        font-size: 13px;
        color: #9ca3af;
        margin-bottom: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# =====================================================
# FUNCIONES AUXILIARES
# =====================================================

def normalize_col(col):
    return (
        str(col)
        .strip()
        .upper()
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
        .replace("Ñ", "N")
    )


def find_column(df, possible_names):
    normalized = {normalize_col(c): c for c in df.columns}

    for name in possible_names:
        key = normalize_col(name)
        if key in normalized:
            return normalized[key]

    return None


def clean_percentage(value):
    if pd.isna(value):
        return 0

    if isinstance(value, str):
        value = value.replace("%", "").replace(",", ".").strip()

    try:
        value = float(value)
    except Exception:
        return 0

    if value <= 1:
        value = value * 100

    return max(0, min(100, value))


def status_tag_class(status):
    status = str(status).lower()

    if "completo" in status or "finalizado" in status:
        return "tag-ok"
    elif "proceso" in status:
        return "tag-progress"
    elif "planificado" in status:
        return "tag-plan"
    elif "stop" in status or "atrasado" in status or "vencido" in status:
        return "tag-danger"
    else:
        return "tag-neutral"


def due_tag(fecha):
    if pd.isna(fecha):
        return "Sin fecha", "tag-neutral"

    today = pd.Timestamp(date.today()).normalize()
    fecha = pd.Timestamp(fecha).normalize()
    days = (fecha - today).days

    if days < 0:
        return f"Vencido hace {abs(days)} días", "tag-danger"
    elif days == 0:
        return "Vence hoy", "tag-danger"
    elif days <= 7:
        return f"Vence en {days} días", "tag-warning"
    elif days <= 30:
        return f"Vence en {days} días", "tag-progress"
    else:
        return f"Vence en {days} días", "tag-neutral"


def get_excel_bytes_from_url(url):
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        return BytesIO(response.content)
    except Exception as e:
        st.error(f"No se pudo descargar el Excel desde GitHub. Error: {e}")
        st.stop()


@st.cache_data(show_spinner=False, ttl=300)
def load_excel_from_url(url):
    excel_bytes = get_excel_bytes_from_url(url)
    return read_excel_base(excel_bytes)


@st.cache_data(show_spinner=False)
def load_excel_from_upload(uploaded_file):
    return read_excel_base(uploaded_file)


@st.cache_data(show_spinner=False)
def load_excel_from_path(path):
    return read_excel_base(path)


def read_excel_base(file_or_buffer):
    try:
        excel_file = pd.ExcelFile(file_or_buffer, engine="openpyxl")
    except Exception as e:
        st.error(f"No se pudo leer el archivo Excel. Error: {e}")
        st.stop()

    sheet_found = None

    for sheet in excel_file.sheet_names:
        if sheet.strip().upper() == EXCEL_SHEET.upper():
            sheet_found = sheet
            break

    if sheet_found is None:
        st.error(
            f"No se encontró la pestaña '{EXCEL_SHEET}'. "
            f"Pestañas encontradas: {', '.join(excel_file.sheet_names)}"
        )
        st.stop()

    df = pd.read_excel(excel_file, sheet_name=sheet_found, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    return df


def prepare_data(df):
    col_anio = find_column(df, ["AÑO", "ANIO"])
    col_clasificacion = find_column(df, ["CLASIFICACION", "CLASIFICACIÓN", "CLASIFICAC"])
    col_tipo = find_column(df, ["TIPO"])
    col_proyecto = find_column(df, ["PROYECTO"])
    col_mini = find_column(df, ["MINI PROYECTO", "MINIPROYECTO", "MINI_PROYECTO"])
    col_descripcion = find_column(df, ["DESCRIPCION", "DESCRIPCIÓN"])
    col_encargado = find_column(df, ["ENCARGADO", "ENCARGA", "RESPONSABLE"])
    col_equipo = find_column(df, ["EQUIPO DE SOPORTE", "EQUIPO DE SOP", "EQUIPO"])
    col_estado = find_column(df, ["ESTADO"])
    col_pct = find_column(df, ["% COMPLETADO", "% COMPLE", "% COMPL", "PORCENTAJE", "AVANCE"])

    col_fecha = find_column(df, [
        "FECHA_ENTREGA",
        "FECHA ENTREGA",
        "FECHA DE ENTREGA",
        "FECHA FINALIZACION",
        "FECHA FINALIZACIÓN",
        "FECHA DE FINALIZACION",
        "FECHA DE FINALIZACIÓN",
        "VENCIMIENTO",
        "FECHA VENCIMIENTO",
        "FECHA DE VENCIMIENTO"
    ])

    required = {
        "TIPO": col_tipo,
        "PROYECTO": col_proyecto,
        "MINI PROYECTO": col_mini,
        "ESTADO": col_estado,
        "FECHA_ENTREGA": col_fecha,
    }

    missing = [k for k, v in required.items() if v is None]

    if missing:
        st.error(
            "Faltan estas columnas obligatorias en la pestaña BASE: "
            + ", ".join(missing)
        )
        st.info("Columnas encontradas en tu Excel:")
        st.write(list(df.columns))
        st.stop()

    data = pd.DataFrame()

    data["anio"] = df[col_anio] if col_anio else ""
    data["clasificacion"] = df[col_clasificacion] if col_clasificacion else ""
    data["tipo"] = df[col_tipo].fillna("Sin tipo").astype(str).str.strip()
    data["proyecto"] = df[col_proyecto].fillna("").astype(str).str.strip()
    data["mini_proyecto"] = df[col_mini].fillna("").astype(str).str.strip()

    if col_descripcion:
        data["descripcion"] = df[col_descripcion].fillna("").astype(str).str.strip()
    else:
        data["descripcion"] = ""

    if col_encargado:
        data["encargado"] = df[col_encargado].fillna("").astype(str).str.strip()
    else:
        data["encargado"] = ""

    if col_equipo:
        data["equipo"] = df[col_equipo].fillna("").astype(str).str.strip()
    else:
        data["equipo"] = ""

    data["estado"] = df[col_estado].fillna("Sin estado").astype(str).str.strip()

    if col_pct:
        data["avance"] = df[col_pct].apply(clean_percentage)
    else:
        data["avance"] = 0

    data["fecha_entrega"] = pd.to_datetime(
        df[col_fecha],
        errors="coerce",
        dayfirst=True
    )

    today = pd.Timestamp(date.today()).normalize()

    data["dias_restantes"] = (data["fecha_entrega"] - today).dt.days
    data["sin_fecha"] = data["fecha_entrega"].isna()

    data["fecha_texto"] = data["fecha_entrega"].dt.strftime("%d/%m/%Y")
    data["fecha_texto"] = data["fecha_texto"].fillna("Sin fecha")

    data = data.sort_values(
        by=["sin_fecha", "fecha_entrega", "avance"],
        ascending=[True, True, True]
    )

    return data


def render_project_card(row):
    titulo = row["mini_proyecto"] or row["proyecto"]
    proyecto = row["proyecto"]
    estado = row["estado"]
    avance = int(row["avance"])
    fecha = row["fecha_entrega"]

    due_text, due_class = due_tag(fecha)
    estado_class = status_tag_class(estado)

    fecha_text = "Sin fecha"
    if not pd.isna(fecha):
        fecha_text = pd.Timestamp(fecha).strftime("%d/%m/%Y")

    titulo = escape(str(titulo))
    proyecto = escape(str(proyecto))
    descripcion = escape(str(row["descripcion"]))
    encargado = escape(str(row["encargado"]))
    equipo = escape(str(row["equipo"]))
    tipo = escape(str(row["tipo"]))
    estado = escape(str(estado))
    fecha_text = escape(str(fecha_text))
    due_text = escape(str(due_text))

    descripcion_html = ""
    if descripcion.strip():
        descripcion_html = f"""
        <div class="project-meta">
            <b>Descripción:</b> {descripcion}
        </div>
        """

    encargado_html = ""
    if encargado.strip():
        encargado_html = f"""
        <div class="project-meta">
            <b>Encargado:</b> {encargado}
        </div>
        """

    equipo_html = ""
    if equipo.strip():
        equipo_html = f"""
        <div class="project-meta">
            <b>Equipo soporte:</b> {equipo}
        </div>
        """

    st.markdown(
        f"""
        <div class="project-card">
            <div class="project-title">○ {titulo}</div>

            <div class="project-meta">
                <b>Tipo:</b> {tipo}
            </div>

            <div class="project-meta">
                <b>Proyecto:</b> {proyecto}
            </div>

            {descripcion_html}
            {encargado_html}
            {equipo_html}

            <div class="project-meta">
                <b>Fecha entrega:</b> {fecha_text}
            </div>

            <span class="tag {estado_class}">{estado}</span>
            <span class="tag {due_class}">{due_text}</span>

            <div class="progress-label">Avance: {avance}%</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.progress(avance / 100)


def apply_plotly_dark_layout(fig):
    fig.update_layout(
        paper_bgcolor="#0f1117",
        plot_bgcolor="#0f1117",
        font=dict(color="#f2f2f2"),
        margin=dict(l=20, r=20, t=50, b=20),
        legend=dict(
            bgcolor="rgba(0,0,0,0)",
            font=dict(color="#f2f2f2")
        )
    )

    fig.update_xaxes(
        gridcolor="#30384a",
        zerolinecolor="#30384a"
    )

    fig.update_yaxes(
        gridcolor="#30384a",
        zerolinecolor="#30384a"
    )

    return fig


# =====================================================
# TÍTULO
# =====================================================

st.markdown(
    '<div class="main-title">📌 Tracker IDM</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">Dashboard interactivo alimentado desde la pestaña BASE del Excel en GitHub.</div>',
    unsafe_allow_html=True
)


# =====================================================
# SIDEBAR
# =====================================================

with st.sidebar:
    st.header("⚙️ Configuración")

    uploaded_file = st.file_uploader(
        "Sube otro Excel manualmente",
        type=["xlsx", "xls"]
    )

    st.caption("Si no subes archivo, se usará automáticamente el Excel de GitHub.")

    if st.button("Actualizar datos desde GitHub"):
        st.cache_data.clear()
        st.rerun()

    st.divider()

    st.caption("Fuente actual por defecto:")
    st.code(EXCEL_URL)

    st.divider()

    st.caption("Columnas recomendadas en la pestaña BASE:")
    st.code(
        """AÑO
TIPO
PROYECTO
MINI PROYECTO
DESCRIPCIÓN
ENCARGADO
EQUIPO DE SOPORTE
ESTADO
% COMPL
FECHA_ENTREGA"""
    )


# =====================================================
# CARGA DE DATOS
# =====================================================

if uploaded_file is not None:
    raw_df = load_excel_from_upload(uploaded_file)
    fuente_datos = "Archivo subido manualmente"

elif EXCEL_URL:
    raw_df = load_excel_from_url(EXCEL_URL)
    fuente_datos = "GitHub"

elif EXCEL_PATH and Path(EXCEL_PATH).exists():
    raw_df = load_excel_from_path(EXCEL_PATH)
    fuente_datos = "Archivo local"

else:
    st.info("Sube un archivo Excel para visualizar el dashboard.")
    st.stop()


df = prepare_data(raw_df)

st.caption(f"Fuente de datos: {fuente_datos} | Pestaña: {EXCEL_SHEET}")


# =====================================================
# FILTROS INTERACTIVOS
# =====================================================

st.markdown('<div class="section-title">Filtros interactivos</div>', unsafe_allow_html=True)

f1, f2, f3, f4 = st.columns(4)

with f1:
    tipos_disponibles = sorted(df["tipo"].dropna().unique())

    filtro_tipo = st.multiselect(
        "Tipo",
        options=tipos_disponibles,
        default=tipos_disponibles
    )

with f2:
    estados_disponibles = sorted(df["estado"].dropna().unique())

    filtro_estado = st.multiselect(
        "Estado",
        options=estados_disponibles,
        default=estados_disponibles
    )

with f3:
    encargados_disponibles = sorted([
        x for x in df["encargado"].dropna().unique()
        if str(x).strip() != ""
    ])

    filtro_encargado = st.multiselect(
        "Encargado",
        options=encargados_disponibles,
        default=encargados_disponibles
    )

with f4:
    busqueda = st.text_input("Buscar proyecto")


df_filtrado = df.copy()

df_filtrado = df_filtrado[df_filtrado["tipo"].isin(filtro_tipo)]
df_filtrado = df_filtrado[df_filtrado["estado"].isin(filtro_estado)]

if encargados_disponibles:
    df_filtrado = df_filtrado[df_filtrado["encargado"].isin(filtro_encargado)]

if busqueda:
    busqueda_lower = busqueda.lower()

    df_filtrado = df_filtrado[
        df_filtrado["proyecto"].astype(str).str.lower().str.contains(busqueda_lower, na=False)
        | df_filtrado["mini_proyecto"].astype(str).str.lower().str.contains(busqueda_lower, na=False)
        | df_filtrado["descripcion"].astype(str).str.lower().str.contains(busqueda_lower, na=False)
        | df_filtrado["encargado"].astype(str).str.lower().str.contains(busqueda_lower, na=False)
    ]


# =====================================================
# FILTRO DE FECHAS
# =====================================================

fechas_validas = df_filtrado["fecha_entrega"].dropna()

if not fechas_validas.empty:
    fecha_min = fechas_validas.min().date()
    fecha_max = fechas_validas.max().date()

    rango_fechas = st.slider(
        "Rango de fecha de entrega",
        min_value=fecha_min,
        max_value=fecha_max,
        value=(fecha_min, fecha_max),
        format="DD/MM/YYYY"
    )

    inicio = pd.Timestamp(rango_fechas[0])
    fin = pd.Timestamp(rango_fechas[1])

    incluir_sin_fecha = st.checkbox("Incluir proyectos sin fecha", value=True)

    if incluir_sin_fecha:
        df_filtrado = df_filtrado[
            (df_filtrado["fecha_entrega"].isna())
            | (
                (df_filtrado["fecha_entrega"] >= inicio)
                & (df_filtrado["fecha_entrega"] <= fin)
            )
        ]
    else:
        df_filtrado = df_filtrado[
            (df_filtrado["fecha_entrega"] >= inicio)
            & (df_filtrado["fecha_entrega"] <= fin)
        ]


# =====================================================
# KPIS
# =====================================================

st.markdown('<div class="section-title">Resumen general</div>', unsafe_allow_html=True)

total = len(df_filtrado)

completos = (
    df_filtrado["estado"]
    .astype(str)
    .str.lower()
    .str.contains("completo|finalizado", na=False)
    .sum()
)

en_proceso = (
    df_filtrado["estado"]
    .astype(str)
    .str.lower()
    .str.contains("proceso", na=False)
    .sum()
)

planificados = (
    df_filtrado["estado"]
    .astype(str)
    .str.lower()
    .str.contains("planificado", na=False)
    .sum()
)

stop = (
    df_filtrado["estado"]
    .astype(str)
    .str.lower()
    .str.contains("stop", na=False)
    .sum()
)

vencidos = (df_filtrado["dias_restantes"] < 0).sum()

avance_promedio = 0
if total > 0:
    avance_promedio = round(df_filtrado["avance"].mean(), 1)

k1, k2, k3, k4, k5, k6 = st.columns(6)

with k1:
    st.metric("Total proyectos", int(total))

with k2:
    st.metric("Completos", int(completos))

with k3:
    st.metric("En proceso", int(en_proceso))

with k4:
    st.metric("Planificados", int(planificados))

with k5:
    st.metric("Stop", int(stop))

with k6:
    st.metric("Vencidos", int(vencidos))

st.metric("Avance promedio", f"{avance_promedio}%")


# =====================================================
# DASHBOARD INTERACTIVO
# =====================================================

st.markdown('<div class="section-title">Dashboard interactivo</div>', unsafe_allow_html=True)

if df_filtrado.empty:
    st.warning("No hay proyectos para mostrar con los filtros seleccionados.")
    st.stop()

g1, g2 = st.columns(2)

with g1:
    estado_count = (
        df_filtrado
        .groupby("estado", as_index=False)
        .size()
        .rename(columns={"size": "cantidad"})
        .sort_values("cantidad", ascending=False)
    )

    fig_estado = px.bar(
        estado_count,
        x="estado",
        y="cantidad",
        text="cantidad",
        title="Distribución de proyectos por estado",
        labels={
            "estado": "Estado",
            "cantidad": "Cantidad de proyectos"
        }
    )

    fig_estado.update_traces(textposition="outside")
    fig_estado = apply_plotly_dark_layout(fig_estado)

    st.plotly_chart(fig_estado, use_container_width=True)

with g2:
    tipo_count = (
        df_filtrado
        .groupby("tipo", as_index=False)
        .size()
        .rename(columns={"size": "cantidad"})
        .sort_values("cantidad", ascending=False)
    )

    fig_tipo = px.pie(
        tipo_count,
        names="tipo",
        values="cantidad",
        title="Distribución de proyectos por tipo",
        hole=0.45
    )

    fig_tipo = apply_plotly_dark_layout(fig_tipo)

    st.plotly_chart(fig_tipo, use_container_width=True)


# =====================================================
# PRÓXIMOS VENCIMIENTOS
# =====================================================

st.markdown('<div class="section-title">Próximos vencimientos</div>', unsafe_allow_html=True)

df_vencimientos = df_filtrado.dropna(subset=["fecha_entrega"]).copy()

df_vencimientos["nombre_tarjeta"] = df_vencimientos["mini_proyecto"]

df_vencimientos.loc[
    df_vencimientos["nombre_tarjeta"].astype(str).str.strip() == "",
    "nombre_tarjeta"
] = df_vencimientos["proyecto"]

df_vencimientos = df_vencimientos.sort_values("fecha_entrega", ascending=True).head(20)

if not df_vencimientos.empty:
    fig_venc = px.scatter(
        df_vencimientos,
        x="fecha_entrega",
        y="nombre_tarjeta",
        size="avance",
        color="estado",
        hover_data={
            "proyecto": True,
            "tipo": True,
            "encargado": True,
            "fecha_texto": True,
            "avance": True,
            "fecha_entrega": False,
            "nombre_tarjeta": False
        },
        title="Top 20 proyectos más próximos a vencer",
        labels={
            "fecha_entrega": "Fecha de entrega",
            "nombre_tarjeta": "Proyecto",
            "estado": "Estado",
            "avance": "Avance"
        }
    )

    fig_venc = apply_plotly_dark_layout(fig_venc)
    st.plotly_chart(fig_venc, use_container_width=True)
else:
    st.info("No hay fechas de entrega registradas para graficar vencimientos.")


# =====================================================
# TABLA INTERACTIVA
# =====================================================

st.markdown('<div class="section-title">Tabla interactiva</div>', unsafe_allow_html=True)

tabla = df_filtrado[
    [
        "tipo",
        "proyecto",
        "mini_proyecto",
        "descripcion",
        "encargado",
        "equipo",
        "estado",
        "avance",
        "fecha_texto",
        "dias_restantes"
    ]
].copy()

tabla = tabla.rename(
    columns={
        "tipo": "Tipo",
        "proyecto": "Proyecto",
        "mini_proyecto": "Mini proyecto",
        "descripcion": "Descripción",
        "encargado": "Encargado",
        "equipo": "Equipo de soporte",
        "estado": "Estado",
        "avance": "% completado",
        "fecha_texto": "Fecha entrega",
        "dias_restantes": "Días restantes"
    }
)

st.dataframe(
    tabla,
    use_container_width=True,
    hide_index=True
)


# =====================================================
# TARJETAS ORDENADAS POR FECHA DE ENTREGA
# =====================================================

st.markdown('<div class="section-title">Detalle de proyectos por prioridad de entrega</div>', unsafe_allow_html=True)

st.markdown(
    '<div class="small-note">Orden automático: primero los proyectos con fecha más próxima, luego los más lejanos, y al final los proyectos sin fecha.</div>',
    unsafe_allow_html=True
)

df_cards = df_filtrado.sort_values(
    by=["sin_fecha", "fecha_entrega", "avance"],
    ascending=[True, True, True]
)

vista = st.radio(
    "Vista de tarjetas",
    options=["Lista general ordenada por fecha", "Agrupada por tipo"],
    horizontal=True
)

if vista == "Lista general ordenada por fecha":
    for _, row in df_cards.iterrows():
        render_project_card(row)

else:
    tipos_cards = list(df_cards["tipo"].dropna().unique())

    orden_preferido = [
        "Estudio Recurrente",
        "Estudios Recurrentes",
        "Iniciativa",
        "Iniciativas",
        "Solicitud Interna",
        "Solicitudes Internas",
        "Tendencias"
    ]

    tipos_ordenados = []

    for tipo_preferido in orden_preferido:
        for tipo_real in tipos_cards:
            if str(tipo_real).strip().lower() == tipo_preferido.lower():
                if tipo_real not in tipos_ordenados:
                    tipos_ordenados.append(tipo_real)

    for tipo_real in tipos_cards:
        if tipo_real not in tipos_ordenados:
            tipos_ordenados.append(tipo_real)

    max_columns = min(len(tipos_ordenados), 4)

    if max_columns == 0:
        st.warning("No hay proyectos para mostrar.")
    else:
        cols = st.columns(max_columns)

        for idx, tipo in enumerate(tipos_ordenados):
            col = cols[idx % max_columns]

            with col:
                st.markdown(
                    f"""
                    <div class="section-title" style="font-size:20px;">
                        {escape(str(tipo))}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                subset = df_cards[df_cards["tipo"] == tipo].copy()

                subset = subset.sort_values(
                    by=["sin_fecha", "fecha_entrega", "avance"],
                    ascending=[True, True, True]
                )

                for _, row in subset.iterrows():
                    render_project_card(row)
