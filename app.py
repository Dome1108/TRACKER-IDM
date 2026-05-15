import streamlit as st
import pandas as pd
import plotly.express as px
import requests
from io import BytesIO
from datetime import date
from html import escape

# =====================================================
# CONFIGURACIÓN GENERAL
# =====================================================

st.set_page_config(
    page_title="Tracker IDM",
    page_icon="📌",
    layout="wide"
)

EXCEL_URL = "https://raw.githubusercontent.com/Dome1108/TRACKER-IDM/main/TRACKER%20IDM.xlsx"
EXCEL_SHEET = "BASE"


# =====================================================
# ESTILOS
# =====================================================

st.markdown(
    """
    <style>
    .stApp {
        background-color: #0f1117;
        color: #f4f4f5;
    }

    section[data-testid="stSidebar"] {
        background-color: #151923;
    }

    .main-title {
        font-size: 38px;
        font-weight: 800;
        margin-bottom: 4px;
        color: white;
    }

    .subtitle {
        color: #a1a1aa;
        font-size: 15px;
        margin-bottom: 25px;
    }

    .section-title {
        font-size: 26px;
        font-weight: 800;
        color: white;
        margin-top: 18px;
        margin-bottom: 15px;
    }

    .project-card {
        background: #1f2430;
        border: 1px solid #30384a;
        border-radius: 18px;
        padding: 22px;
        margin-bottom: 18px;
        box-shadow: 0 6px 18px rgba(0,0,0,0.25);
    }

    .project-title {
        font-size: 22px;
        font-weight: 800;
        color: white;
        margin-bottom: 14px;
        line-height: 1.35;
    }

    .project-text {
        font-size: 15px;
        color: #d4d4d8;
        margin-bottom: 6px;
        line-height: 1.4;
    }

    .tag {
        display: inline-block;
        padding: 6px 11px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
        margin-right: 6px;
        margin-top: 12px;
    }

    .tag-completo {
        background-color: #123d2a;
        color: #67e8a3;
    }

    .tag-proceso {
        background-color: #3d3312;
        color: #ffd166;
    }

    .tag-planificado {
        background-color: #13294b;
        color: #83b9ff;
    }

    .tag-stop {
        background-color: #4a1515;
        color: #ff8a8a;
    }

    .tag-neutro {
        background-color: #3a3a3a;
        color: #d0d0d0;
    }

    .note {
        color: #a1a1aa;
        font-size: 14px;
        margin-bottom: 15px;
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


def find_column(df, names):
    normalized_cols = {normalize_col(c): c for c in df.columns}

    for name in names:
        key = normalize_col(name)
        if key in normalized_cols:
            return normalized_cols[key]

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


def es_completo(estado):
    estado = str(estado).lower()
    return "completo" in estado or "finalizado" in estado


def estado_class(estado):
    estado = str(estado).lower()

    if "completo" in estado or "finalizado" in estado:
        return "tag-completo"

    if "proceso" in estado:
        return "tag-proceso"

    if "planificado" in estado:
        return "tag-planificado"

    if "stop" in estado or "vencido" in estado or "atrasado" in estado:
        return "tag-stop"

    return "tag-neutro"


def vencimiento_texto(fecha, estado):
    if es_completo(estado):
        return "Completado", "tag-completo"

    if pd.isna(fecha):
        return "Sin fecha", "tag-neutro"

    hoy = pd.Timestamp(date.today()).normalize()
    fecha = pd.Timestamp(fecha).normalize()
    dias = (fecha - hoy).days

    if dias < 0:
        return f"Vencido hace {abs(dias)} días", "tag-stop"

    if dias == 0:
        return "Vence hoy", "tag-stop"

    if dias <= 7:
        return f"Vence en {dias} días", "tag-stop"

    if dias <= 30:
        return f"Vence en {dias} días", "tag-proceso"

    return f"Vence en {dias} días", "tag-neutro"


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
        st.error(f"No se encontró la pestaña '{EXCEL_SHEET}' en el Excel.")
        st.write("Pestañas encontradas:", excel_file.sheet_names)
        st.stop()

    df = pd.read_excel(excel_file, sheet_name=sheet_found, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    return df


@st.cache_data(ttl=300, show_spinner=False)
def load_excel_from_github(url):
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
    except Exception as e:
        st.error(f"No se pudo descargar el Excel desde GitHub. Error: {e}")
        st.stop()

    excel_bytes = BytesIO(response.content)
    return read_excel_base(excel_bytes)


def prepare_data(df):
    col_anio = find_column(df, ["AÑO", "ANIO"])
    col_tipo = find_column(df, ["TIPO"])
    col_proyecto = find_column(df, ["PROYECTO"])
    col_mini = find_column(df, ["MINI PROYECTO", "MINIPROYECTO", "MINI_PROYECTO"])
    col_descripcion = find_column(df, ["DESCRIPCIÓN", "DESCRIPCION"])
    col_encargado = find_column(df, ["ENCARGADO", "ENCARGA", "RESPONSABLE"])
    col_equipo = find_column(df, ["EQUIPO DE SOPORTE", "EQUIPO DE SOP", "EQUIPO"])
    col_estado = find_column(df, ["ESTADO"])
    col_avance = find_column(df, ["% COMPL", "% COMPLE", "% COMPLETADO", "AVANCE", "PORCENTAJE"])

    col_fecha = find_column(df, [
        "FECHA_ENTREGA",
        "FECHA ENTREGA",
        "FECHA DE ENTREGA",
        "FECHA FINALIZACIÓN",
        "FECHA FINALIZACION",
        "FECHA DE FINALIZACIÓN",
        "FECHA DE FINALIZACION",
        "VENCIMIENTO",
        "FECHA VENCIMIENTO",
        "FECHA DE VENCIMIENTO"
    ])

    required = {
        "TIPO": col_tipo,
        "PROYECTO": col_proyecto,
        "MINI PROYECTO": col_mini,
        "ESTADO": col_estado,
        "FECHA_ENTREGA": col_fecha
    }

    missing = [name for name, col in required.items() if col is None]

    if missing:
        st.error("Faltan columnas obligatorias en la pestaña BASE: " + ", ".join(missing))
        st.write("Columnas encontradas:", list(df.columns))
        st.stop()

    data = pd.DataFrame()

    data["Año"] = df[col_anio] if col_anio else ""
    data["Tipo"] = df[col_tipo].fillna("Sin tipo").astype(str).str.strip()
    data["Proyecto"] = df[col_proyecto].fillna("").astype(str).str.strip()
    data["Mini proyecto"] = df[col_mini].fillna("").astype(str).str.strip()

    if col_descripcion:
        data["Descripción"] = df[col_descripcion].fillna("").astype(str).str.strip()
    else:
        data["Descripción"] = ""

    if col_encargado:
        data["Encargado"] = df[col_encargado].fillna("").astype(str).str.strip()
    else:
        data["Encargado"] = ""

    if col_equipo:
        data["Equipo"] = df[col_equipo].fillna("").astype(str).str.strip()
    else:
        data["Equipo"] = ""

    data["Estado"] = df[col_estado].fillna("Sin estado").astype(str).str.strip()

    if col_avance:
        data["Avance"] = df[col_avance].apply(clean_percentage)
    else:
        data["Avance"] = 0

    data["Fecha entrega"] = pd.to_datetime(
        df[col_fecha],
        errors="coerce",
        dayfirst=True
    )

    hoy = pd.Timestamp(date.today()).normalize()

    data["Días restantes"] = (data["Fecha entrega"] - hoy).dt.days
    data["Sin fecha"] = data["Fecha entrega"].isna()
    data["Fecha texto"] = data["Fecha entrega"].dt.strftime("%d/%m/%Y")
    data["Fecha texto"] = data["Fecha texto"].fillna("Sin fecha")

    data["Nombre tarjeta"] = data["Mini proyecto"]
    data.loc[data["Nombre tarjeta"].str.strip() == "", "Nombre tarjeta"] = data["Proyecto"]

    data["Es completo"] = data["Estado"].apply(es_completo)
    data["Es pendiente"] = ~data["Es completo"]

    data = data.sort_values(
        by=["Sin fecha", "Fecha entrega"],
        ascending=[True, True]
    )

    return data


def plotly_layout(fig):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="#0f1117",
        plot_bgcolor="#0f1117",
        font=dict(color="#f4f4f5"),
        margin=dict(l=20, r=20, t=60, b=20),
        legend=dict(
            bgcolor="rgba(0,0,0,0)",
            font=dict(color="#f4f4f5")
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


def render_card(row):
    titulo = escape(str(row["Nombre tarjeta"]))
    tipo = escape(str(row["Tipo"]))
    proyecto = escape(str(row["Proyecto"]))
    descripcion = escape(str(row["Descripción"]))
    encargado = escape(str(row["Encargado"]))
    equipo = escape(str(row["Equipo"]))
    estado = escape(str(row["Estado"]))
    fecha_texto = escape(str(row["Fecha texto"]))

    avance = int(row["Avance"])

    venc_text, venc_class = vencimiento_texto(row["Fecha entrega"], row["Estado"])
    estado_css = estado_class(row["Estado"])

    venc_text = escape(str(venc_text))

    descripcion_html = ""
    if descripcion.strip():
        descripcion_html = f'<div class="project-text"><b>Descripción:</b> {descripcion}</div>'

    encargado_html = ""
    if encargado.strip():
        encargado_html = f'<div class="project-text"><b>Encargado:</b> {encargado}</div>'

    equipo_html = ""
    if equipo.strip():
        equipo_html = f'<div class="project-text"><b>Equipo soporte:</b> {equipo}</div>'

    html = (
        '<div class="project-card">'
        f'<div class="project-title">📌 {titulo}</div>'
        f'<div class="project-text"><b>Tipo:</b> {tipo}</div>'
        f'<div class="project-text"><b>Proyecto:</b> {proyecto}</div>'
        f'{descripcion_html}'
        f'{encargado_html}'
        f'{equipo_html}'
        f'<div class="project-text"><b>Fecha entrega:</b> {fecha_texto}</div>'
        f'<span class="tag {estado_css}">{estado}</span>'
        f'<span class="tag {venc_class}">{venc_text}</span>'
        f'<div class="project-text" style="margin-top:14px;"><b>Avance:</b> {avance}%</div>'
        '</div>'
    )

    st.markdown(html, unsafe_allow_html=True)
    st.progress(avance / 100)


# =====================================================
# APP
# =====================================================

st.markdown(
    '<div class="main-title">📌 Tracker IDM</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">Dashboard interactivo conectado al Excel de GitHub, pestaña BASE.</div>',
    unsafe_allow_html=True
)


# =====================================================
# SIDEBAR
# =====================================================

with st.sidebar:
    st.header("⚙️ Configuración")

    uploaded_file = st.file_uploader(
        "Subir Excel manualmente",
        type=["xlsx", "xls"]
    )

    if st.button("🔄 Actualizar datos"):
        st.cache_data.clear()
        st.rerun()

    st.caption("Fuente GitHub:")
    st.code(EXCEL_URL)

    st.caption("Pestaña leída:")
    st.code(EXCEL_SHEET)


# =====================================================
# CARGA DE DATOS
# =====================================================

if uploaded_file is not None:
    raw_df = read_excel_base(uploaded_file)
    fuente = "Archivo subido manualmente"
else:
    raw_df = load_excel_from_github(EXCEL_URL)
    fuente = "GitHub"

df = prepare_data(raw_df)

st.caption(f"Fuente de datos: {fuente} | Pestaña: {EXCEL_SHEET}")


# =====================================================
# FILTROS
# =====================================================

st.markdown('<div class="section-title">Filtros</div>', unsafe_allow_html=True)

f1, f2, f3, f4 = st.columns(4)

with f1:
    tipos_disponibles = sorted(df["Tipo"].dropna().unique())

    filtro_tipo = st.multiselect(
        "Tipo",
        options=tipos_disponibles,
        default=tipos_disponibles
    )

with f2:
    estados_disponibles = sorted(df["Estado"].dropna().unique())

    filtro_estado = st.multiselect(
        "Estado",
        options=estados_disponibles,
        default=estados_disponibles
    )

with f3:
    encargados_disponibles = sorted([
        x for x in df["Encargado"].dropna().unique()
        if str(x).strip() != ""
    ])

    filtro_encargado = st.multiselect(
        "Encargado",
        options=encargados_disponibles,
        default=encargados_disponibles
    )

with f4:
    buscar = st.text_input("Buscar")


df_f = df.copy()

df_f = df_f[df_f["Tipo"].isin(filtro_tipo)]
df_f = df_f[df_f["Estado"].isin(filtro_estado)]

if encargados_disponibles:
    df_f = df_f[df_f["Encargado"].isin(filtro_encargado)]

if buscar:
    buscar_lower = buscar.lower()

    df_f = df_f[
        df_f["Proyecto"].astype(str).str.lower().str.contains(buscar_lower, na=False)
        | df_f["Mini proyecto"].astype(str).str.lower().str.contains(buscar_lower, na=False)
        | df_f["Descripción"].astype(str).str.lower().str.contains(buscar_lower, na=False)
        | df_f["Encargado"].astype(str).str.lower().str.contains(buscar_lower, na=False)
    ]


# =====================================================
# FILTRO POR FECHA
# =====================================================

fechas_validas = df_f["Fecha entrega"].dropna()

if not fechas_validas.empty:
    min_fecha = fechas_validas.min().date()
    max_fecha = fechas_validas.max().date()

    if min_fecha != max_fecha:
        rango = st.slider(
            "Rango de fecha de entrega",
            min_value=min_fecha,
            max_value=max_fecha,
            value=(min_fecha, max_fecha),
            format="DD/MM/YYYY"
        )

        incluir_sin_fecha = st.checkbox("Incluir proyectos sin fecha", value=True)

        inicio = pd.Timestamp(rango[0])
        fin = pd.Timestamp(rango[1])

        if incluir_sin_fecha:
            df_f = df_f[
                df_f["Fecha entrega"].isna()
                | (
                    (df_f["Fecha entrega"] >= inicio)
                    & (df_f["Fecha entrega"] <= fin)
                )
            ]
        else:
            df_f = df_f[
                (df_f["Fecha entrega"] >= inicio)
                & (df_f["Fecha entrega"] <= fin)
            ]
    else:
        st.caption(f"Todos los proyectos con fecha tienen la misma fecha de entrega: {min_fecha.strftime('%d/%m/%Y')}")


if df_f.empty:
    st.warning("No hay proyectos con los filtros seleccionados.")
    st.stop()


# =====================================================
# PESTAÑAS
# =====================================================

tab1, tab2, tab3 = st.tabs(
    [
        "📊 Dashboard interactivo",
        "🗂️ Pendientes por fecha de entrega",
        "📋 Tabla"
    ]
)


# =====================================================
# TAB 1 - DASHBOARD
# =====================================================

with tab1:
    st.markdown('<div class="section-title">Resumen general</div>', unsafe_allow_html=True)

    df_pendientes = df_f[df_f["Es pendiente"]].copy()

    total = len(df_f)
    pendientes = len(df_pendientes)

    completos = df_f["Estado"].astype(str).str.lower().str.contains("completo|finalizado", na=False).sum()
    proceso = df_f["Estado"].astype(str).str.lower().str.contains("proceso", na=False).sum()
    planificado = df_f["Estado"].astype(str).str.lower().str.contains("planificado", na=False).sum()
    stop = df_f["Estado"].astype(str).str.lower().str.contains("stop", na=False).sum()

    vencidos = (
        (df_pendientes["Fecha entrega"].notna())
        & (df_pendientes["Días restantes"] < 0)
    ).sum()

    avance_promedio = round(df_f["Avance"].mean(), 1)

    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)

    k1.metric("Total", int(total))
    k2.metric("Pendientes", int(pendientes))
    k3.metric("Completos", int(completos))
    k4.metric("En proceso", int(proceso))
    k5.metric("Planificados", int(planificado))
    k6.metric("Stop", int(stop))
    k7.metric("Vencidos pendientes", int(vencidos))

    st.metric("Avance promedio", f"{avance_promedio}%")

    c1, c2 = st.columns(2)

    with c1:
        estado_df = (
            df_f.groupby("Estado", as_index=False)
            .size()
            .rename(columns={"size": "Cantidad"})
            .sort_values("Cantidad", ascending=False)
        )

        fig_estado = px.bar(
            estado_df,
            x="Estado",
            y="Cantidad",
            text="Cantidad",
            title="Proyectos por estado",
            labels={
                "Estado": "Estado",
                "Cantidad": "Cantidad de proyectos"
            }
        )

        fig_estado.update_traces(textposition="outside")
        fig_estado = plotly_layout(fig_estado)

        st.plotly_chart(
            fig_estado,
            use_container_width=True,
            config={"displayModeBar": True}
        )

    with c2:
        tipo_df = (
            df_f.groupby("Tipo", as_index=False)
            .size()
            .rename(columns={"size": "Cantidad"})
            .sort_values("Cantidad", ascending=False)
        )

        fig_tipo = px.pie(
            tipo_df,
            names="Tipo",
            values="Cantidad",
            title="Proyectos por tipo",
            hole=0.45
        )

        fig_tipo = plotly_layout(fig_tipo)

        st.plotly_chart(
            fig_tipo,
            use_container_width=True,
            config={"displayModeBar": True}
        )

    st.markdown('<div class="section-title">Línea de vencimientos pendientes</div>', unsafe_allow_html=True)

    df_time = df_pendientes.dropna(subset=["Fecha entrega"]).copy()
    df_time = df_time.sort_values("Fecha entrega").head(30)

    if not df_time.empty:
        fig_time = px.scatter(
            df_time,
            x="Fecha entrega",
            y="Nombre tarjeta",
            color="Estado",
            size="Avance",
            hover_data={
                "Tipo": True,
                "Proyecto": True,
                "Encargado": True,
                "Fecha texto": True,
                "Días restantes": True,
                "Fecha entrega": False,
                "Nombre tarjeta": False
            },
            title="Proyectos pendientes ordenados por fecha de entrega"
        )

        fig_time = plotly_layout(fig_time)

        st.plotly_chart(
            fig_time,
            use_container_width=True,
            config={"displayModeBar": True}
        )
    else:
        st.info("No hay proyectos pendientes con fecha de entrega para mostrar.")


# =====================================================
# TAB 2 - TARJETAS SOLO PENDIENTES
# =====================================================

with tab2:
    st.markdown(
        '<div class="section-title">Pendientes ordenados por fecha de entrega</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="note">Aquí solo aparecen proyectos pendientes. Los completos o finalizados no se muestran en esta vista.</div>',
        unsafe_allow_html=True
    )

    vista = st.radio(
        "Vista",
        ["Lista general", "Agrupado por tipo"],
        horizontal=True
    )

    df_cards = df_f[df_f["Es pendiente"]].copy()

    df_cards = df_cards.sort_values(
        by=["Sin fecha", "Fecha entrega"],
        ascending=[True, True]
    )

    if df_cards.empty:
        st.success("No hay proyectos pendientes con los filtros seleccionados.")

    elif vista == "Lista general":
        for _, row in df_cards.iterrows():
            render_card(row)

    else:
        tipos = list(df_cards["Tipo"].dropna().unique())

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
            for tipo_real in tipos:
                if str(tipo_real).strip().lower() == tipo_preferido.lower():
                    if tipo_real not in tipos_ordenados:
                        tipos_ordenados.append(tipo_real)

        for tipo_real in tipos:
            if tipo_real not in tipos_ordenados:
                tipos_ordenados.append(tipo_real)

        num_cols = min(len(tipos_ordenados), 4)

        if num_cols == 0:
            st.warning("No hay proyectos pendientes para mostrar.")
        else:
            cols = st.columns(num_cols)

            for i, tipo in enumerate(tipos_ordenados):
                with cols[i % num_cols]:
                    st.subheader(tipo)

                    subset = df_cards[df_cards["Tipo"] == tipo].sort_values(
                        by=["Sin fecha", "Fecha entrega"],
                        ascending=[True, True]
                    )

                    for _, row in subset.iterrows():
                        render_card(row)


# =====================================================
# TAB 3 - TABLA
# =====================================================

with tab3:
    st.markdown('<div class="section-title">Tabla filtrada</div>', unsafe_allow_html=True)

    tabla = df_f[
        [
            "Tipo",
            "Proyecto",
            "Mini proyecto",
            "Descripción",
            "Encargado",
            "Equipo",
            "Estado",
            "Avance",
            "Fecha texto",
            "Días restantes"
        ]
    ].copy()

    tabla = tabla.rename(
        columns={
            "Avance": "% completado",
            "Fecha texto": "Fecha entrega"
        }
    )

    st.dataframe(
        tabla,
        use_container_width=True,
        hide_index=True
    )
