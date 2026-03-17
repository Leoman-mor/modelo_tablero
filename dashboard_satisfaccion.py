"""
TABLERO DE SATISFACCIÓN DE CLIENTES - QUIMPAC DE COLOMBIA S.A.
==============================================================
Dashboard interactivo para Streamlit Cloud.

Para correr localmente:
    streamlit run dashboard_satisfaccion.py

Para Streamlit Cloud:
    1. Subir este archivo + requirements.txt + Consolidado_Encuestas_Satisfaccion.xlsx
       al mismo repositorio de GitHub.
    2. Conectar el repositorio en share.streamlit.io
"""

import os
import warnings
warnings.filterwarnings("ignore")

import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# ─────────────────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Satisfacción Clientes · QUIMPAC",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────
# CSS PERSONALIZADO
# ─────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Fondo general */
    .stApp { background-color: #F0F4F8; }

    /* Header principal */
    .header-box {
        background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
        padding: 18px 28px;
        border-radius: 10px;
        margin-bottom: 18px;
        color: white;
    }
    .header-box h1 { margin: 0; font-size: 24px; color: white; }
    .header-box p  { margin: 4px 0 0 0; font-size: 13px; color: #BDD7EE; }

    /* Tarjetas KPI */
    .kpi-box {
        background: white;
        border-radius: 10px;
        padding: 14px 18px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        text-align: center;
        min-height: 100px;
    }
    .kpi-icon  { font-size: 22px; margin-bottom: 2px; }
    .kpi-label { font-size: 11px; color: #595959; font-weight: 700;
                 text-transform: uppercase; letter-spacing: .4px; margin: 0; }
    .kpi-value { font-size: 30px; font-weight: 800; margin: 2px 0; line-height: 1.1; }
    .kpi-sub   { font-size: 11px; color: #8E8E8E; margin: 0; }

    /* Sección title */
    .section-title {
        font-size: 13px; font-weight: 700; color: #1F4E79;
        border-left: 4px solid #2E75B6; padding-left: 8px;
        margin: 8px 0 4px 0;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] { background-color: #1F4E79; }
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] .stMarkdown p { color: white !important; }
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 { color: white !important; }

    /* Ocultar toolbar de Plotly */
    .modebar { display: none !important; }

    /* Remove padding */
    .block-container { padding-top: 1rem; padding-bottom: 2rem; }

    /* Tabla */
    .stDataFrame { font-size: 12px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    base = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base, "Consolidado_Encuestas_Satisfaccion.xlsx")
    df = pd.read_excel(path, header=1)
    df = df.rename(columns={
        "Año": "anio", "Tipo de Encuesta": "tipo",
        "Empresa / Cliente": "empresa", "Cargo": "cargo",
        "Nombre Encuestado": "nombre", "Fecha Encuesta": "fecha",
        "Producto Evaluado": "producto",
        "Cantidad Entregada": "c_cantidad",
        "Tiempo de Entrega": "c_tiempo_entrega",
        "Calidad": "c_calidad", "Precio": "c_precio",
        "Apoyo Técnico": "c_apoyo_tec",
        "Servicio al Cliente": "c_servicio",
        "Tiempo Tránsito": "c_tiempo_transito",
        "Doc. Logística": "c_doc_log",
        "Servicio Técnico": "c_servicio_tec",
        "Soporte Comercial": "c_soporte",
        "Mercadeo": "c_mercadeo",
        "Devoluciones/Averías": "c_devoluciones",
        "Asesoría Vendedor": "c_asesoria",
        "Dinámicas Comerciales": "c_dinamicas",
        "Comentario Criterios": "comentario_crit",
        "NPS (0-10)": "nps",
        "NPS Comentario": "nps_comentario",
        "Tiene Quejas": "tiene_quejas",
        "Razón Quejas": "razon_queja",
        "Especificación Queja": "especif_queja",
        "Gestión Queja": "gestion_queja",
        "Aspectos a Mejorar": "aspectos_mejorar",
        "Productos Otra Cía": "prods_otra",
        "Razón Compra Otra": "razon_otra",
    })
    df["nps"]  = pd.to_numeric(df["nps"], errors="coerce")
    df["anio"] = df["anio"].astype(int)
    return df

df_full = load_data()

# ─────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────
TIPO_COLORS = {
    "Químicos Nacional":     "#2E75B6",
    "Químicos Exportación":  "#70AD47",
    "Sal Industrial":        "#ED7D31",
    "Sales Mineralizadas":   "#9E3EA8",
}
SCORE_COLORS = {
    "Excelente": "#375623",
    "Bueno":     "#70AD47",
    "Regular":   "#ED7D31",
    "Malo":      "#C00000",
    "No sabe/ No aplica": "#BFBFBF",
}
CRITERIOS = {
    "c_cantidad":        "Cantidad Entregada",
    "c_tiempo_entrega":  "Tiempo de Entrega",
    "c_calidad":         "Calidad",
    "c_precio":          "Precio",
    "c_apoyo_tec":       "Apoyo Técnico",
    "c_servicio":        "Servicio al Cliente",
    "c_tiempo_transito": "Tiempo Tránsito",
    "c_doc_log":         "Doc. Logística",
    "c_servicio_tec":    "Servicio Técnico",
    "c_soporte":         "Soporte Comercial",
    "c_mercadeo":        "Mercadeo",
    "c_devoluciones":    "Devoluciones/Averías",
    "c_asesoria":        "Asesoría Vendedor",
    "c_dinamicas":       "Dinámicas Comerciales",
}

def nps_score(serie):
    cats  = serie.dropna().apply(lambda v: "P" if v >= 9 else ("N" if v >= 7 else "D"))
    total = len(cats)
    if total == 0: return 0
    return round((cats=="P").sum()/total*100 - (cats=="D").sum()/total*100, 1)

# ─────────────────────────────────────────────────────────
# SIDEBAR — FILTROS
# ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🎛️ Filtros")
    st.markdown("---")

    anios_disp = sorted(df_full["anio"].unique())
    anios_sel  = st.multiselect(
        "Año de evaluación",
        options=anios_disp,
        default=anios_disp,
    )

    tipos_disp = ["Todos"] + sorted(df_full["tipo"].unique())
    tipo_sel   = st.selectbox("Tipo de encuesta", tipos_disp)

    # Empresas filtradas por lo anterior
    dff_prev = df_full[df_full["anio"].isin(anios_sel or anios_disp)]
    if tipo_sel != "Todos":
        dff_prev = dff_prev[dff_prev["tipo"] == tipo_sel]
    empresas_disp = ["Todas"] + sorted(dff_prev["empresa"].dropna().unique())
    empresa_sel   = st.selectbox("Empresa / Cliente", empresas_disp)

    st.markdown("---")
    st.caption("📊 QUIMPAC DE COLOMBIA S.A.\nEncuestas 2024 – 2025")

# ─────────────────────────────────────────────────────────
# FILTRADO PRINCIPAL
# ─────────────────────────────────────────────────────────
anios_sel = anios_sel or anios_disp
dff = df_full[df_full["anio"].isin(anios_sel)]
if tipo_sel != "Todos":
    dff = dff[dff["tipo"] == tipo_sel]
if empresa_sel != "Todas":
    dff = dff[dff["empresa"] == empresa_sel]

# ─────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────
st.markdown("""
<div class="header-box">
  <h1>📋 Satisfacción de Clientes</h1>
  <p>QUIMPAC DE COLOMBIA S.A. · Tablero Gerencial y Comercial · 2024–2025</p>
</div>
""", unsafe_allow_html=True)

if dff.empty:
    st.warning("Sin datos para los filtros seleccionados.")
    st.stop()

# ─────────────────────────────────────────────────────────
# KPIs
# ─────────────────────────────────────────────────────────
total     = len(dff)
nps_prom  = round(dff["nps"].mean(), 1)
nps_sc    = nps_score(dff["nps"])
crit_cal  = dff["c_calidad"].dropna()
pct_exc   = round((crit_cal == "Excelente").sum() / len(crit_cal) * 100, 1) if len(crit_cal) else 0
n_quejas  = (dff["tiene_quejas"] == "Si").sum()
pct_quej  = round(n_quejas / total * 100, 1)
pct_prom  = round((dff["nps"] >= 9).sum() / total * 100, 1)

def kpi_html(icon, label, value, sub, color):
    return f"""
    <div class="kpi-box">
      <div class="kpi-icon">{icon}</div>
      <p class="kpi-label">{label}</p>
      <p class="kpi-value" style="color:{color};">{value}</p>
      <p class="kpi-sub">{sub}</p>
    </div>"""

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.markdown(kpi_html("📋", "Total Encuestas", f"{total:,}",
    f"{dff['tipo'].nunique()} tipos · {len(anios_sel)} año(s)", "#1F4E79"), unsafe_allow_html=True)
k2.markdown(kpi_html("⭐", "NPS Promedio", f"{nps_prom}",
    "Escala 0 – 10",
    "#375623" if nps_prom >= 9 else "#ED7D31"), unsafe_allow_html=True)
k3.markdown(kpi_html("📈", "NPS Score", f"{nps_sc:+.0f}%",
    "Promotores − Detractores",
    "#375623" if nps_sc >= 50 else "#ED7D31"), unsafe_allow_html=True)
k4.markdown(kpi_html("✅", "Calidad Excelente", f"{pct_exc}%",
    "De respuestas en Calidad", "#375623"), unsafe_allow_html=True)
k5.markdown(kpi_html("⚠️", "Con Quejas", f"{pct_quej}%",
    f"{n_quejas} clientes",
    "#C00000" if pct_quej > 20 else "#ED7D31"), unsafe_allow_html=True)
k6.markdown(kpi_html("👍", "Promotores", f"{pct_prom}%",
    "NPS ≥ 9", "#2E75B6"), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────
# FILA 1: NPS por tipo · Distribución scores · Gauge
# ─────────────────────────────────────────────────────────
col_a, col_b, col_c = st.columns([5, 4, 3])

# — NPS por tipo —
with col_a:
    st.markdown('<p class="section-title">NPS Promedio por Tipo de Encuesta</p>', unsafe_allow_html=True)
    nps_tipo = (dff.groupby("tipo")["nps"].mean().round(1)
                .reset_index().sort_values("nps", ascending=True))
    fig1 = go.Figure()
    for _, row in nps_tipo.iterrows():
        color = TIPO_COLORS.get(row["tipo"], "#2E75B6")
        fig1.add_trace(go.Bar(
            y=[row["tipo"]], x=[row["nps"]], orientation="h",
            marker_color=color,
            text=[f"  {row['nps']}"], textposition="inside",
            textfont=dict(color="white", size=12, family="Arial Bold"),
            showlegend=False,
        ))
    fig1.add_vline(x=9, line_dash="dash", line_color="#375623", line_width=2,
                   annotation_text="Meta 9", annotation_position="top",
                   annotation_font_color="#375623")
    fig1.update_layout(
        xaxis=dict(range=[0, 10.5], tickfont=dict(size=11)),
        yaxis=dict(tickfont=dict(size=11)),
        height=230, margin=dict(l=5, r=15, t=10, b=25),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    st.plotly_chart(fig1, use_container_width=True, config={"displayModeBar": False})

# — Donut distribución —
with col_b:
    st.markdown('<p class="section-title">Distribución de Calificaciones</p>', unsafe_allow_html=True)
    crit_cols = [c for c in CRITERIOS if c in dff.columns]
    all_scores = pd.concat([dff[c].dropna() for c in crit_cols])
    sc = all_scores.value_counts().reset_index()
    sc.columns = ["score", "cnt"]
    order_map = {"Excelente": 0, "Bueno": 1, "Regular": 2, "Malo": 3, "No sabe/ No aplica": 4}
    sc["ord"] = sc["score"].map(order_map).fillna(9)
    sc = sc.sort_values("ord")

    fig2 = go.Figure(go.Pie(
        labels=sc["score"], values=sc["cnt"], hole=0.55,
        marker_colors=[SCORE_COLORS.get(s, "#999") for s in sc["score"]],
        textinfo="percent", textfont=dict(size=11),
        hovertemplate="%{label}: %{value} (%{percent})<extra></extra>",
    ))
    fig2.update_layout(
        legend=dict(orientation="v", x=1.0, y=0.5, font=dict(size=10)),
        height=230, margin=dict(l=5, r=5, t=10, b=10),
        paper_bgcolor="white",
    )
    st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

# — Gauge NPS Score —
with col_c:
    st.markdown('<p class="section-title">NPS Score</p>', unsafe_allow_html=True)
    fig3 = go.Figure(go.Indicator(
        mode="gauge+number",
        value=nps_sc,
        number={"suffix": "%", "font": {"size": 28, "color": "#1F4E79"}},
        gauge={
            "axis": {"range": [-100, 100], "tickfont": {"size": 9}},
            "bar": {"color": "#2E75B6", "thickness": 0.25},
            "steps": [
                {"range": [-100, 0],  "color": "#FFCCCC"},
                {"range": [0, 50],    "color": "#FFE699"},
                {"range": [50, 100],  "color": "#C6EFCE"},
            ],
            "threshold": {"line": {"color": "#375623", "width": 3},
                          "thickness": 0.8, "value": 50},
        },
    ))
    fig3.update_layout(
        height=230, margin=dict(l=10, r=10, t=20, b=10),
        paper_bgcolor="white",
    )
    st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})

# ─────────────────────────────────────────────────────────
# FILA 2: Criterios · Tendencia
# ─────────────────────────────────────────────────────────
col_d, col_e = st.columns([7, 5])

# — Barras apiladas por criterio —
with col_d:
    st.markdown('<p class="section-title">Calificación por Criterio (% respuestas)</p>',
                unsafe_allow_html=True)
    crit_rows = []
    for col_k, label in CRITERIOS.items():
        if col_k not in dff.columns: continue
        sub = dff[col_k].dropna()
        if len(sub) == 0: continue
        row = {"criterio": label}
        for s in ["Excelente", "Bueno", "Regular", "Malo"]:
            row[s] = round((sub == s).sum() / len(sub) * 100, 1)
        crit_rows.append(row)

    df_cr = pd.DataFrame(crit_rows).sort_values("Excelente", ascending=True)

    fig4 = go.Figure()
    for score_name, color in [("Excelente","#375623"),("Bueno","#70AD47"),
                               ("Regular","#ED7D31"),("Malo","#C00000")]:
        if score_name not in df_cr.columns: continue
        fig4.add_trace(go.Bar(
            y=df_cr["criterio"], x=df_cr[score_name],
            name=score_name, orientation="h",
            marker_color=color,
            text=[f"{v:.0f}%" if v > 3 else "" for v in df_cr[score_name]],
            textposition="inside", textfont=dict(color="white", size=9),
        ))
    fig4.update_layout(
        barmode="stack",
        xaxis=dict(ticksuffix="%", range=[0, 100], tickfont=dict(size=10)),
        yaxis=dict(tickfont=dict(size=10)),
        legend=dict(orientation="h", y=-0.12, font=dict(size=10)),
        height=370, margin=dict(l=5, r=10, t=10, b=50),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar": False})

# — Tendencia NPS —
with col_e:
    st.markdown('<p class="section-title">Evolución NPS por Año y Tipo</p>',
                unsafe_allow_html=True)
    tipos_activos = dff["tipo"].unique()
    tend = df_full[df_full["tipo"].isin(tipos_activos)].groupby(["anio","tipo"])["nps"].mean().round(1).reset_index()

    fig5 = go.Figure()
    for t in sorted(tipos_activos):
        sub = tend[tend["tipo"] == t].sort_values("anio")
        color = TIPO_COLORS.get(t, "#595959")
        fig5.add_trace(go.Scatter(
            x=sub["anio"].astype(str), y=sub["nps"],
            mode="lines+markers+text", name=t,
            line=dict(color=color, width=2.5),
            marker=dict(size=9, color=color),
            text=[f"{v}" for v in sub["nps"]],
            textposition="top center", textfont=dict(size=10),
        ))
    fig5.add_hline(y=9, line_dash="dash", line_color="#375623", line_width=2,
                   annotation_text="Meta 9", annotation_position="top right",
                   annotation_font_color="#375623")
    fig5.update_layout(
        xaxis=dict(title="Año", tickfont=dict(size=11)),
        yaxis=dict(title="NPS Promedio", range=[7, 10.5]),
        legend=dict(font=dict(size=10), orientation="v", x=1.0, y=1.0),
        height=370, margin=dict(l=5, r=5, t=10, b=30),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    st.plotly_chart(fig5, use_container_width=True, config={"displayModeBar": False})

# ─────────────────────────────────────────────────────────
# FILA 3: Razones quejas · Aspectos mejorar · Quejas por tipo
# ─────────────────────────────────────────────────────────
col_f, col_g, col_h = st.columns([5, 4, 3])

# — Razones de quejas —
with col_f:
    st.markdown('<p class="section-title">Principales Razones de Quejas</p>',
                unsafe_allow_html=True)
    quejas_df = dff[dff["tiene_quejas"] == "Si"]["razon_queja"].dropna()
    if len(quejas_df) > 0:
        qc = quejas_df.str.strip().value_counts().head(8).reset_index()
        qc.columns = ["razon", "n"]
        qc = qc.sort_values("n", ascending=True)
        fig6 = go.Figure(go.Bar(
            x=qc["n"], y=qc["razon"], orientation="h",
            marker_color="#C00000",
            text=qc["n"], textposition="outside", textfont=dict(size=11),
        ))
        fig6.update_layout(
            xaxis=dict(title="N° quejas"),
            yaxis=dict(tickfont=dict(size=10)),
            height=300, margin=dict(l=5, r=30, t=10, b=30),
            plot_bgcolor="white", paper_bgcolor="white",
        )
        st.plotly_chart(fig6, use_container_width=True, config={"displayModeBar": False})
    else:
        st.success("✅ Sin quejas registradas.")

# — Aspectos a mejorar —
with col_g:
    st.markdown('<p class="section-title">Aspectos a Mejorar</p>', unsafe_allow_html=True)
    asp_s = dff["aspectos_mejorar"].dropna().str.strip().value_counts().head(6).reset_index()
    asp_s.columns = ["aspecto", "n"]
    asp_s["pct"] = (asp_s["n"] / asp_s["n"].sum() * 100).round(1)
    paleta = ["#ED7D31","#FFC000","#2E75B6","#70AD47","#9E3EA8","#595959"]

    fig7 = go.Figure(go.Bar(
        x=asp_s["aspecto"], y=asp_s["n"],
        marker_color=paleta[:len(asp_s)],
        text=[f"{p}%" for p in asp_s["pct"]],
        textposition="outside", textfont=dict(size=11),
    ))
    fig7.update_layout(
        yaxis=dict(title="Menciones"),
        xaxis=dict(tickfont=dict(size=10), tickangle=-15),
        height=300, margin=dict(l=5, r=10, t=10, b=50),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    st.plotly_chart(fig7, use_container_width=True, config={"displayModeBar": False})

# — Quejas Si/No por tipo —
with col_h:
    st.markdown('<p class="section-title">Quejas por Tipo</p>', unsafe_allow_html=True)
    qt = dff.groupby("tipo")["tiene_quejas"].value_counts().unstack(fill_value=0).reset_index()
    if "Si" not in qt.columns: qt["Si"] = 0
    if "No" not in qt.columns: qt["No"] = 0

    fig8 = go.Figure()
    fig8.add_trace(go.Bar(
        x=qt["tipo"], y=qt["Si"], name="Con queja", marker_color="#C00000",
        text=qt["Si"], textposition="inside", textfont=dict(color="white", size=11),
    ))
    fig8.add_trace(go.Bar(
        x=qt["tipo"], y=qt["No"], name="Sin queja", marker_color="#70AD47",
        text=qt["No"], textposition="inside", textfont=dict(color="white", size=11),
    ))
    fig8.update_layout(
        barmode="stack",
        xaxis=dict(tickangle=-20, tickfont=dict(size=8)),
        yaxis=dict(title="Clientes"),
        legend=dict(orientation="h", y=-0.22, font=dict(size=10)),
        height=300, margin=dict(l=5, r=5, t=10, b=55),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    st.plotly_chart(fig8, use_container_width=True, config={"displayModeBar": False})

# ─────────────────────────────────────────────────────────
# FILA 4: Tabla de atención prioritaria
# ─────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<p class="section-title">⚠️ Clientes con Atención Prioritaria — Calificaciones Regular / Malo o NPS &lt; 8</p>',
            unsafe_allow_html=True)

crit_all = [c for c in CRITERIOS if c in dff.columns]
mask_bad = dff[crit_all].isin(["Regular", "Malo"]).any(axis=1)
mask_nps = dff["nps"] < 8
criticos = dff[mask_bad | mask_nps].copy()

if criticos.empty:
    st.success("✅ No hay registros críticos para los filtros seleccionados.")
else:
    # Construir columna "Criterios Críticos"
    def get_bad(row):
        bads = []
        for c, lbl in CRITERIOS.items():
            if c in row.index and row[c] in ["Regular", "Malo"]:
                bads.append(f"{lbl}: {row[c]}")
        return " | ".join(bads) if bads else "—"

    criticos["Criterios Críticos"] = criticos.apply(get_bad, axis=1)

    tabla = criticos[[
        "tipo", "anio", "empresa", "producto", "nps",
        "Criterios Críticos", "comentario_crit",
        "tiene_quejas", "razon_queja", "aspectos_mejorar",
    ]].rename(columns={
        "tipo": "Tipo", "anio": "Año", "empresa": "Empresa",
        "producto": "Producto", "nps": "NPS",
        "comentario_crit": "Comentario",
        "tiene_quejas": "Queja", "razon_queja": "Razón Queja",
        "aspectos_mejorar": "A Mejorar",
    }).sort_values(["NPS", "Empresa"]).reset_index(drop=True)

    st.info(f"📌 {len(tabla)} registros requieren atención.")

    # Color condicional en Streamlit
    def color_row(row):
        if row["NPS"] < 8:
            bg = "background-color: #FFCCCC"
        elif row["Queja"] == "Si":
            bg = "background-color: #FCE4D6"
        else:
            bg = ""
        return [bg] * len(row)

    st.dataframe(
        tabla.style.apply(color_row, axis=1),
        use_container_width=True,
        height=320,
    )

# ─────────────────────────────────────────────────────────
# FILA 5: Competencia (si hay datos)
# ─────────────────────────────────────────────────────────
if "prods_otra" in dff.columns:
    prods_otra = dff["prods_otra"].dropna()
    prods_otra = prods_otra[prods_otra.str.strip().str.lower() != "ninguno"]
    if len(prods_otra) > 0:
        st.markdown("---")
        col_i, col_j = st.columns([6, 6])

        with col_i:
            st.markdown('<p class="section-title">Productos que Compran a la Competencia</p>',
                        unsafe_allow_html=True)
            poc = prods_otra.str.strip().value_counts().head(10).reset_index()
            poc.columns = ["producto", "n"]
            poc["pct"] = (poc["n"] / len(dff) * 100).round(1)
            fig9 = go.Figure(go.Bar(
                x=poc["n"], y=poc["producto"], orientation="h",
                marker_color="#9E3EA8",
                text=[f"{p}%" for p in poc["pct"]],
                textposition="outside", textfont=dict(size=10),
            ))
            fig9.update_layout(
                xaxis=dict(title="Menciones"),
                yaxis=dict(tickfont=dict(size=10)),
                height=280, margin=dict(l=5, r=40, t=10, b=30),
                plot_bgcolor="white", paper_bgcolor="white",
            )
            st.plotly_chart(fig9, use_container_width=True, config={"displayModeBar": False})

        if "razon_otra" in dff.columns:
            with col_j:
                st.markdown('<p class="section-title">Razón Principal para Comprar a Otra Empresa</p>',
                            unsafe_allow_html=True)
                raz = dff["razon_otra"].dropna()
                raz = raz[raz.str.strip().str.lower().isin(
                    ["n.a", "n.a.", "na", "n/a", "ninguno"]
                ) == False]
                if len(raz) > 0:
                    rc = raz.str.strip().value_counts().head(8).reset_index()
                    rc.columns = ["razon", "n"]
                    rc["pct"] = (rc["n"] / rc["n"].sum() * 100).round(1)
                    fig10 = go.Figure(go.Bar(
                        x=rc["n"], y=rc["razon"], orientation="h",
                        marker_color="#FFC000",
                        text=[f"{p}%" for p in rc["pct"]],
                        textposition="outside", textfont=dict(size=10),
                    ))
                    fig10.update_layout(
                        xaxis=dict(title="Menciones"),
                        yaxis=dict(tickfont=dict(size=10)),
                        height=280, margin=dict(l=5, r=40, t=10, b=30),
                        plot_bgcolor="white", paper_bgcolor="white",
                    )
                    st.plotly_chart(fig10, use_container_width=True, config={"displayModeBar": False})

# ─────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    f"📊 QUIMPAC DE COLOMBIA S.A. · Tablero de Satisfacción de Clientes · "
    f"{total} registros · Años: {', '.join(map(str, sorted(anios_sel)))} · "
    f"Generado automáticamente desde Consolidado_Encuestas_Satisfaccion.xlsx"
)
