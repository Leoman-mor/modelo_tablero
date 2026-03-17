"""
TABLERO DE SATISFACCIÓN DE CLIENTES - QUIMPAC DE COLOMBIA S.A.
==============================================================
Borrador de tablero gerencial y comercial.

INSTRUCCIONES DE USO:
  1. Instalar dependencias (solo la primera vez):
       pip install dash plotly dash-bootstrap-components pandas openpyxl

  2. Ejecutar el script:
       python dashboard_satisfaccion.py

  3. Abrir en el navegador:
       http://127.0.0.1:8050/
"""

import os
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import dash
from dash import dcc, html, Input, Output, dash_table
import dash_bootstrap_components as dbc

# ─────────────────────────────────────────────────────────
# 1. CARGA Y PREPARACIÓN DE DATOS
# ─────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE  = os.path.join(SCRIPT_DIR, "Consolidado_Encuestas_Satisfaccion.xlsx")

df_raw = pd.read_excel(DATA_FILE, header=1)

# Renombrar columnas para uso interno
RENAME = {
    "Año":                   "anio",
    "Tipo de Encuesta":      "tipo",
    "Empresa / Cliente":     "empresa",
    "Cargo":                 "cargo",
    "Nombre Encuestado":     "nombre",
    "Fecha Encuesta":        "fecha",
    "Producto Evaluado":     "producto",
    "Cantidad Entregada":    "c_cantidad",
    "Tiempo de Entrega":     "c_tiempo_entrega",
    "Calidad":               "c_calidad",
    "Precio":                "c_precio",
    "Apoyo Técnico":         "c_apoyo_tec",
    "Servicio al Cliente":   "c_servicio",
    "Tiempo Tránsito":       "c_tiempo_transito",
    "Doc. Logística":        "c_doc_log",
    "Servicio Técnico":      "c_servicio_tec",
    "Soporte Comercial":     "c_soporte",
    "Mercadeo":              "c_mercadeo",
    "Devoluciones/Averías":  "c_devoluciones",
    "Asesoría Vendedor":     "c_asesoria",
    "Dinámicas Comerciales": "c_dinamicas",
    "Comentario Criterios":  "comentario_crit",
    "NPS (0-10)":            "nps",
    "NPS Comentario":        "nps_comentario",
    "Tiene Quejas":          "tiene_quejas",
    "Razón Quejas":          "razon_queja",
    "Especificación Queja":  "especif_queja",
    "Gestión Queja":         "gestion_queja",
    "Aspectos a Mejorar":    "aspectos_mejorar",
    "Productos Otra Cía":    "prods_otra",
    "Razón Compra Otra":     "razon_otra",
}
df = df_raw.rename(columns=RENAME)
df["nps"] = pd.to_numeric(df["nps"], errors="coerce")
df["anio"] = df["anio"].astype(int)

# Mapa score → valor numérico para promedios
SCORE_NUM = {"Excelente": 4, "Bueno": 3, "Regular": 2, "Malo": 1}

CRITERIOS_LABELS = {
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

# Colores por tipo de encuesta
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

# ─────────────────────────────────────────────────────────
# 2. HELPERS
# ─────────────────────────────────────────────────────────
def calcular_nps_cat(nps_val):
    if pd.isna(nps_val): return "Sin dato"
    if nps_val >= 9:     return "Promotor"
    if nps_val >= 7:     return "Neutro"
    return "Detractor"

def nps_score(serie):
    """Calcula NPS = %Promotores - %Detractores."""
    cats  = serie.dropna().apply(calcular_nps_cat)
    total = len(cats)
    if total == 0: return 0
    prom = (cats == "Promotor").sum()  / total * 100
    detr = (cats == "Detractor").sum() / total * 100
    return round(prom - detr, 1)

def score_numerico(serie):
    return serie.map(SCORE_NUM).mean()

def kpi_card(titulo, valor, subtitulo="", color="#1F4E79", icon="📊"):
    return dbc.Card([
        dbc.CardBody([
            html.P(f"{icon} {titulo}", className="kpi-title"),
            html.H3(valor, className="kpi-value", style={"color": color}),
            html.P(subtitulo, className="kpi-sub"),
        ])
    ], className="kpi-card")

# ─────────────────────────────────────────────────────────
# 3. APP DASH
# ─────────────────────────────────────────────────────────
app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.FLATLY],
    title="Satisfacción Clientes · QUIMPAC",
    suppress_callback_exceptions=True,
)

ANIOS   = sorted(df["anio"].unique())
TIPOS   = ["Todos"] + sorted(df["tipo"].unique())

# ─── LAYOUT ───────────────────────────────────────────────
app.layout = dbc.Container([

    # HEADER
    dbc.Row([
        dbc.Col([
            html.Div([
                html.H2("📋 Satisfacción de Clientes", style={"margin": 0, "color": "white"}),
                html.P("QUIMPAC DE COLOMBIA S.A. · Tablero Gerencial y Comercial",
                       style={"margin": 0, "color": "#BDD7EE", "fontSize": "13px"}),
            ])
        ], width=8),
        dbc.Col([
            html.Div([
                html.Span("Encuestas 2024–2025", style={"color": "#BDD7EE", "fontSize": "12px"}),
            ], style={"textAlign": "right", "paddingTop": "12px"})
        ], width=4),
    ], className="header-bar"),

    # FILTROS
    dbc.Row([
        dbc.Col([
            html.Label("Año de evaluación", className="filter-label"),
            dcc.Checklist(
                id="filtro-anio",
                options=[{"label": f" {a}", "value": a} for a in ANIOS],
                value=ANIOS,
                inline=True,
                className="filter-checklist",
            ),
        ], width=3),
        dbc.Col([
            html.Label("Tipo de encuesta", className="filter-label"),
            dcc.Dropdown(
                id="filtro-tipo",
                options=[{"label": t, "value": t} for t in TIPOS],
                value="Todos",
                clearable=False,
                style={"fontSize": "13px"},
            ),
        ], width=3),
        dbc.Col([
            html.Label("Empresa / Cliente", className="filter-label"),
            dcc.Dropdown(
                id="filtro-empresa",
                options=[{"label": "Todas", "value": "Todas"}],
                value="Todas",
                clearable=False,
                style={"fontSize": "13px"},
            ),
        ], width=4),
        dbc.Col([
            html.Label(" ", className="filter-label"),
            dbc.Button("🔄 Limpiar filtros", id="btn-reset", color="secondary",
                       size="sm", className="mt-1", style={"width": "100%"}),
        ], width=2),
    ], className="filter-bar"),

    # KPIs
    dbc.Row(id="row-kpis", className="kpi-row"),

    html.Hr(style={"margin": "8px 0"}),

    # FILA 1: NPS + Distribución scores
    dbc.Row([
        dbc.Col(dcc.Graph(id="chart-nps-tipo", config={"displayModeBar": False}), width=5),
        dbc.Col(dcc.Graph(id="chart-scores-dist", config={"displayModeBar": False}), width=4),
        dbc.Col(dcc.Graph(id="chart-nps-gauge", config={"displayModeBar": False}), width=3),
    ]),

    # FILA 2: Criterios + Tendencia
    dbc.Row([
        dbc.Col(dcc.Graph(id="chart-criterios", config={"displayModeBar": False}), width=7),
        dbc.Col(dcc.Graph(id="chart-tendencia", config={"displayModeBar": False}), width=5),
    ]),

    # FILA 3: Quejas + Aspectos a mejorar
    dbc.Row([
        dbc.Col(dcc.Graph(id="chart-quejas-razones", config={"displayModeBar": False}), width=5),
        dbc.Col(dcc.Graph(id="chart-aspectos", config={"displayModeBar": False}), width=4),
        dbc.Col(dcc.Graph(id="chart-quejas-tipo", config={"displayModeBar": False}), width=3),
    ]),

    # FILA 4: Tabla de atención prioritaria
    dbc.Row([
        dbc.Col([
            html.H6("⚠️ Clientes con Calificaciones Regular o Malo — Atención Prioritaria",
                    className="section-title"),
            html.Div(id="tabla-criticos"),
        ], width=12),
    ], className="mt-2"),

    html.Div(style={"height": "30px"}),

], fluid=True, className="main-container")


# ─── CALLBACKS ────────────────────────────────────────────
def filtrar_df(anios, tipo, empresa):
    dff = df[df["anio"].isin(anios)] if anios else df.copy()
    if tipo != "Todos":
        dff = dff[dff["tipo"] == tipo]
    if empresa != "Todas":
        dff = dff[dff["empresa"] == empresa]
    return dff


@app.callback(
    Output("filtro-empresa", "options"),
    Output("filtro-empresa", "value"),
    Input("filtro-anio", "value"),
    Input("filtro-tipo", "value"),
    Input("btn-reset", "n_clicks"),
)
def update_empresas(anios, tipo, _):
    dff = df[df["anio"].isin(anios or ANIOS)]
    if tipo and tipo != "Todos":
        dff = dff[dff["tipo"] == tipo]
    empresas = sorted(dff["empresa"].dropna().unique())
    opts = [{"label": "Todas", "value": "Todas"}] + [{"label": e, "value": e} for e in empresas]
    return opts, "Todas"


@app.callback(
    Output("filtro-anio", "value"),
    Output("filtro-tipo", "value"),
    Input("btn-reset", "n_clicks"),
    prevent_initial_call=True,
)
def reset_filters(_):
    return ANIOS, "Todos"


@app.callback(
    Output("row-kpis", "children"),
    Output("chart-nps-tipo", "figure"),
    Output("chart-scores-dist", "figure"),
    Output("chart-nps-gauge", "figure"),
    Output("chart-criterios", "figure"),
    Output("chart-tendencia", "figure"),
    Output("chart-quejas-razones", "figure"),
    Output("chart-aspectos", "figure"),
    Output("chart-quejas-tipo", "figure"),
    Output("tabla-criticos", "children"),
    Input("filtro-anio", "value"),
    Input("filtro-tipo", "value"),
    Input("filtro-empresa", "value"),
)
def update_all(anios, tipo, empresa):
    anios = anios or ANIOS
    dff = filtrar_df(anios, tipo, empresa)

    if dff.empty:
        empty = go.Figure()
        empty.add_annotation(text="Sin datos para los filtros seleccionados",
                             showarrow=False, font=dict(size=14))
        return [], empty, empty, empty, empty, empty, empty, empty, empty, html.P("Sin datos.")

    # ── KPIs ──────────────────────────────────────────────
    total    = len(dff)
    nps_prom = round(dff["nps"].mean(), 1)
    nps_sc   = nps_score(dff["nps"])
    pct_exc  = round((dff["c_calidad"] == "Excelente").sum() / dff["c_calidad"].notna().sum() * 100, 1)
    pct_quej = round((dff["tiene_quejas"] == "Si").sum() / total * 100, 1)
    tipos_n  = dff["tipo"].nunique()

    kpis = dbc.Row([
        dbc.Col(kpi_card("Total Encuestas", f"{total:,}", f"{tipos_n} tipos · {len(anios)} año(s)", "#1F4E79", "📋"), width=2),
        dbc.Col(kpi_card("NPS Promedio", f"{nps_prom}", "Escala 0 – 10", "#2E75B6" if nps_prom >= 9 else "#ED7D31", "⭐"), width=2),
        dbc.Col(kpi_card("NPS Score", f"{nps_sc:+.0f}%", "Promotores − Detractores", "#375623" if nps_sc >= 50 else "#ED7D31", "📈"), width=2),
        dbc.Col(kpi_card("Calidad Excelente", f"{pct_exc}%", "De respuestas en Calidad", "#375623", "✅"), width=2),
        dbc.Col(kpi_card("Con Quejas", f"{pct_quej}%", f"{(dff['tiene_quejas']=='Si').sum()} clientes", "#C00000" if pct_quej > 20 else "#ED7D31", "⚠️"), width=2),
        dbc.Col(kpi_card("Recomienda Producto", f"{round((dff['nps'] >= 9).sum()/total*100,1)}%", "NPS ≥ 9 (Promotores)", "#2E75B6", "👍"), width=2),
    ])

    # ── CHART 1: NPS por tipo (barras horizontales) ───────
    nps_tipo = (
        dff.groupby("tipo")["nps"].mean().round(1)
        .reset_index().sort_values("nps", ascending=True)
    )
    fig_nps_tipo = go.Figure()
    for _, row in nps_tipo.iterrows():
        color = TIPO_COLORS.get(row["tipo"], "#2E75B6")
        fig_nps_tipo.add_trace(go.Bar(
            y=[row["tipo"]], x=[row["nps"]],
            orientation="h",
            marker_color=color,
            text=[f"{row['nps']}"],
            textposition="inside",
            textfont=dict(color="white", size=11, family="Arial"),
            showlegend=False,
            name=row["tipo"],
        ))
    fig_nps_tipo.update_layout(
        title=dict(text="NPS Promedio por Tipo de Encuesta", font=dict(size=13)),
        xaxis=dict(title="NPS (0-10)", range=[0, 10.5], tickfont=dict(size=11)),
        yaxis=dict(tickfont=dict(size=11)),
        height=260, margin=dict(l=10, r=20, t=40, b=30),
        plot_bgcolor="white", paper_bgcolor="white",
        shapes=[dict(type="line", x0=9, x1=9, y0=-0.5, y1=len(nps_tipo)-0.5,
                     line=dict(color="#375623", width=2, dash="dash"))],
        annotations=[dict(x=9.05, y=len(nps_tipo)-0.5, xanchor="left",
                          text="Meta 9", font=dict(color="#375623", size=10), showarrow=False)],
    )

    # ── CHART 2: Distribución de scores (donut) ───────────
    crit_cols = [c for c in CRITERIOS_LABELS if c in dff.columns]
    all_scores = pd.concat([dff[c].dropna() for c in crit_cols])
    score_counts = all_scores.value_counts().reset_index()
    score_counts.columns = ["score", "count"]
    order = ["Excelente", "Bueno", "Regular", "Malo", "No sabe/ No aplica"]
    score_counts["score"] = pd.Categorical(score_counts["score"], categories=order, ordered=True)
    score_counts = score_counts.sort_values("score")

    fig_dist = go.Figure(go.Pie(
        labels=score_counts["score"],
        values=score_counts["count"],
        hole=0.55,
        marker_colors=[SCORE_COLORS.get(s, "#999") for s in score_counts["score"]],
        textinfo="percent",
        textfont=dict(size=11),
        hovertemplate="%{label}: %{value} (%{percent})<extra></extra>",
    ))
    fig_dist.update_layout(
        title=dict(text="Distribución de Calificaciones", font=dict(size=13)),
        legend=dict(orientation="v", x=1.0, y=0.5, font=dict(size=10)),
        height=260, margin=dict(l=10, r=10, t=40, b=10),
        paper_bgcolor="white",
    )

    # ── CHART 3: Gauge NPS Score ──────────────────────────
    nps_sc_val = nps_score(dff["nps"])
    fig_gauge = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=nps_sc_val,
        delta={"reference": 50, "valueformat": ".0f"},
        number={"suffix": "%", "font": {"size": 22}},
        title={"text": "NPS Score<br><span style='font-size:11px'>Promotores − Detractores</span>",
               "font": {"size": 13}},
        gauge={
            "axis": {"range": [-100, 100], "tickwidth": 1, "tickcolor": "#595959",
                     "tickfont": {"size": 9}},
            "bar": {"color": "#2E75B6"},
            "steps": [
                {"range": [-100, 0],  "color": "#FFCCCC"},
                {"range": [0, 50],    "color": "#FFE699"},
                {"range": [50, 100],  "color": "#C6EFCE"},
            ],
            "threshold": {"line": {"color": "#375623", "width": 3},
                          "thickness": 0.75, "value": 50},
        },
    ))
    fig_gauge.update_layout(height=260, margin=dict(l=20, r=20, t=40, b=10), paper_bgcolor="white")

    # ── CHART 4: Criterios por tipo (barras agrupadas) ────
    crit_disponibles = {k: v for k, v in CRITERIOS_LABELS.items() if k in dff.columns}
    records = []
    for col, label in crit_disponibles.items():
        for t in dff["tipo"].unique():
            sub = dff[dff["tipo"] == t][col].dropna()
            if len(sub) == 0:
                continue
            pct_exc_c = (sub == "Excelente").sum() / len(sub) * 100
            pct_bue_c = (sub == "Bueno").sum()     / len(sub) * 100
            pct_reg_c = (sub == "Regular").sum()   / len(sub) * 100
            pct_mal_c = (sub == "Malo").sum()       / len(sub) * 100
            records.append({"tipo": t, "criterio": label,
                            "Excelente": pct_exc_c, "Bueno": pct_bue_c,
                            "Regular": pct_reg_c,   "Malo": pct_mal_c,
                            "score_num": score_numerico(sub)})

    df_crit = pd.DataFrame(records)

    # Promedio general por criterio (todas las filas)
    crit_avg = []
    for col, label in crit_disponibles.items():
        sub = dff[col].dropna()
        if len(sub) == 0: continue
        pct_exc_c = (sub == "Excelente").sum() / len(sub) * 100
        pct_bue_c = (sub == "Bueno").sum()     / len(sub) * 100
        pct_reg_c = (sub == "Regular").sum()   / len(sub) * 100
        pct_mal_c = (sub == "Malo").sum()       / len(sub) * 100
        crit_avg.append({"criterio": label, "Excelente": pct_exc_c,
                         "Bueno": pct_bue_c, "Regular": pct_reg_c, "Malo": pct_mal_c})

    df_ca = pd.DataFrame(crit_avg).sort_values("Excelente", ascending=True)

    fig_crit = go.Figure()
    for score_name, color in [("Excelente","#375623"), ("Bueno","#70AD47"),
                               ("Regular","#ED7D31"), ("Malo","#C00000")]:
        if score_name in df_ca.columns:
            fig_crit.add_trace(go.Bar(
                y=df_ca["criterio"], x=df_ca[score_name],
                name=score_name, orientation="h",
                marker_color=color,
                text=[f"{v:.0f}%" if v > 3 else "" for v in df_ca[score_name]],
                textposition="inside", textfont=dict(color="white", size=9),
            ))
    fig_crit.update_layout(
        barmode="stack",
        title=dict(text="Calificación por Criterio (% respuestas)", font=dict(size=13)),
        xaxis=dict(title="%", ticksuffix="%", range=[0, 100]),
        yaxis=dict(tickfont=dict(size=10)),
        legend=dict(orientation="h", y=-0.15, font=dict(size=10)),
        height=370, margin=dict(l=10, r=10, t=40, b=50),
        plot_bgcolor="white", paper_bgcolor="white",
    )

    # ── CHART 5: Tendencia NPS 2024 vs 2025 ──────────────
    if len(anios) > 1:
        tend = df.groupby(["anio", "tipo"])["nps"].mean().round(1).reset_index()
        tend = tend[tend["tipo"].isin(dff["tipo"].unique())]
    else:
        tend = dff.groupby(["tipo"])["nps"].mean().round(1).reset_index()
        tend["anio"] = anios[0]

    fig_tend = go.Figure()
    for t in tend["tipo"].unique():
        sub = tend[tend["tipo"] == t].sort_values("anio")
        color = TIPO_COLORS.get(t, "#595959")
        fig_tend.add_trace(go.Scatter(
            x=sub["anio"].astype(str), y=sub["nps"],
            mode="lines+markers+text",
            name=t,
            line=dict(color=color, width=2),
            marker=dict(size=8, color=color),
            text=[f"{v}" for v in sub["nps"]],
            textposition="top center",
            textfont=dict(size=10),
        ))
    fig_tend.add_hline(y=9, line_dash="dash", line_color="#375623",
                       annotation_text="Meta 9", annotation_position="top right",
                       annotation_font_color="#375623")
    fig_tend.update_layout(
        title=dict(text="Evolución NPS por Año y Tipo", font=dict(size=13)),
        xaxis=dict(title="Año"),
        yaxis=dict(title="NPS Promedio", range=[7, 10.5]),
        legend=dict(orientation="v", x=1.0, font=dict(size=9)),
        height=300, margin=dict(l=10, r=10, t=40, b=30),
        plot_bgcolor="white", paper_bgcolor="white",
    )

    # ── CHART 6: Razones de quejas ────────────────────────
    quejas = dff[dff["tiene_quejas"] == "Si"]["razon_queja"].dropna()
    # Normalizar razones similares
    quejas = quejas.str.strip().str.title()
    qc = quejas.value_counts().head(8).reset_index()
    qc.columns = ["razon", "count"]

    fig_quejas = go.Figure(go.Bar(
        x=qc["count"], y=qc["razon"],
        orientation="h",
        marker_color="#C00000",
        text=qc["count"], textposition="outside",
        textfont=dict(size=10),
    ))
    fig_quejas.update_layout(
        title=dict(text="Principales Razones de Quejas", font=dict(size=13)),
        xaxis=dict(title="N° de quejas"),
        yaxis=dict(tickfont=dict(size=10)),
        height=310, margin=dict(l=10, r=30, t=40, b=30),
        plot_bgcolor="white", paper_bgcolor="white",
    )

    # ── CHART 7: Aspectos a mejorar ──────────────────────
    asp = dff["aspectos_mejorar"].dropna().str.strip().str.title()
    asp_c = asp.value_counts().head(6).reset_index()
    asp_c.columns = ["aspecto", "count"]
    asp_c["pct"] = (asp_c["count"] / asp_c["count"].sum() * 100).round(1)

    fig_asp = go.Figure(go.Bar(
        x=asp_c["aspecto"], y=asp_c["count"],
        marker_color=["#ED7D31", "#FFC000", "#2E75B6", "#70AD47", "#9E3EA8", "#595959"],
        text=[f"{v}%" for v in asp_c["pct"]], textposition="outside",
        textfont=dict(size=11),
    ))
    fig_asp.update_layout(
        title=dict(text="Aspectos a Mejorar", font=dict(size=13)),
        xaxis=dict(tickfont=dict(size=10)),
        yaxis=dict(title="Menciones"),
        height=310, margin=dict(l=10, r=10, t=40, b=40),
        plot_bgcolor="white", paper_bgcolor="white",
    )

    # ── CHART 8: % Quejas por tipo de encuesta ────────────
    qt = dff.groupby("tipo").apply(
        lambda x: pd.Series({
            "Si": (x["tiene_quejas"] == "Si").sum(),
            "No": (x["tiene_quejas"] == "No").sum(),
        })
    ).reset_index()

    fig_qt = go.Figure()
    fig_qt.add_trace(go.Bar(
        x=qt["tipo"], y=qt["Si"],
        name="Con queja", marker_color="#C00000",
        text=qt["Si"], textposition="inside", textfont=dict(color="white", size=11),
    ))
    fig_qt.add_trace(go.Bar(
        x=qt["tipo"], y=qt["No"],
        name="Sin queja", marker_color="#70AD47",
        text=qt["No"], textposition="inside", textfont=dict(color="white", size=11),
    ))
    fig_qt.update_layout(
        barmode="stack",
        title=dict(text="Quejas por Tipo", font=dict(size=13)),
        xaxis=dict(tickangle=-20, tickfont=dict(size=9)),
        yaxis=dict(title="Clientes"),
        legend=dict(orientation="h", y=-0.2, font=dict(size=10)),
        height=310, margin=dict(l=10, r=10, t=40, b=50),
        plot_bgcolor="white", paper_bgcolor="white",
    )

    # ── TABLA CRÍTICOS ───────────────────────────────────
    # Registros con Regular o Malo en ALGÚN criterio
    crit_cols_all = [c for c in CRITERIOS_LABELS if c in dff.columns]
    mask_bad = dff[crit_cols_all].isin(["Regular", "Malo"]).any(axis=1)
    mask_nps = dff["nps"] < 8
    criticos = dff[mask_bad | mask_nps].copy()

    if criticos.empty:
        tabla_out = dbc.Alert("✅ No hay registros con calificaciones críticas para los filtros seleccionados.",
                              color="success")
    else:
        # Qué criterios son malos
        def get_bad_crit(row):
            bads = []
            for c, lbl in CRITERIOS_LABELS.items():
                if c in row.index and row[c] in ["Regular", "Malo"]:
                    bads.append(f"{lbl}: {row[c]}")
            return " | ".join(bads) if bads else "—"

        criticos["Criterios Críticos"] = criticos[crit_cols_all + list(CRITERIOS_LABELS.keys())].apply(
            lambda r: " | ".join([f"{CRITERIOS_LABELS[c]}: {r[c]}" for c in crit_cols_all if c in r.index and r[c] in ["Regular","Malo"]]),
            axis=1
        )

        tabla_df = criticos[[
            "tipo", "anio", "empresa", "producto", "nps",
            "Criterios Críticos", "comentario_crit", "tiene_quejas", "razon_queja"
        ]].rename(columns={
            "tipo": "Tipo", "anio": "Año", "empresa": "Empresa",
            "producto": "Producto", "nps": "NPS",
            "comentario_crit": "Comentario",
            "tiene_quejas": "¿Queja?", "razon_queja": "Razón Queja",
        }).sort_values(["NPS", "Empresa"])

        tabla_out = dash_table.DataTable(
            data=tabla_df.to_dict("records"),
            columns=[{"name": c, "id": c} for c in tabla_df.columns],
            page_size=8,
            sort_action="native",
            filter_action="native",
            style_table={"overflowX": "auto"},
            style_cell={
                "fontFamily": "Arial", "fontSize": "11px",
                "padding": "5px 8px", "textAlign": "left",
                "maxWidth": "200px", "overflow": "hidden",
                "textOverflow": "ellipsis",
            },
            style_header={
                "backgroundColor": "#1F4E79", "color": "white",
                "fontWeight": "bold", "fontSize": "11px",
            },
            style_data_conditional=[
                {"if": {"filter_query": '{NPS} < 8'},
                 "backgroundColor": "#FFCCCC"},
                {"if": {"filter_query": '{¿Queja?} = "Si"'},
                 "backgroundColor": "#FCE4D6"},
                {"if": {"row_index": "odd"},
                 "backgroundColor": "#F5F5F5"},
            ],
            tooltip_data=[
                {c: {"value": str(row.get(c, "")), "type": "markdown"}
                 for c in tabla_df.columns}
                for row in tabla_df.to_dict("records")
            ],
            tooltip_duration=None,
        )

    return kpis, fig_nps_tipo, fig_dist, fig_gauge, fig_crit, fig_tend, fig_quejas, fig_asp, fig_qt, tabla_out


# ─────────────────────────────────────────────────────────
# 4. CSS INLINE
# ─────────────────────────────────────────────────────────
app.index_string = """
<!DOCTYPE html>
<html>
<head>
    {%metas%}
    <title>{%title%}</title>
    {%favicon%}
    {%css%}
    <style>
        body { background-color: #F0F4F8; font-family: Arial, sans-serif; }

        .main-container { padding: 0 12px 20px 12px; max-width: 1600px; }

        .header-bar {
            background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
            padding: 12px 20px;
            margin: 0 -12px 12px -12px;
            border-radius: 0;
        }

        .filter-bar {
            background: white;
            padding: 10px 15px;
            border-radius: 8px;
            margin-bottom: 12px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.1);
        }

        .filter-label {
            font-size: 11px;
            font-weight: bold;
            color: #595959;
            margin-bottom: 3px;
        }

        .filter-checklist label {
            font-size: 13px !important;
            margin-right: 12px !important;
        }

        .kpi-row { margin-bottom: 10px; }

        .kpi-card {
            border-radius: 8px !important;
            border: none !important;
            box-shadow: 0 2px 6px rgba(0,0,0,0.1);
            background: white !important;
            margin: 0 4px;
        }

        .kpi-card .card-body { padding: 10px 14px !important; }

        .kpi-title {
            font-size: 11px;
            color: #595959;
            font-weight: bold;
            margin-bottom: 2px;
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }

        .kpi-value {
            font-size: 26px;
            font-weight: bold;
            margin: 0;
            line-height: 1.1;
        }

        .kpi-sub {
            font-size: 10px;
            color: #8E8E8E;
            margin-bottom: 0;
        }

        .section-title {
            font-size: 13px;
            font-weight: bold;
            color: #1F4E79;
            border-left: 3px solid #2E75B6;
            padding-left: 8px;
            margin: 10px 0 6px 0;
        }

        .js-plotly-plot .plotly .modebar { display: none; }

        .dash-spreadsheet-container { font-size: 11px; }
    </style>
</head>
<body>
    {%app_entry%}
    <footer>
        {%config%}
        {%scripts%}
        {%renderer%}
    </footer>
</body>
</html>
"""

# ─────────────────────────────────────────────────────────
# 5. PUNTO DE ENTRADA
# ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 60)
    print("  TABLERO SATISFACCIÓN CLIENTES - QUIMPAC DE COLOMBIA")
    print("=" * 60)
    print(f"  Registros cargados: {len(df):,}")
    print(f"  Años: {sorted(df['anio'].unique())}")
    print(f"  Tipos: {sorted(df['tipo'].unique())}")
    print("=" * 60)
    print("  Abriendo en: http://127.0.0.1:8050/")
    print("  Presiona Ctrl+C para detener el servidor.")
    print("=" * 60)
    app.run(debug=False, host="127.0.0.1", port=8050)
