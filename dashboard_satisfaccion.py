"""
TABLERO DE SATISFACCION DE CLIENTES - QUIMPAC DE COLOMBIA S.A.
Graficas estilo Power BI + Analisis de Sentimientos en espanol.
streamlit run dashboard_satisfaccion.py
"""
import os, re, warnings
from collections import Counter
warnings.filterwarnings("ignore")

import pandas as pd
import numpy as np
import plotly.graph_objects as go
import streamlit as st

# ── Pagina ──────────────────────────────────────────────
st.set_page_config(page_title="Satisfaccion Clientes · QUIMPAC",
                   page_icon="📋", layout="wide",
                   initial_sidebar_state="expanded")

st.markdown("""
<style>
  .stApp{background:#F0F4F8}
  .hdr{background:linear-gradient(135deg,#1F4E79,#2E75B6);padding:16px 24px;
       border-radius:10px;margin-bottom:14px;color:#fff}
  .hdr h1{margin:0;font-size:22px;color:#fff}
  .hdr p{margin:4px 0 0;font-size:12px;color:#BDD7EE}
  .kpi{background:#fff;border-radius:8px;padding:12px 14px;
       box-shadow:0 2px 6px rgba(0,0,0,.08);text-align:center;min-height:96px}
  .kpi-ic{font-size:20px;margin-bottom:1px}
  .kpi-lb{font-size:10px;color:#595959;font-weight:700;text-transform:uppercase;
          letter-spacing:.4px;margin:0}
  .kpi-v{font-size:28px;font-weight:800;margin:2px 0;line-height:1.1}
  .kpi-s{font-size:10px;color:#8E8E8E;margin:0}
  .sec{font-size:12px;font-weight:700;color:#1F4E79;border-left:4px solid #2E75B6;
       padding-left:8px;margin:6px 0 4px}
  section[data-testid="stSidebar"]{background:#1F4E79}
  section[data-testid="stSidebar"] label,
  section[data-testid="stSidebar"] .stMarkdown p{color:#fff!important}
  section[data-testid="stSidebar"] h1,h2,h3{color:#fff!important}
  .modebar{display:none!important}
  .block-container{padding-top:1rem;padding-bottom:2rem}
  .stTabs [data-baseweb="tab"]{font-size:13px;font-weight:600}
</style>
""", unsafe_allow_html=True)

# ── Datos ────────────────────────────────────────────────
@st.cache_data
def load_data():
    base = os.path.dirname(os.path.abspath(__file__))
    df = pd.read_excel(os.path.join(base, "Consolidado_Encuestas_Satisfaccion.xlsx"), header=1)
    df = df.rename(columns={
        "Año":"anio","Tipo de Encuesta":"tipo","Empresa / Cliente":"empresa",
        "Cargo":"cargo","Nombre Encuestado":"nombre","Fecha Encuesta":"fecha",
        "Producto Evaluado":"producto",
        "Cantidad Entregada":"c_cantidad","Tiempo de Entrega":"c_tiempo_entrega",
        "Calidad":"c_calidad","Precio":"c_precio","Apoyo Técnico":"c_apoyo_tec",
        "Servicio al Cliente":"c_servicio","Tiempo Tránsito":"c_tiempo_transito",
        "Doc. Logística":"c_doc_log","Servicio Técnico":"c_servicio_tec",
        "Soporte Comercial":"c_soporte","Mercadeo":"c_mercadeo",
        "Devoluciones/Averías":"c_devoluciones","Asesoría Vendedor":"c_asesoria",
        "Dinámicas Comerciales":"c_dinamicas",
        "Comentario Criterios":"comentario_crit","NPS (0-10)":"nps",
        "NPS Comentario":"nps_comentario","Tiene Quejas":"tiene_quejas",
        "Razón Quejas":"razon_queja","Especificación Queja":"especif_queja",
        "Gestión Queja":"gestion_queja","Aspectos a Mejorar":"aspectos_mejorar",
        "Ampliación Mejora":"ampliacion_mejora",
        "Productos Otra Cía":"prods_otra","Razón Compra Otra":"razon_otra",
        "Comentario Prefer.":"comentario_pref",
    })
    df["nps"]  = pd.to_numeric(df["nps"], errors="coerce")
    df["anio"] = df["anio"].astype(int)
    return df

df_full = load_data()

# ── Constantes ───────────────────────────────────────────
TC = {"Químicos Nacional":"#2E75B6","Químicos Exportación":"#70AD47",
      "Sal Industrial":"#ED7D31","Sales Mineralizadas":"#9E3EA8"}
SC = {"Excelente":"#375623","Bueno":"#70AD47","Regular":"#ED7D31",
      "Malo":"#C00000","No sabe/ No aplica":"#BFBFBF"}
CRIT = {
    "c_cantidad":"Cantidad Entregada","c_tiempo_entrega":"Tiempo de Entrega",
    "c_calidad":"Calidad","c_precio":"Precio","c_apoyo_tec":"Apoyo Técnico",
    "c_servicio":"Servicio al Cliente","c_tiempo_transito":"Tiempo Tránsito",
    "c_doc_log":"Doc. Logística","c_servicio_tec":"Servicio Técnico",
    "c_soporte":"Soporte Comercial","c_mercadeo":"Mercadeo",
    "c_devoluciones":"Devoluciones/Averías","c_asesoria":"Asesoría Vendedor",
    "c_dinamicas":"Dinámicas Comerciales",
}

# ── Sentimientos ─────────────────────────────────────────
POS = {
    "excelente":3,"excelentes":3,"perfecto":3,"perfecta":3,
    "muy bien":2,"muy bueno":2,"muy buena":2,"bueno":2,"buena":2,
    "buenos":2,"buenas":2,"bien":1,"satisfecho":2,"satisfecha":2,
    "satisfechos":2,"puntual":2,"puntualidad":2,"oportuno":2,"oportuna":2,
    "eficiente":2,"eficiencia":2,"eficaz":2,
    "buena calidad":2,"recomiendo":2,"recomendable":2,
    "agradable":2,"amable":2,"cordial":2,"atento":2,"atenta":2,
    "rapido":2,"rapida":2,"pronto":1,"sin problemas":2,"sin novedad":2,
    "contento":2,"contenta":2,"conforme":2,
    "excelente servicio":3,"buen servicio":2,
}
NEG = {
    # Calidad baja / servicio deficiente
    "malo":3,"mala":3,"malos":3,"malas":3,
    "regular":2,"regulares":2,          # ← en encuestas = mediocre/insatisfactorio
    "deficiente":3,"deficientes":3,"deficiencia":2,
    # Tiempos
    "demora":2,"demoras":2,"demoran":2,"demorado":2,
    "retraso":2,"retrasos":2,"tarde":2,"tardanza":2,"atraso":2,
    "lento":2,"lenta":2,"lentitud":2,
    # Fallas
    "falla":2,"fallas":2,"fallo":2,"falló":2,
    "error":2,"errores":2,"inconveniente":2,"inconvenientes":2,
    # Incumplimiento
    "incumple":3,"incumplimiento":3,"no cumple":3,"no llega":2,
    # Precios
    "caro":2,"caros":2,"costoso":2,"costosa":2,"precio alto":2,"precios altos":2,
    # Atención
    "mal servicio":3,"mala atencion":3,"mala atención":3,
    "no responde":2,"no atiende":2,"no contesta":2,
    "complicado":1,"complicada":1,"dificil":1,"difícil":1,
    # Quejas generales
    "problema":2,"problemas":2,"queja":2,"quejas":2,
    "insatisfecho":3,"insatisfecha":3,
    "falta":2,"faltan":2,"faltó":2,
    "mejorar":1,"debe mejorar":2,"hay que mejorar":2,
}
STOP = {
    "de","la","el","en","y","a","los","del","se","las","un","por","con","no",
    "una","su","para","es","al","lo","como","mas","pero","sus","le","ya","o",
    "este","ha","si","porque","esta","son","entre","cuando","muy","sin","sobre",
    "ser","tiene","tambien","me","hasta","hay","donde","han","que","nos","desde",
    "todo","todos","uno","les","ni","contra","otros","ese","eso","ante","ellos",
    "e","esto","antes","algunos","unos","yo","otro","otras","tanto","esa",
    "estos","mucho","cual","poco","ella","estar","estas","algunas","algo",
    "nosotros","mi","mis","tu","fue","sido","cada","nuestro","nuestra",
    "solo","hacia","durante","despues","dicho","none","nan","n/a",
    "producto","productos","empresa","cliente","clientes","quimpac",
}

def _limpiar(text):
    """Limpia el texto: elimina 'None', 'nan', whitespace."""
    t = re.sub(r'\bnone\b|\bnan\b|\bn/a\b', ' ', str(text).lower())
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def sentimiento(text):
    if pd.isna(text): return None
    t = _limpiar(text)
    if len(t) < 6: return None          # texto demasiado corto → no clasificar
    p = sum(v for k,v in POS.items() if k in t)
    n = sum(v for k,v in NEG.items() if k in t)
    if p==0 and n==0: return "Neutro"
    return "Positivo" if p>n else ("Negativo" if n>p else "Neutro")

def word_freq(series, top=20):
    txt = " ".join(series.dropna().astype(str))
    txt = _limpiar(txt)
    txt = re.sub(r"[^a-z\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00f1\s]"," ",txt)
    words = [w for w in txt.split() if len(w)>3 and w not in STOP]
    return Counter(words).most_common(top)

def treemap_palabras(series, title, colorscale):
    """Genera un Plotly Treemap de frecuencia de palabras (estilo nube)."""
    wf = word_freq(series)
    if not wf:
        return None
    df_w = pd.DataFrame(wf, columns=["palabra","freq"])
    fig = go.Figure(go.Treemap(
        labels=df_w["palabra"],
        parents=[""] * len(df_w),
        values=df_w["freq"],
        textinfo="label+value",
        textfont=dict(size=14, family="Segoe UI, Arial"),
        marker=dict(
            colors=df_w["freq"],
            colorscale=colorscale,
            showscale=False,
        ),
        hovertemplate="<b>%{label}</b><br>Frecuencia: %{value}<extra></extra>",
    ))
    fig.update_layout(
        height=380, margin=dict(l=5,r=5,t=12,b=5),
        paper_bgcolor="white",
    )
    return fig

def nps_sc(s):
    cats = s.dropna().apply(lambda v: "P" if v>=9 else("N" if v>=7 else "D"))
    t = len(cats)
    if t==0: return 0
    return round((cats=="P").sum()/t*100-(cats=="D").sum()/t*100,1)

def fl(fig, h=300, m=None):
    fig.update_layout(height=h, margin=m or dict(l=5,r=15,t=12,b=30),
                      plot_bgcolor="white", paper_bgcolor="white",
                      font=dict(family="Segoe UI, Arial", size=11))
    return fig

# ── Sidebar ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("## Filtros")
    st.markdown("---")
    anios_d = sorted(df_full["anio"].unique())
    anios   = st.multiselect("Año", options=anios_d, default=anios_d)
    tipos_d = ["Todos"]+sorted(df_full["tipo"].unique())
    tipo    = st.selectbox("Tipo de encuesta", tipos_d)
    prev    = df_full[df_full["anio"].isin(anios or anios_d)]
    if tipo!="Todos": prev=prev[prev["tipo"]==tipo]
    emp_d   = ["Todas"]+sorted(prev["empresa"].dropna().unique())
    empresa = st.selectbox("Empresa / Cliente", emp_d)
    st.markdown("---")
    st.caption("QUIMPAC DE COLOMBIA S.A.\nEncuestas 2024–2025")

# ── Filtrado ─────────────────────────────────────────────
anios = anios or anios_d
dff   = df_full[df_full["anio"].isin(anios)]
if tipo!="Todos":    dff = dff[dff["tipo"]==tipo]
if empresa!="Todas": dff = dff[dff["empresa"]==empresa]

# ── Header ───────────────────────────────────────────────
st.markdown("""<div class="hdr"><h1>📋 Satisfaccion de Clientes</h1>
<p>QUIMPAC DE COLOMBIA S.A. &nbsp;·&nbsp; Tablero Gerencial y Comercial &nbsp;·&nbsp; 2024–2025</p>
</div>""", unsafe_allow_html=True)

if dff.empty:
    st.warning("Sin datos para los filtros seleccionados.")
    st.stop()

# ── KPIs ─────────────────────────────────────────────────
total    = len(dff)
np_prom  = round(dff["nps"].mean(),1)
np_score = nps_sc(dff["nps"])
cal      = dff["c_calidad"].dropna()
pct_exc  = round((cal=="Excelente").sum()/len(cal)*100,1) if len(cal) else 0
n_q      = (dff["tiene_quejas"]=="Si").sum()
pct_q    = round(n_q/total*100,1)
pct_p    = round((dff["nps"]>=9).sum()/total*100,1)

def kpi(ic,lb,val,sub,col):
    return (f'<div class="kpi"><div class="kpi-ic">{ic}</div>'
            f'<p class="kpi-lb">{lb}</p>'
            f'<p class="kpi-v" style="color:{col};">{val}</p>'
            f'<p class="kpi-s">{sub}</p></div>')

k1,k2,k3,k4,k5,k6 = st.columns(6)
k1.markdown(kpi("📋","Total Encuestas",f"{total:,}",
    f"{dff['tipo'].nunique()} tipos · {len(anios)} año(s)","#1F4E79"),unsafe_allow_html=True)
k2.markdown(kpi("⭐","NPS Promedio",f"{np_prom}","Escala 0–10",
    "#375623" if np_prom>=9 else "#ED7D31"),unsafe_allow_html=True)
k3.markdown(kpi("📈","NPS Score",f"{np_score:+.0f}%","Promotores − Detractores",
    "#375623" if np_score>=50 else "#ED7D31"),unsafe_allow_html=True)
k4.markdown(kpi("✅","Calidad Excelente",f"{pct_exc}%","Respuestas Calidad","#375623"),
    unsafe_allow_html=True)
k5.markdown(kpi("⚠️","Con Quejas",f"{pct_q}%",f"{n_q} clientes",
    "#C00000" if pct_q>20 else "#ED7D31"),unsafe_allow_html=True)
k6.markdown(kpi("👍","Promotores",f"{pct_p}%","NPS ≥ 9","#2E75B6"),unsafe_allow_html=True)
st.markdown("<br>",unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════
t1,t2,t3,t4,t5 = st.tabs([
    "📊 Resumen NPS","📉 Criterios","⚠️ Quejas y Mejoras",
    "💬 Análisis de Texto","🎯 Clientes Prioritarios"])

# ──────── TAB 1: RESUMEN NPS ─────────────────────────────
with t1:
    cA, cB = st.columns([5,7])
    with cA:
        st.markdown('<p class="sec">NPS Promedio por Tipo</p>',unsafe_allow_html=True)
        nt = dff.groupby("tipo")["nps"].mean().round(1).reset_index().sort_values("nps")
        f = go.Figure()
        for _,r in nt.iterrows():
            f.add_trace(go.Bar(y=[r["tipo"]],x=[r["nps"]],orientation="h",
                showlegend=False,marker_color=TC.get(r["tipo"],"#2E75B6"),
                text=[f"  {r['nps']}"],textposition="inside",
                textfont=dict(color="white",size=12,family="Arial Bold")))
        f.add_vline(x=9,line_dash="dash",line_color="#375623",line_width=2,
                    annotation_text="Meta 9",annotation_font_color="#375623",
                    annotation_position="top")
        fl(f,260,dict(l=5,r=20,t=20,b=20))
        f.update_layout(xaxis=dict(range=[0,10.5]),yaxis=dict(tickfont=dict(size=11)))
        st.plotly_chart(f,use_container_width=True,config={"displayModeBar":False})

    with cB:
        st.markdown('<p class="sec">Promotores · Pasivos · Detractores por Tipo</p>',
                    unsafe_allow_html=True)
        def seg(v):
            if pd.isna(v): return None
            return "Promotores (≥9)" if v>=9 else ("Pasivos (7-8)" if v>=7 else "Detractores (<7)")
        tmp = dff.copy(); tmp["seg"]=tmp["nps"].apply(seg)
        sg  = tmp.dropna(subset=["seg"]).groupby(["tipo","seg"]).size().unstack(fill_value=0).reset_index()
        seg_ord  = ["Promotores (≥9)","Pasivos (7-8)","Detractores (<7)"]
        seg_cols = {"Promotores (≥9)":"#375623","Pasivos (7-8)":"#FFC000","Detractores (<7)":"#C00000"}
        f2=go.Figure()
        for sn,sc2 in seg_cols.items():
            if sn not in sg.columns: continue
            avail=[c for c in seg_ord if c in sg.columns]
            tot=sg[avail].sum(axis=1).replace(0,1)
            pcts=(sg[sn]/tot*100).round(1)
            f2.add_trace(go.Bar(name=sn,x=sg["tipo"],y=pcts,marker_color=sc2,
                text=[f"{p:.0f}%" for p in pcts],textposition="inside",
                textfont=dict(color="white",size=11)))
        fl(f2,260,dict(l=5,r=10,t=20,b=65))
        f2.update_layout(barmode="stack",yaxis=dict(ticksuffix="%",range=[0,105]),
            xaxis=dict(tickangle=-10,tickfont=dict(size=10)),
            legend=dict(orientation="h",y=-0.28,font=dict(size=10)))
        st.plotly_chart(f2,use_container_width=True,config={"displayModeBar":False})

    st.markdown("<br>",unsafe_allow_html=True)
    cC,cD=st.columns([5,7])
    with cC:
        st.markdown('<p class="sec">Distribución de Calificaciones</p>',unsafe_allow_html=True)
        cc=[c for c in CRIT if c in dff.columns]
        scores=pd.concat([dff[c].dropna() for c in cc])
        sc2=scores.value_counts().reset_index(); sc2.columns=["s","n"]
        om={"Excelente":0,"Bueno":1,"Regular":2,"Malo":3,"No sabe/ No aplica":4}
        sc2["o"]=sc2["s"].map(om).fillna(9); sc2=sc2.sort_values("o")
        f3=go.Figure(go.Pie(labels=sc2["s"],values=sc2["n"],hole=0.55,
            marker_colors=[SC.get(x,"#999") for x in sc2["s"]],
            textinfo="percent",textfont=dict(size=11),
            hovertemplate="%{label}: %{value} (%{percent})<extra></extra>"))
        f3.update_layout(legend=dict(orientation="v",x=1.0,y=0.5,font=dict(size=10)))
        fl(f3,260,dict(l=5,r=5,t=12,b=10))
        st.plotly_chart(f3,use_container_width=True,config={"displayModeBar":False})

    with cD:
        st.markdown('<p class="sec">Evolución NPS por Año y Tipo</p>',unsafe_allow_html=True)
        ta=dff["tipo"].unique()
        te=df_full[df_full["tipo"].isin(ta)].groupby(["anio","tipo"])["nps"].mean().round(1).reset_index()
        f4=go.Figure()
        for tp in sorted(ta):
            s=te[te["tipo"]==tp].sort_values("anio")
            col2=TC.get(tp,"#595959")
            f4.add_trace(go.Scatter(x=s["anio"].astype(str),y=s["nps"],
                mode="lines+markers+text",name=tp,
                line=dict(color=col2,width=2.5),marker=dict(size=9,color=col2),
                text=[str(v) for v in s["nps"]],
                textposition="top center",textfont=dict(size=10)))
        f4.add_hline(y=9,line_dash="dash",line_color="#375623",line_width=2,
                     annotation_text="Meta 9",annotation_position="top right",
                     annotation_font_color="#375623")
        fl(f4,260,dict(l=5,r=5,t=20,b=30))
        f4.update_layout(xaxis=dict(title="Año"),yaxis=dict(title="NPS Prom.",range=[7,10.5]),
            legend=dict(font=dict(size=10)))
        st.plotly_chart(f4,use_container_width=True,config={"displayModeBar":False})

# ──────── TAB 2: CRITERIOS ───────────────────────────────
with t2:
    cE,cF=st.columns([7,5])
    with cE:
        st.markdown('<p class="sec">Calificación por Criterio — ordenado por % Excelente</p>',
                    unsafe_allow_html=True)
        rows=[]
        for ck,lb in CRIT.items():
            if ck not in dff.columns: continue
            sub=dff[ck].dropna()
            if len(sub)==0: continue
            row={"c":lb}
            for s in ["Excelente","Bueno","Regular","Malo"]:
                row[s]=round((sub==s).sum()/len(sub)*100,1)
            rows.append(row)
        dc=pd.DataFrame(rows).sort_values("Excelente",ascending=True)
        f5=go.Figure()
        for sn,col3 in [("Malo","#C00000"),("Regular","#ED7D31"),
                         ("Bueno","#70AD47"),("Excelente","#375623")]:
            if sn not in dc.columns: continue
            f5.add_trace(go.Bar(y=dc["c"],x=dc[sn],name=sn,orientation="h",
                marker_color=col3,
                text=[f"{v:.0f}%" if v>3 else "" for v in dc[sn]],
                textposition="inside",textfont=dict(color="white",size=9)))
        fl(f5,420,dict(l=5,r=10,t=12,b=50))
        f5.update_layout(barmode="stack",
            xaxis=dict(ticksuffix="%",range=[0,101]),yaxis=dict(tickfont=dict(size=10)),
            legend=dict(orientation="h",y=-0.12,font=dict(size=10)))
        st.plotly_chart(f5,use_container_width=True,config={"displayModeBar":False})

    with cF:
        st.markdown('<p class="sec">% Excelente por Criterio y Año</p>',unsafe_allow_html=True)
        acols={2024:"#2E75B6",2025:"#ED7D31"}
        f6=go.Figure()
        for yr in sorted(dff["anio"].unique()):
            dy=dff[dff["anio"]==yr]; ev,lb2=[],[]
            for ck,lb in CRIT.items():
                if ck not in dy.columns: continue
                sub=dy[ck].dropna()
                if len(sub)==0: continue
                ev.append(round((sub=="Excelente").sum()/len(sub)*100,1)); lb2.append(lb)
            f6.add_trace(go.Bar(name=str(yr),x=lb2,y=ev,
                marker_color=acols.get(yr,"#595959"),
                text=[f"{v:.0f}%" for v in ev],textposition="outside",textfont=dict(size=9)))
        fl(f6,420,dict(l=5,r=10,t=12,b=80))
        f6.update_layout(barmode="group",
            yaxis=dict(title="% Excelente",ticksuffix="%",range=[0,115]),
            xaxis=dict(tickangle=-35,tickfont=dict(size=9)),
            legend=dict(orientation="h",y=-0.22,font=dict(size=10)))
        st.plotly_chart(f6,use_container_width=True,config={"displayModeBar":False})

# ──────── TAB 3: QUEJAS ──────────────────────────────────
with t3:
    cG,cH,cI=st.columns(3)
    with cG:
        st.markdown('<p class="sec">Razones de Quejas (top 8)</p>',unsafe_allow_html=True)
        qdf=dff[dff["tiene_quejas"]=="Si"]["razon_queja"].dropna()
        if len(qdf)>0:
            qc=qdf.str.strip().value_counts().head(8).reset_index()
            qc.columns=["r","n"]; qc=qc.sort_values("n",ascending=True)
            f7=go.Figure(go.Bar(x=qc["n"],y=qc["r"],orientation="h",
                marker_color="#C00000",text=qc["n"],textposition="outside"))
            fl(f7,340,dict(l=5,r=35,t=12,b=30))
            f7.update_layout(xaxis=dict(title="N° quejas"))
            st.plotly_chart(f7,use_container_width=True,config={"displayModeBar":False})
        else:
            st.success("Sin quejas registradas.")

    with cH:
        st.markdown('<p class="sec">Aspectos a Mejorar (top 6)</p>',unsafe_allow_html=True)
        asp=dff["aspectos_mejorar"].dropna().str.strip().value_counts().head(6).reset_index()
        asp.columns=["a","n"]; asp["p"]=(asp["n"]/asp["n"].sum()*100).round(1)
        pal=["#ED7D31","#FFC000","#2E75B6","#70AD47","#9E3EA8","#595959"]
        f8=go.Figure(go.Bar(x=asp["a"],y=asp["n"],marker_color=pal[:len(asp)],
            text=[f"{p}%" for p in asp["p"]],textposition="outside"))
        fl(f8,340,dict(l=5,r=10,t=12,b=60))
        f8.update_layout(yaxis=dict(title="Menciones"),
            xaxis=dict(tickangle=-20,tickfont=dict(size=10)))
        st.plotly_chart(f8,use_container_width=True,config={"displayModeBar":False})

    with cI:
        st.markdown('<p class="sec">Tasa de Quejas por Tipo</p>',unsafe_allow_html=True)
        qt=dff.groupby("tipo")["tiene_quejas"].value_counts().unstack(fill_value=0).reset_index()
        if "Si" not in qt.columns: qt["Si"]=0
        if "No" not in qt.columns: qt["No"]=0
        qt["tot"]=qt["Si"]+qt["No"]
        qt["ps"]=(qt["Si"]/qt["tot"]*100).round(1)
        qt["pn"]=100-qt["ps"]
        f9=go.Figure()
        f9.add_trace(go.Bar(x=qt["tipo"],y=qt["ps"],name="Con queja",
            marker_color="#C00000",text=[f"{p:.0f}%" for p in qt["ps"]],
            textposition="inside",textfont=dict(color="white",size=10)))
        f9.add_trace(go.Bar(x=qt["tipo"],y=qt["pn"],name="Sin queja",
            marker_color="#70AD47",text=[f"{p:.0f}%" for p in qt["pn"]],
            textposition="inside",textfont=dict(color="white",size=10)))
        fl(f9,340,dict(l=5,r=5,t=12,b=65))
        f9.update_layout(barmode="stack",yaxis=dict(ticksuffix="%",range=[0,105]),
            xaxis=dict(tickangle=-20,tickfont=dict(size=8)),
            legend=dict(orientation="h",y=-0.22,font=dict(size=10)))
        st.plotly_chart(f9,use_container_width=True,config={"displayModeBar":False})

    # Competencia
    if "prods_otra" in dff.columns:
        po=dff["prods_otra"].dropna()
        po=po[~po.str.strip().str.lower().isin(["ninguno","ninguna","n/a","na","n.a","no"])]
        if len(po)>0:
            st.markdown("---")
            cJ,cK=st.columns(2)
            with cJ:
                st.markdown('<p class="sec">Productos que Compran a la Competencia</p>',
                            unsafe_allow_html=True)
                pc=po.str.strip().value_counts().head(10).reset_index()
                pc.columns=["p","n"]; pc["pct"]=(pc["n"]/len(dff)*100).round(1)
                pc=pc.sort_values("n",ascending=True)
                fc1=go.Figure(go.Bar(x=pc["n"],y=pc["p"],orientation="h",
                    marker_color="#9E3EA8",text=[f"{p}%" for p in pc["pct"]],
                    textposition="outside"))
                fl(fc1,300,dict(l=5,r=45,t=12,b=30))
                st.plotly_chart(fc1,use_container_width=True,config={"displayModeBar":False})
            if "razon_otra" in dff.columns:
                with cK:
                    st.markdown('<p class="sec">Razón para Comprar a Otra Empresa</p>',
                                unsafe_allow_html=True)
                    rz=dff["razon_otra"].dropna()
                    rz=rz[~rz.str.strip().str.lower().isin(["n.a","n.a.","na","n/a","ninguno","no"])]
                    if len(rz)>0:
                        rc=rz.str.strip().value_counts().head(8).reset_index()
                        rc.columns=["r","n"]; rc["pct"]=(rc["n"]/rc["n"].sum()*100).round(1)
                        rc=rc.sort_values("n",ascending=True)
                        fc2=go.Figure(go.Bar(x=rc["n"],y=rc["r"],orientation="h",
                            marker_color="#FFC000",text=[f"{p}%" for p in rc["pct"]],
                            textposition="outside"))
                        fl(fc2,300,dict(l=5,r=45,t=12,b=30))
                        st.plotly_chart(fc2,use_container_width=True,config={"displayModeBar":False})

# ──────── TAB 4: ANÁLISIS DE TEXTO ───────────────────────
with t4:
    st.markdown("""<p style="font-size:12px;color:#595959;margin-bottom:8px;">
    Análisis automático de sentimientos sobre los comentarios abiertos de las encuestas.
    Clasifica cada respuesta como <b>Positiva</b>, <b>Negativa</b> o <b>Neutra</b>
    mediante un diccionario de palabras clave en español — sin conexión a internet.</p>""",
    unsafe_allow_html=True)

    TCOLS={"nps_comentario":"Comentario NPS","comentario_crit":"Comentario Criterios",
           "especif_queja":"Especificación Queja","ampliacion_mejora":"Ampliación Mejora",
           "comentario_pref":"Comentario Preferencia"}
    tcp={k:v for k,v in TCOLS.items() if k in dff.columns}

    if not tcp:
        st.info("No hay columnas de texto disponibles.")
    else:
        dtx=dff.copy()
        # Combinar columnas de texto limpiando None/nan antes de unir
        dtx["txt_comb"]=(dtx[list(tcp.keys())]
                         .apply(lambda col: col.apply(
                             lambda v: "" if pd.isna(v) else _limpiar(str(v))))
                         .agg(" ".join, axis=1)
                         .str.strip())
        dtx["txt_comb"]=dtx["txt_comb"].replace("",pd.NA)
        dtx["sent"]=dtx["txt_comb"].apply(sentimiento)
        sv=dtx.dropna(subset=["sent"])
        tot_c=len(sv)

        cS1,cS2,cS3=st.columns([4,6,2])

        with cS1:
            st.markdown('<p class="sec">Distribución de Sentimientos</p>',unsafe_allow_html=True)
            if tot_c>0:
                sc3=sv["sent"].value_counts().reset_index(); sc3.columns=["s","n"]
                scm={"Positivo":"#375623","Neutro":"#BFBFBF","Negativo":"#C00000"}
                fs1=go.Figure(go.Pie(labels=sc3["s"],values=sc3["n"],hole=0.5,
                    marker_colors=[scm.get(x,"#999") for x in sc3["s"]],
                    textinfo="percent+label",textfont=dict(size=12),
                    hovertemplate="%{label}: %{value} (%{percent})<extra></extra>"))
                fl(fs1,260,dict(l=5,r=5,t=10,b=10)); fs1.update_layout(showlegend=False)
                st.plotly_chart(fs1,use_container_width=True,config={"displayModeBar":False})
                si=sc3.set_index("s")["n"]
                for s,col4 in [("Positivo","#375623"),("Neutro","#595959"),("Negativo","#C00000")]:
                    n=si.get(s,0); p=round(n/tot_c*100,1)
                    st.markdown(f"<span style='color:{col4};font-weight:700;'>{s}</span>: "
                                f"{n} ({p}%)<br>",unsafe_allow_html=True)
            else:
                st.info("Sin comentarios.")

        with cS2:
            st.markdown('<p class="sec">Sentimiento por Tipo de Encuesta</p>',unsafe_allow_html=True)
            if tot_c>0:
                stp=sv.groupby(["tipo","sent"]).size().unstack(fill_value=0).reset_index()
                fs2=go.Figure()
                for s,col5 in [("Positivo","#375623"),("Neutro","#BFBFBF"),("Negativo","#C00000")]:
                    if s not in stp.columns: continue
                    ec=[c for c in ["Positivo","Neutro","Negativo"] if c in stp.columns]
                    tot2=stp[ec].sum(axis=1).replace(0,1)
                    pc=(stp[s]/tot2*100).round(1)
                    fs2.add_trace(go.Bar(name=s,x=stp["tipo"],y=pc,marker_color=col5,
                        text=[f"{x:.0f}%" for x in pc],textposition="inside",
                        textfont=dict(color="white",size=10)))
                fl(fs2,260,dict(l=5,r=10,t=10,b=70))
                fs2.update_layout(barmode="stack",yaxis=dict(ticksuffix="%",range=[0,105]),
                    xaxis=dict(tickangle=-15,tickfont=dict(size=9)),
                    legend=dict(orientation="h",y=-0.32,font=dict(size=10)))
                st.plotly_chart(fs2,use_container_width=True,config={"displayModeBar":False})

        with cS3:
            st.markdown('<p class="sec">NPS por sentimiento</p>',unsafe_allow_html=True)
            if tot_c>0:
                ns=sv.groupby("sent")["nps"].mean().round(1)
                for s in ["Positivo","Neutro","Negativo"]:
                    if s not in ns.index: continue
                    col6={"Positivo":"#375623","Neutro":"#595959","Negativo":"#C00000"}[s]
                    st.markdown(
                        f"<div style='text-align:center;margin:10px 0;padding:8px;"
                        f"background:white;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.08);'>"
                        f"<span style='font-size:10px;font-weight:700;color:{col6};'>{s}</span>"
                        f"<br><span style='font-size:26px;font-weight:800;color:{col6};'>{ns[s]}</span>"
                        f"<br><span style='font-size:9px;color:#8E8E8E;'>NPS prom.</span></div>",
                        unsafe_allow_html=True)

        st.markdown("<br>",unsafe_allow_html=True)
        cW1,cW2=st.columns(2)

        with cW1:
            st.markdown('<p class="sec">🟩 Mapa de Palabras — Comentarios Positivos</p>',
                        unsafe_allow_html=True)
            pt=sv[sv["sent"]=="Positivo"]["txt_comb"]
            if len(pt)>0:
                fw1=treemap_palabras(pt,"Positivos","Greens")
                if fw1:
                    st.plotly_chart(fw1,use_container_width=True,config={"displayModeBar":False})
                else:
                    st.info("Pocas palabras para generar el mapa.")
            else:
                st.info("Sin comentarios positivos.")

        with cW2:
            st.markdown('<p class="sec">🟥 Mapa de Palabras — Comentarios Negativos</p>',
                        unsafe_allow_html=True)
            nt2=sv[sv["sent"]=="Negativo"]["txt_comb"]
            if len(nt2)>0:
                fw2=treemap_palabras(nt2,"Negativos","Reds")
                if fw2:
                    st.plotly_chart(fw2,use_container_width=True,config={"displayModeBar":False})
            else:
                st.info("Sin comentarios negativos.")

        st.markdown("---")
        st.markdown('<p class="sec">Comentarios con Clasificación</p>',unsafe_allow_html=True)
        sf=st.selectbox("Filtrar sentimiento",["Todos","Positivo","Neutro","Negativo"],key="sf")
        tb=sv[sv["txt_comb"].str.strip().str.len()>10].copy()
        if sf!="Todos": tb=tb[tb["sent"]==sf]

        cm=["tipo","anio","empresa","producto","nps","sent"]+[k for k in tcp if k in tb.columns]
        tbv=(tb[cm].rename(columns={"tipo":"Tipo","anio":"Año","empresa":"Empresa",
                "producto":"Producto","nps":"NPS","sent":"Sentimiento",**tcp})
             .sort_values(["Sentimiento","NPS"]).head(100).reset_index(drop=True))

        def crow(row):
            bg={"Positivo":"background-color:#C6EFCE","Negativo":"background-color:#FFCCCC"
                }.get(row.get("Sentimiento",""),"")
            return [bg]*len(row)
        st.dataframe(tbv.style.apply(crow,axis=1),use_container_width=True,height=400)
        st.caption(f"Mostrando hasta 100 de {len(tb)} comentarios.")

# ──────── TAB 5: CLIENTES PRIORITARIOS ───────────────────
with t5:
    st.markdown("""<p style="font-size:12px;color:#595959;margin-bottom:8px;">
    Registros con criterio <b>Regular/Malo</b> o NPS &lt; 8, ordenados por NPS.</p>""",
    unsafe_allow_html=True)
    ca=[c for c in CRIT if c in dff.columns]
    mb=dff[ca].isin(["Regular","Malo"]).any(axis=1)
    mn=dff["nps"]<8
    crit=dff[mb|mn].copy()

    if crit.empty:
        st.success("No hay registros críticos para los filtros seleccionados.")
    else:
        def gbad(row):
            b=[f"{lb}: {row[c]}" for c,lb in CRIT.items()
               if c in row.index and row[c] in ["Regular","Malo"]]
            return " | ".join(b) if b else "—"
        crit["Criterios Criticos"]=crit.apply(gbad,axis=1)

        m1,m2,m3,m4=st.columns(4)
        m1.metric("Registros críticos",len(crit))
        m2.metric("NPS promedio",round(crit["nps"].mean(),1))
        m3.metric("Empresas afectadas",crit["empresa"].nunique())
        m4.metric("Con quejas activas",(crit["tiene_quejas"]=="Si").sum())
        st.markdown("<br>",unsafe_allow_html=True)

        tb2=(crit[["tipo","anio","empresa","producto","nps","Criterios Criticos",
                    "comentario_crit","tiene_quejas","razon_queja","aspectos_mejorar"]]
             .rename(columns={"tipo":"Tipo","anio":"Año","empresa":"Empresa",
                 "producto":"Producto","nps":"NPS","comentario_crit":"Comentario",
                 "tiene_quejas":"Queja","razon_queja":"Razón Queja",
                 "aspectos_mejorar":"A Mejorar"})
             .sort_values(["NPS","Empresa"]).reset_index(drop=True))

        def cr2(row):
            bg="#FFCCCC" if row["NPS"]<8 else ("#FCE4D6" if row.get("Queja")=="Si" else "")
            return [f"background-color:{bg}" if bg else ""]*len(row)
        st.dataframe(tb2.style.apply(cr2,axis=1),use_container_width=True,height=400)

        cP1,cP2=st.columns(2)
        with cP1:
            st.markdown('<p class="sec">Distribución NPS — Registros Críticos</p>',
                        unsafe_allow_html=True)
            nd=crit["nps"].dropna().astype(int).value_counts().sort_index().reset_index()
            nd.columns=["v","n"]
            def nc(v): return "#C00000" if v<7 else("#FFC000" if v<9 else "#375623")
            fp1=go.Figure(go.Bar(x=nd["v"].astype(str),y=nd["n"],
                marker_color=[nc(v) for v in nd["v"]],
                text=nd["n"],textposition="outside"))
            fl(fp1,280,dict(l=5,r=10,t=12,b=30))
            fp1.update_layout(xaxis=dict(title="Valor NPS"),yaxis=dict(title="N° registros"))
            st.plotly_chart(fp1,use_container_width=True,config={"displayModeBar":False})
        with cP2:
            st.markdown('<p class="sec">Registros Críticos por Tipo</p>',unsafe_allow_html=True)
            ct=crit.groupby("tipo").size().reset_index(name="n")
            fp2=go.Figure(go.Bar(x=ct["tipo"],y=ct["n"],
                marker_color=[TC.get(t,"#595959") for t in ct["tipo"]],
                text=ct["n"],textposition="outside"))
            fl(fp2,280,dict(l=5,r=10,t=12,b=55))
            fp2.update_layout(yaxis=dict(title="N° registros"),
                xaxis=dict(tickangle=-15,tickfont=dict(size=9)))
            st.plotly_chart(fp2,use_container_width=True,config={"displayModeBar":False})

# ── Footer ───────────────────────────────────────────────
st.markdown("---")
st.caption(f"QUIMPAC DE COLOMBIA S.A. · {total} registros · "
           f"Años: {', '.join(map(str,sorted(anios)))} · "
           f"Generado desde Consolidado_Encuestas_Satisfaccion.xlsx")
