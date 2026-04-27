import base64
import io
from pathlib import Path
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="QSport Valuation Tool", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
#MainMenu {visibility:hidden;} footer {visibility:hidden;} header {visibility:hidden;}
[data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"] {display:none;}
.block-container {padding-top:1.4rem; padding-bottom:2rem;}
section[data-testid="stSidebar"] {background:#f7f8fb; border-right:1px solid #e6e8ef;}
.qs-card {background:white; border:1px solid #e8eaf0; border-radius:18px; padding:18px 20px; box-shadow:0 4px 18px rgba(20,30,55,.05); margin-bottom:14px;}
.qs-title {font-size:30px; font-weight:800; letter-spacing:-.03em; color:#111827; line-height:1.1;}
.qs-subtitle {color:#667085; font-size:14px; margin-top:6px;}
.qs-pill {display:inline-block; padding:5px 10px; border-radius:999px; font-size:12px; font-weight:700; border:1px solid #e6e8ef; background:#f9fafb; color:#344054; margin-right:6px;}
.metric-box {background:white; border:1px solid #e8eaf0; border-radius:18px; padding:18px; box-shadow:0 3px 14px rgba(20,30,55,.04);}
.metric-label {color:#667085; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:.04em;}
.metric-value {color:#111827; font-size:26px; font-weight:800; margin-top:4px;}
.metric-note {color:#667085; font-size:12px; margin-top:4px;}
.section-title {color:#111827; font-size:19px; font-weight:800; margin-top:10px; margin-bottom:6px;}
.small-muted {color:#667085; font-size:13px;}
</style>
""", unsafe_allow_html=True)
try:
    st.set_option("client.showErrorDetails", False)
except Exception:
    pass

ASSET_DIR = Path("assets")
LOGO_MELI = ASSET_DIR / "logo_meli.png"
LOGO_QSPORT = ASSET_DIR / "logo_qsport.png"

def img_b64(path: Path):
    return base64.b64encode(path.read_bytes()).decode("utf-8") if path.exists() else None

def show_logo(path: Path, width=180):
    if path.exists():
        st.image(str(path), width=width)

SCORE_LABELS = {"Muy Bajo":1, "Bajo":2, "Medio":3, "Alto":4, "Muy Alto":5}
SCORE_FROM_VALUE = {v:k for k,v in SCORE_LABELS.items()}
SCORE_COLS = ["Masividad", "Construcción de Marca", "Potencial de Negocio", "Valor Agregado", "Uso Esperado"]
INPUT_COLS = ["Asset", "Tipo de Asset", *SCORE_COLS]

ASSET_BASES = {
    "Naming Rights": 10.0,
    "Exclusividad": 7.0,
    "Hospitality": 5.0,
    "Contenido": 4.0,
    "PR / Comunicación": 3.0,
    "Signage / Visibilidad": 2.0,
    "Digital / Social": 3.5,
    "Experiencias": 4.5,
    "Data / CRM": 4.0,
    "Otros": 2.5,
}

DEFAULT_DATA = [
    ["Naming Rights Arena", "Naming Rights", "Muy Alto", "Muy Alto", "Alto", "Muy Alto", "Muy Alto"],
    ["Naming Rights Pitch / Field", "Naming Rights", "Alto", "Alto", "Medio", "Alto", "Medio"],
    ["Exclusividad de categoría", "Exclusividad", "Medio", "Muy Alto", "Muy Alto", "Alto", "Muy Alto"],
    ["Preventa MELI+ / Mercado Pago", "Data / CRM", "Medio", "Alto", "Muy Alto", "Alto", "Muy Alto"],
    ["Mercado Pago método oficial", "Exclusividad", "Medio", "Alto", "Muy Alto", "Alto", "Alto"],
    ["Hospitality premium", "Hospitality", "Medio", "Alto", "Alto", "Muy Alto", "Alto"],
    ["Experiencias VIP", "Experiencias", "Medio", "Alto", "Alto", "Muy Alto", "Alto"],
    ["Contenido Mercado Play", "Contenido", "Alto", "Alto", "Medio", "Alto", "Medio"],
    ["PR Mentions", "PR / Comunicación", "Alto", "Alto", "Medio", "Medio", "Alto"],
    ["Signage estadio", "Signage / Visibilidad", "Alto", "Medio", "Bajo", "Medio", "Alto"],
    ["Social Media Rights", "Digital / Social", "Alto", "Medio", "Medio", "Medio", "Medio"],
    ["Activaciones día de evento", "Experiencias", "Medio", "Medio", "Alto", "Alto", "Medio"],
]

def default_df():
    return pd.DataFrame(DEFAULT_DATA, columns=INPUT_COLS)

def score_to_label(v):
    if isinstance(v, str) and v in SCORE_LABELS:
        return v
    try:
        n = int(round(float(str(v).split("·")[0].strip())))
        n = max(1, min(5, n))
        return SCORE_FROM_VALUE[n]
    except Exception:
        return "Medio"

def normalize_df(df):
    df = df.copy()
    if "Marca" in df.columns and "Construcción de Marca" not in df.columns:
        df["Construcción de Marca"] = df["Marca"]
    if "Tipo de Asset" not in df.columns:
        df["Tipo de Asset"] = "Otros"
    for col in INPUT_COLS:
        if col not in df.columns:
            df[col] = "Medio" if col in SCORE_COLS else ("Otros" if col == "Tipo de Asset" else "")
    df = df[INPUT_COLS]
    df["Asset"] = df["Asset"].fillna("").astype(str)
    df = df[df["Asset"].str.strip() != ""]
    if df.empty:
        df = default_df()
    for col in SCORE_COLS:
        df[col] = df[col].apply(score_to_label)
    df["Tipo de Asset"] = df["Tipo de Asset"].apply(lambda x: x if x in ASSET_BASES else "Otros")
    return df.reset_index(drop=True)

def classify(score):
    if score >= 4.5:
        return "Core Premium"
    if score >= 3.5:
        return "Strategic Value"
    if score >= 2.5:
        return "Supporting Asset"
    return "Low Priority"

def calculate(df, unit_value):
    df = normalize_df(df).copy()
    for col in SCORE_COLS:
        df[col + " Num"] = df[col].map(SCORE_LABELS).astype(float)
    df["Base de Valor"] = df["Tipo de Asset"].map(ASSET_BASES).fillna(ASSET_BASES["Otros"])
    df["Score Final"] = df[[c + " Num" for c in SCORE_COLS]].mean(axis=1).round(2)
    df["Clasificación"] = df["Score Final"].apply(classify)
    df["Valor Estimado"] = (df["Base de Valor"] * (df["Score Final"] / 5) * unit_value).round(2)
    total = df["Valor Estimado"].sum()
    df["Participación %"] = (df["Valor Estimado"] / total * 100).fillna(0).round(1)
    return df

def money(v, currency):
    return f"{currency} {v:,.1f}M"

def recommendation(estimated, price, currency):
    if price <= 0:
        return "Ingresá el precio solicitado para comparar la propuesta contra el valor estimado del paquete."
    gap = estimated - price
    pct = gap / price * 100
    if pct >= 10:
        return f"El paquete presenta una oportunidad: el valor estimado supera el precio solicitado en {money(abs(gap), currency)} ({abs(pct):.1f}%). La recomendación es avanzar, asegurando los assets Core Premium."
    if pct <= -10:
        return f"El paquete muestra sobreprecio: el precio solicitado supera el valor estimado en {money(abs(gap), currency)} ({abs(pct):.1f}%). La recomendación es negociar reducción, sumar derechos o mejorar activos clave."
    return "El precio solicitado está en una zona razonable frente al valor estimado. La recomendación es avanzar con foco en garantías de ejecución y uso efectivo de los assets."

def build_excel(df, summary, currency):
    wb = Workbook()
    ws = wb.active
    ws.title = "Valuation V8"
    ws2 = wb.create_sheet("Methodology")
    blue="3483FA"; grey="F3F4F6"; white="FFFFFF"
    thin=Side(style="thin", color="D0D5DD")
    border=Border(left=thin,right=thin,top=thin,bottom=thin)
    ws.merge_cells("A1:L1")
    ws["A1"] = "QSport Sponsorship Valuation Tool — V8"
    ws["A1"].font = Font(bold=True, size=16, color="111827")
    meta = [["Cliente", summary["client"]], ["Proyecto", summary["project"]], ["Precio solicitado", summary["asking_price"]], ["Valor estimado", summary["estimated"]], ["Gap", summary["gap"]], ["Gap %", summary["gap_pct"]], ["Fecha", datetime.now().strftime("%Y-%m-%d %H:%M")]]
    for r,row in enumerate(meta,3):
        ws.cell(r,1,row[0]).font = Font(bold=True)
        ws.cell(r,1).fill = PatternFill("solid", fgColor=grey)
        ws.cell(r,2,row[1])
    cols = ["Asset","Tipo de Asset","Base de Valor",*SCORE_COLS,"Score Final","Clasificación","Valor Estimado","Participación %"]
    start=12
    for j,col in enumerate(cols,1):
        cell=ws.cell(start,j,col)
        cell.font=Font(bold=True,color=white)
        cell.fill=PatternFill("solid", fgColor=blue)
        cell.border=border
        cell.alignment=Alignment(horizontal="center", wrap_text=True)
    for i,(_,row) in enumerate(df[cols].iterrows(),start+1):
        for j,col in enumerate(cols,1):
            cell=ws.cell(i,j,row[col])
            cell.border=border
            cell.alignment=Alignment(wrap_text=True)
    for i in range(1,len(cols)+1):
        ws.column_dimensions[get_column_letter(i)].width=22
    ws2["A1"]="Metodología V8"
    ws2["A1"].font=Font(bold=True,size=16)
    notes=[
        ["Principio", "Primero estima el valor independiente del paquete y luego lo compara contra el precio solicitado."],
        ["Fórmula", "Valor Asset = Base de Valor × (Score Final / 5) × Valor por unidad"],
        ["Base de Valor", "Referencia QSport por tipo de asset. No depende del precio pedido."],
        ["Score Final", "Promedio de Masividad, Construcción de Marca, Potencial de Negocio, Valor Agregado y Uso Esperado."],
        ["Clasificación", "Core Premium >=4.5; Strategic Value >=3.5; Supporting Asset >=2.5; Low Priority <2.5"],
    ]
    for r,row in enumerate(notes,3):
        ws2.cell(r,1,row[0]).font=Font(bold=True)
        ws2.cell(r,2,row[1])
    ws2.column_dimensions["A"].width=24
    ws2.column_dimensions["B"].width=110
    bio=io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.read()

with st.sidebar:
    show_logo(LOGO_QSPORT, 190)
    st.markdown("### Sponsorship Valuation")
    st.caption("Modelo V8 — valor estimado independiente vs precio solicitado")
    st.divider()
    client_name = st.text_input("Cliente", "Mercado Livre")
    project_name = st.text_input("Proyecto", "Arena Pacaembu")
    currency = st.selectbox("Moneda", ["USD","BRL","EUR"], index=0)
    asking_price = st.number_input("Precio solicitado por la propiedad (M)", min_value=0.0, value=28.0, step=0.5)
    unit_value = st.number_input("Valor por unidad QSport (M)", min_value=0.1, value=1.0, step=0.1, help="Convierte la escala de bases en dinero. Ejemplo: 1 unidad = USD 1M.")
    st.divider()
    uploaded = st.file_uploader("Cargar Excel de assets", type=["xlsx","xls","csv"])
    st.divider()
    with st.expander("Referencia metodológica", expanded=False):
        st.markdown("""
        La herramienta **no reparte el precio del contrato**.

        Primero estima el valor independiente de cada asset:

        `Valor Asset = Base × (Score / 5) × Valor por unidad`

        Luego compara el total contra el precio solicitado.
        """)
        st.dataframe(pd.DataFrame({"Score Final":["≥ 4.5","≥ 3.5","≥ 2.5","< 2.5"],"Clasificación":["Core Premium","Strategic Value","Supporting Asset","Low Priority"]}), hide_index=True, use_container_width=True)
        st.dataframe(pd.DataFrame({"Tipo de Asset":list(ASSET_BASES.keys()),"Base":list(ASSET_BASES.values())}), hide_index=True, use_container_width=True)

if "working_df" not in st.session_state:
    st.session_state.working_df = default_df()
if uploaded is not None:
    try:
        up = pd.read_csv(uploaded) if uploaded.name.lower().endswith(".csv") else pd.read_excel(uploaded)
        st.session_state.working_df = normalize_df(up)
        st.sidebar.success("Archivo cargado correctamente.")
    except Exception:
        st.sidebar.error("No se pudo leer el archivo. Revisá estructura o formato.")

meli = img_b64(LOGO_MELI)
if meli:
    st.markdown(f"""
    <div class="qs-card"><div style="display:flex; align-items:center; gap:18px;">
      <div style="width:170px; min-width:170px;"><img src="data:image/png;base64,{meli}" style="max-width:160px; max-height:58px; object-fit:contain;"></div>
      <div><div class="qs-title">Sponsorship / Naming Rights Valuation Tool</div>
      <div class="qs-subtitle">Valoración estratégica independiente del paquete de assets vs precio solicitado</div>
      <div style="margin-top:10px;"><span class="qs-pill">V8 Valuation Model</span><span class="qs-pill">QSport Methodology</span><span class="qs-pill">{client_name}</span></div></div>
    </div></div>""", unsafe_allow_html=True)
else:
    st.markdown('<div class="qs-card"><div class="qs-title">Sponsorship / Naming Rights Valuation Tool</div><div class="qs-subtitle">Valoración estratégica independiente del paquete de assets vs precio solicitado</div></div>', unsafe_allow_html=True)

st.markdown('<div class="section-title">1. Carga y scoring de assets</div>', unsafe_allow_html=True)
st.markdown('<div class="small-muted">Asigná tipo de asset y puntajes. La base de valor viene de la metodología QSport, no del precio pedido.</div>', unsafe_allow_html=True)

edited = st.data_editor(
    st.session_state.working_df,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        "Asset": st.column_config.TextColumn("Asset", width="large"),
        "Tipo de Asset": st.column_config.SelectboxColumn("Tipo de Asset", options=list(ASSET_BASES.keys()), required=True),
        **{col: st.column_config.SelectboxColumn(col, options=list(SCORE_LABELS.keys()), required=True) for col in SCORE_COLS},
    },
    key="asset_editor_v8",
)
st.session_state.working_df = normalize_df(edited)
result = calculate(st.session_state.working_df, unit_value)
estimated = float(result["Valor Estimado"].sum())
gap = estimated - asking_price
gap_pct = gap / asking_price * 100 if asking_price else 0
core_count = int((result["Clasificación"] == "Core Premium").sum())

st.markdown('<div class="section-title">2. Resultado ejecutivo</div>', unsafe_allow_html=True)
c1,c2,c3,c4 = st.columns(4)
with c1:
    st.markdown(f'<div class="metric-box"><div class="metric-label">Valor estimado del paquete</div><div class="metric-value">{money(estimated,currency)}</div><div class="metric-note">Suma independiente de assets</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="metric-box"><div class="metric-label">Precio solicitado</div><div class="metric-value">{money(asking_price,currency)}</div><div class="metric-note">Oferta presentada por la propiedad</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="metric-box"><div class="metric-label">Gap vs precio</div><div class="metric-value">{money(gap,currency)}</div><div class="metric-note">{gap_pct:.1f}% vs precio pedido</div></div>', unsafe_allow_html=True)
with c4:
    st.markdown(f'<div class="metric-box"><div class="metric-label">Assets / Core Premium</div><div class="metric-value">{len(result)} / {core_count}</div><div class="metric-note">Cantidad total y activos premium</div></div>', unsafe_allow_html=True)

st.markdown(f'<div class="qs-card"><div style="font-weight:800; color:#111827; margin-bottom:6px;">Lectura estratégica</div><div style="color:#475467; font-size:15px;">{recommendation(estimated, asking_price, currency)}</div></div>', unsafe_allow_html=True)

st.markdown('<div class="section-title">3. Tabla de valoración recalculada</div>', unsafe_allow_html=True)
display_cols = ["Asset","Tipo de Asset","Base de Valor",*SCORE_COLS,"Score Final","Clasificación","Valor Estimado","Participación %"]
display = result[display_cols].sort_values("Valor Estimado", ascending=False).reset_index(drop=True)
st.dataframe(display, use_container_width=True, hide_index=True)

ch1,ch2 = st.columns(2)
with ch1:
    st.markdown('<div class="section-title">Valor estimado por asset</div>', unsafe_allow_html=True)
    fig = px.bar(display.head(15).sort_values("Valor Estimado"), x="Valor Estimado", y="Asset", orientation="h", text="Valor Estimado")
    fig.update_layout(height=480, margin=dict(l=10,r=10,t=20,b=10), xaxis_title=f"Valor estimado ({currency} M)", yaxis_title="", plot_bgcolor="white", paper_bgcolor="white")
    st.plotly_chart(fig, use_container_width=True)
with ch2:
    st.markdown('<div class="section-title">Valor por clasificación</div>', unsafe_allow_html=True)
    class_df = result.groupby("Clasificación", as_index=False)["Valor Estimado"].sum().sort_values("Valor Estimado", ascending=False)
    fig2 = px.pie(class_df, values="Valor Estimado", names="Clasificación", hole=.55)
    fig2.update_layout(height=480, margin=dict(l=10,r=10,t=20,b=10), paper_bgcolor="white")
    st.plotly_chart(fig2, use_container_width=True)

st.markdown('<div class="section-title">4. Escenarios de negociación</div>', unsafe_allow_html=True)
sc = pd.DataFrame({"Escenario":["Oferta actual","Renovación optimizada","Escenario agresivo"], "Precio propuesto":[asking_price, asking_price*.90, asking_price*.80]})
sc["Gap vs valor"] = estimated - sc["Precio propuesto"]
sc["Gap %"] = (sc["Gap vs valor"] / sc["Precio propuesto"] * 100).replace([float("inf"), -float("inf")], 0).fillna(0)
sc["Lectura"] = sc["Gap %"].apply(lambda x: "Oportunidad" if x >= 10 else ("Razonable" if x > -10 else "Sobreprecio"))
st.dataframe(sc.round({"Precio propuesto":2,"Gap vs valor":2,"Gap %":1}), use_container_width=True, hide_index=True)

st.markdown('<div class="section-title">5. Exportables</div>', unsafe_allow_html=True)
summary = {"client":client_name,"project":project_name,"asking_price":asking_price,"estimated":estimated,"gap":gap,"gap_pct":gap_pct}
excel = build_excel(result, summary, currency)
st.download_button("Descargar Excel actualizado", data=excel, file_name=f"{client_name}_{project_name}_valuation_v8.xlsx".replace(" ","_"), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
st.info("V8 prioriza la lógica de valuación independiente. El PDF ejecutivo premium puede sumarse en la próxima iteración.")

st.markdown('<div style="margin-top:24px; padding-top:14px; border-top:1px solid #e6e8ef; color:#98a2b3; font-size:12px;">QSport Sponsorship Valuation Tool — V8. Modelo estratégico de referencia. No reemplaza una medición de media value con exposición real.</div>', unsafe_allow_html=True)
