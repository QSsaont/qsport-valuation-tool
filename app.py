import base64
import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

SCORE_COLUMNS = [
    "Masividad",
    "Construcción de Marca",
    "Potencial de Negocio",
    "Valor Agregado",
    "Uso Esperado",
]

OUTPUT_COLUMNS = [
    "Asset",
    *SCORE_COLUMNS,
    "Score Final",
    "Clasificación",
    "Valor Económico",
]

SCORE_LABELS = {
    1: "1 · Muy Bajo",
    2: "2 · Bajo",
    3: "3 · Medio",
    4: "4 · Alto",
    5: "5 · Muy Alto",
}
LABEL_TO_SCORE = {v: k for k, v in SCORE_LABELS.items()}
CLASSIFICATION_ORDER = ["Core Premium", "Strategic Value", "Supporting Asset", "Low Priority"]
CLASSIFICATION_COLORS = {
    "Core Premium": "#00A650",
    "Strategic Value": "#3483FA",
    "Supporting Asset": "#F5B700",
    "Low Priority": "#9CA3AF",
}
SCENARIO_MULTIPLIERS = {
    "Contrato actual": 1.00,
    "Renovación optimizada": 1.15,
    "Escenario agresivo": 1.30,
}

DEFAULT_ASSETS = [
    ["Naming Rights Arena", 5, 5, 5, 5, 5],
    ["Naming Rights Pitch / Field", 4, 3, 5, 4, 3],
    ["Naming Rights Events Center", 4, 4, 4, 4, 4],
    ["PR Mentions", 5, 5, 5, 5, 5],
    ["Branding complejo / signage", 1, 1, 1, 1, 1],
    ["Bicicletero branded", 1, 1, 1, 1, 1],
    ["Exclusividad categoría", 5, 5, 5, 5, 5],
    ["Mercado Pago método oficial", 1, 1, 1, 1, 1],
    ["Mercado Play Content Studio", 5, 5, 5, 5, 5],
    ["Preventa MELI+ / MPago", 5, 5, 5, 5, 5],
    ["Hospitality / palcos premium", 3, 4, 5, 5, 4],
    ["Experiencias para sellers y clientes", 3, 4, 5, 5, 4],
    ["Activaciones en días de evento", 4, 4, 5, 4, 4],
    ["Contenido digital y redes sociales", 4, 4, 4, 3, 4],
    ["Presencia LED / pantallas internas", 4, 3, 4, 3, 4],
    ["Entradas y beneficios comerciales", 3, 3, 4, 4, 4],
    ["Uso de imagen del venue", 3, 4, 3, 4, 4],
    ["Acciones B2B y networking", 2, 3, 5, 4, 3],
    ["Integración app / wallet", 4, 4, 5, 4, 3],
    ["Promociones co-brandeadas", 3, 3, 5, 3, 4],
    ["Base de datos / leads opt-in", 2, 3, 5, 5, 3],
]

VARIABLE_DEFINITIONS = {
    "Masividad": "Capacidad del asset para generar alcance amplio, visibilidad sostenida y exposición frecuente.",
    "Construcción de Marca": "Capacidad de fortalecer percepción, prestigio, liderazgo y posicionamiento.",
    "Potencial de Negocio": "Capacidad de generar ventas, usuarios, transacciones, promociones o leads.",
    "Valor Agregado": "Beneficios diferenciales como exclusividad, experiencias premium, contenido único o acceso preferencial.",
    "Uso Esperado": "Probabilidad real de uso efectivo por parte de la marca.",
}


def clean_score(value) -> float:
    if isinstance(value, str):
        value = LABEL_TO_SCORE.get(value, value.split("·", 1)[0].strip() if "·" in value else value)
    try:
        value = float(value)
    except Exception:
        return 1.0
    return min(max(value, 1.0), 5.0)


def score_to_label(value) -> str:
    return SCORE_LABELS.get(int(round(clean_score(value))), "1 · Muy Bajo")


def classify_score(score: float) -> str:
    if pd.isna(score):
        return "Low Priority"
    if score >= 4.5:
        return "Core Premium"
    if score >= 3.5:
        return "Strategic Value"
    if score >= 2.5:
        return "Supporting Asset"
    return "Low Priority"


def normalize_input(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "Marca" in df.columns and "Construcción de Marca" not in df.columns:
        df["Construcción de Marca"] = df["Marca"]
    for col in ["Asset", *SCORE_COLUMNS]:
        if col not in df.columns:
            df[col] = "" if col == "Asset" else 1
    df = df[["Asset", *SCORE_COLUMNS]].copy()
    df["Asset"] = df["Asset"].fillna("").astype(str)
    for col in SCORE_COLUMNS:
        df[col] = df[col].apply(clean_score)
    df = df[df["Asset"].str.strip() != ""].reset_index(drop=True)
    if df.empty:
        return pd.DataFrame(DEFAULT_ASSETS, columns=["Asset", *SCORE_COLUMNS])
    return df


def recalculate(df: pd.DataFrame, contract_value: float) -> pd.DataFrame:
    df = normalize_input(df)
    for col in SCORE_COLUMNS:
        df[col] = df[col].apply(clean_score)
    df["Score Final"] = df[SCORE_COLUMNS].mean(axis=1).round(2)
    total_score = df["Score Final"].sum()
    df["Valor Económico"] = 0.0 if total_score <= 0 else (df["Score Final"] / total_score * contract_value).round(3)
    df["Clasificación"] = df["Score Final"].apply(classify_score)
    return df[OUTPUT_COLUMNS]


def editor_dataframe(df: pd.DataFrame, expert_mode: bool) -> pd.DataFrame:
    df = normalize_input(df)
    if expert_mode:
        return df
    display = df.copy()
    for col in SCORE_COLUMNS:
        display[col] = display[col].apply(score_to_label)
    return display


def make_output_table(df: pd.DataFrame, currency: str) -> pd.DataFrame:
    out = df.copy()
    for col in SCORE_COLUMNS:
        out[col] = out[col].apply(score_to_label)
    out["Score Final"] = out["Score Final"].map(lambda x: f"{x:.1f}")
    out["Valor Económico"] = out["Valor Económico"].map(lambda x: f"{currency} {x:,.2f} M")
    return out


def make_excel_download(df: pd.DataFrame, client_name: str, project_name: str, currency: str, contract_value: float) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Modelo Scoring"
    ws2 = wb.create_sheet("Metodología")
    ws3 = wb.create_sheet("Escenarios")

    dark = "111827"
    yellow = "FFE600"
    blue = "3483FA"
    grey = "F9FAFB"
    white = "FFFFFF"
    border = Border(left=Side(style="thin", color="E5E7EB"), right=Side(style="thin", color="E5E7EB"), top=Side(style="thin", color="E5E7EB"), bottom=Side(style="thin", color="E5E7EB"))

    ws.merge_cells("A1:I1")
    ws["A1"] = "Sponsorship / Naming Rights Valuation Model"
    ws["A1"].font = Font(bold=True, size=16, color=dark)
    ws["A1"].fill = PatternFill("solid", fgColor=yellow)
    ws["A1"].alignment = Alignment(horizontal="center")

    meta = [["Cliente", client_name], ["Proyecto", project_name], ["Moneda", currency], ["Valor total contrato", contract_value], ["Fecha exportación", datetime.now().strftime("%Y-%m-%d %H:%M")]]
    for r, row in enumerate(meta, start=3):
        ws.cell(r, 1).value = row[0]
        ws.cell(r, 2).value = row[1]
        ws.cell(r, 1).font = Font(bold=True)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=grey)

    start_row = 10
    for c, h in enumerate(OUTPUT_COLUMNS, start=1):
        cell = ws.cell(start_row, c)
        cell.value = f"Valor Económico {currency} M" if h == "Valor Económico" else h
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=blue)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    export_df = df.copy()
    for col in SCORE_COLUMNS:
        export_df[col] = export_df[col].apply(score_to_label)

    for r_idx, row in enumerate(export_df.itertuples(index=False), start=start_row + 1):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(r_idx, c_idx)
            cell.value = value
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if c_idx == 7:
                cell.number_format = "0.0"
            if c_idx == 9:
                cell.number_format = "#,##0.00"

    widths = [42, 18, 22, 22, 18, 18, 14, 22, 20]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A11"

    ws2.append(["Metodología de scoring"])
    ws2["A1"].font = Font(bold=True, size=15)
    ws2.append([])
    ws2.append(["Escala", "Descripción"])
    for cell in ws2[3]:
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=blue)
    for k, v in SCORE_LABELS.items():
        ws2.append([k, v.split("·", 1)[1].strip()])
    ws2.append([])
    ws2.append(["Score Final", "Clasificación"])
    ws2.append([">= 4.5", "Core Premium"])
    ws2.append([">= 3.5", "Strategic Value"])
    ws2.append([">= 2.5", "Supporting Asset"])
    ws2.append(["< 2.5", "Low Priority"])
    ws2.append([])
    ws2.append(["Variable", "Definición"])
    for var, desc in VARIABLE_DEFINITIONS.items():
        ws2.append([var, desc])
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 90
    for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws3.append(["Escenario", "Multiplicador", f"Valor Contrato {currency} M"])
    for cell in ws3[1]:
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=blue)
    for scenario, multiplier in SCENARIO_MULTIPLIERS.items():
        ws3.append([scenario, multiplier, contract_value * multiplier])
    ws3.column_dimensions["A"].width = 28
    ws3.column_dimensions["B"].width = 18
    ws3.column_dimensions["C"].width = 26

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def make_pdf_download(df: pd.DataFrame, client_name: str, project_name: str, currency: str, contract_value: float) -> bytes | None:
    if not REPORTLAB_AVAILABLE:
        return None
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=32, leftMargin=32, topMargin=32, bottomMargin=32)
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Sponsorship / Naming Rights Valuation", styles["Title"]),
        Paragraph(f"Cliente: {client_name} | Proyecto: {project_name}", styles["Normal"]),
        Paragraph(f"Valor total contrato: {currency} {contract_value:,.2f} M", styles["Normal"]),
        Spacer(1, 14),
    ]
    summary = df.groupby("Clasificación", as_index=False)["Valor Económico"].sum().sort_values("Valor Económico", ascending=False)
    summary_data = [["Clasificación", f"Valor {currency} M"]] + [[r["Clasificación"], f"{r['Valor Económico']:,.2f}"] for _, r in summary.iterrows()]
    t = Table(summary_data, colWidths=[260, 140])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#111827")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D1D5DB")), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("ALIGN", (1, 1), (-1, -1), "RIGHT")]))
    story.append(t)
    story.append(Spacer(1, 14))

    top = df.sort_values("Valor Económico", ascending=False).head(10)
    table_data = [["Asset", "Score", "Clasificación", f"Valor {currency} M"]]
    for _, r in top.iterrows():
        table_data.append([r["Asset"], f"{r['Score Final']:.1f}", r["Clasificación"], f"{r['Valor Económico']:,.2f}"])
    t2 = Table(table_data, colWidths=[220, 55, 125, 95])
    t2.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#3483FA")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D1D5DB")), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (1, 1), (-1, -1), "RIGHT")]))
    story.append(Paragraph("Top assets por valor económico", styles["Heading2"]))
    story.append(t2)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def image_as_base64(path: str) -> str | None:
    logo_path = Path(path)
    if not logo_path.exists():
        return None
    return base64.b64encode(logo_path.read_bytes()).decode()


def render_agency_logo():
    encoded = image_as_base64("assets/logo_qsport.png")
    if encoded:
        st.markdown(
            f'<img src="data:image/png;base64,{encoded}" class="agency-logo" />',
            unsafe_allow_html=True,
        )
    else:
        st.markdown("**QSport · ZETABÉ**")


def render_main_header():
    encoded = image_as_base64("assets/logo_meli.png")
    if encoded:
        logo_html = f'<img src="data:image/png;base64,{encoded}" class="header-client-logo" />'
    else:
        logo_html = '<div class="client-logo-fallback">Mercado Livre</div>'
    st.markdown(
        f'''
        <section class="hero-header">
            <div class="hero-logo-wrap">{logo_html}</div>
            <div class="hero-divider"></div>
            <div class="hero-copy">
                <h1>Sponsorship / Naming Rights Valuation Tool</h1>
                <p>Modelo de scoring estratégico para valorar assets comerciales y redistribuir el valor económico del contrato.</p>
            </div>
        </section>
        ''',
        unsafe_allow_html=True,
    )

st.set_page_config(page_title="Mercado Livre Sponsorship Valuation", page_icon="📊", layout="wide")
st.markdown("""
<style>
.block-container {padding-top: 2.2rem; padding-bottom: 2rem; max-width: 1500px;}
h1, h2, h3 {letter-spacing: -0.035em;}
div[data-testid="stMetric"] {background:#FFFFFF; border:1px solid #E5E7EB; padding:18px 18px; border-radius:18px; box-shadow:0 1px 2px rgba(17,24,39,.04);}
div[data-testid="stMetricValue"] {font-size:1.85rem;}
.agency-logo {width:210px; max-width:100%; height:auto; display:block; margin:18px 0 26px 0;}
.hero-header {display:flex; align-items:center; gap:26px; padding:14px 0 28px 0; margin-bottom:28px; border-bottom:1px solid #D9DEE8;}
.hero-logo-wrap {display:flex; align-items:center; justify-content:center; min-width:220px; height:auto;}
.header-client-logo {width:215px; max-width:100%; height:auto; display:block; object-fit:contain;}
.hero-divider {width:2px; align-self:stretch; min-height:74px; background:#0B1F3A; opacity:.85;}
.hero-copy h1 {font-size:44px; line-height:1.02; margin:0; color:#111827; font-weight:800;}
.hero-copy p {font-size:16px; color:#6B7280; margin:14px 0 0 0;}
.client-logo-fallback {display:inline-block; background:#FFE600; color:#2D3277; font-weight:800; font-size:22px; padding:12px 20px; border-radius:999px; box-shadow:0 1px 2px rgba(17,24,39,.12);}
.method-box {background:#F9FAFB; border:1px solid #E5E7EB; border-radius:14px; padding:12px 14px; font-size:14px;}
@media (max-width: 900px) {
  .hero-header {align-items:flex-start; flex-direction:column; gap:14px;}
  .hero-divider {display:none;}
  .hero-copy h1 {font-size:34px;}
  .hero-logo-wrap {min-width:0;}
  .header-client-logo {width:190px;}
}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    render_agency_logo()
    st.header("Inputs globales")
    client_name = st.text_input("Nombre cliente", "Mercado Livre")
    project_name = st.text_input("Nombre proyecto", "Mercado Livre Arena Pacaembu")
    currency = st.selectbox("Moneda", ["USD", "BRL", "EUR"], index=0)
    contract_value = st.number_input("Valor total contrato (en millones)", min_value=0.0, value=10.0, step=0.5)
    uploaded_file = st.file_uploader("Cargar Excel de assets", type=["xlsx"])
    expert_mode = st.toggle("Modo experto: editar números manualmente", value=False)

    st.divider()
    st.subheader("Referencia metodológica")
    st.markdown("""
<div class="method-box">
<b>Escala de puntaje</b><br>
1 · Muy Bajo<br>
2 · Bajo<br>
3 · Medio<br>
4 · Alto<br>
5 · Muy Alto<br><br>
<b>Clasificación</b><br>
≥ 4.5 · Core Premium<br>
≥ 3.5 · Strategic Value<br>
≥ 2.5 · Supporting Asset<br>
&lt; 2.5 · Low Priority
</div>
""", unsafe_allow_html=True)

render_main_header()

if uploaded_file:
    input_df = pd.read_excel(uploaded_file)
    base_df = normalize_input(input_df)
else:
    base_df = pd.DataFrame(DEFAULT_ASSETS, columns=["Asset", *SCORE_COLUMNS])

st.subheader("Tabla editable de assets")
st.write("Usá los desplegables para puntuar cada asset. El modelo recalcula automáticamente score, clasificación y valor económico.")

editor_df = editor_dataframe(base_df, expert_mode)
if expert_mode:
    column_config = {"Asset": st.column_config.TextColumn("Asset", width="large", required=True)}
    column_config.update({col: st.column_config.NumberColumn(col, min_value=1, max_value=5, step=1, format="%d") for col in SCORE_COLUMNS})
else:
    column_config = {"Asset": st.column_config.TextColumn("Asset", width="large", required=True)}
    column_config.update({col: st.column_config.SelectboxColumn(col, options=list(SCORE_LABELS.values()), required=True) for col in SCORE_COLUMNS})

edited_df = st.data_editor(editor_df, use_container_width=True, num_rows="dynamic", column_config=column_config, hide_index=True)
result_df = recalculate(edited_df, contract_value)

core_value = result_df.loc[result_df["Clasificación"] == "Core Premium", "Valor Económico"].sum()
avg_score = result_df["Score Final"].mean()
strategic_assets = len(result_df[result_df["Score Final"] >= 3.5])

m1, m2, m3, m4 = st.columns(4)
m1.metric("Valor contrato", f"{currency} {contract_value:,.1f} M")
m2.metric("Assets", f"{len(result_df)}")
m3.metric("Score promedio", f"{avg_score:.2f}")
m4.metric("Core Premium", f"{currency} {core_value:,.1f} M")

st.subheader("Resumen ejecutivo recalculado")
st.write("Vista limpia para negocio: sin decimales técnicos, con escala cualitativa y valor económico en millones.")
output_view = make_output_table(result_df, currency)
st.dataframe(output_view, use_container_width=True, hide_index=True)

chart_col1, chart_col2 = st.columns(2)
with chart_col1:
    st.subheader("Valor por asset")
    fig_asset = px.bar(result_df.sort_values("Valor Económico", ascending=True), x="Valor Económico", y="Asset", orientation="h", color="Clasificación", color_discrete_map=CLASSIFICATION_COLORS, text="Valor Económico", labels={"Valor Económico": f"Valor {currency} M"})
    fig_asset.update_traces(texttemplate="%{text:.2f} M")
    fig_asset.update_layout(height=560, margin=dict(l=10, r=10, t=30, b=10), legend_title_text="Clasificación")
    st.plotly_chart(fig_asset, use_container_width=True)

with chart_col2:
    st.subheader("Valor por clasificación")
    class_df = result_df.groupby("Clasificación", as_index=False)["Valor Económico"].sum()
    class_df["Clasificación"] = pd.Categorical(class_df["Clasificación"], categories=CLASSIFICATION_ORDER, ordered=True)
    class_df = class_df.sort_values("Clasificación")
    fig_class = px.pie(class_df, names="Clasificación", values="Valor Económico", hole=0.55, color="Clasificación", color_discrete_map=CLASSIFICATION_COLORS)
    fig_class.update_traces(textinfo="percent+label")
    fig_class.update_layout(height=560, margin=dict(l=10, r=10, t=30, b=10))
    st.plotly_chart(fig_class, use_container_width=True)

st.subheader("Escenarios")
scenario_df = pd.DataFrame([{"Escenario": name, "Multiplicador": mult, f"Valor Contrato {currency} M": round(contract_value * mult, 2)} for name, mult in SCENARIO_MULTIPLIERS.items()])
st.dataframe(scenario_df, use_container_width=True, hide_index=True)
fig_scenario = px.bar(scenario_df, x="Escenario", y=f"Valor Contrato {currency} M", text=f"Valor Contrato {currency} M")
fig_scenario.update_traces(texttemplate="%{text:.1f} M")
fig_scenario.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10))
st.plotly_chart(fig_scenario, use_container_width=True)

with st.expander("Definición de variables"):
    for var, desc in VARIABLE_DEFINITIONS.items():
        st.markdown(f"**{var}:** {desc}")

with st.expander("Cómo se calcula"):
    st.markdown("""
**Score Final** = promedio simple de Masividad, Construcción de Marca, Potencial de Negocio, Valor Agregado y Uso Esperado.  
**Valor Económico** = `(Score Final / suma total de scores) × valor total del contrato`.  
**Clasificación** se asigna por rangos: Core Premium, Strategic Value, Supporting Asset o Low Priority.
""")

st.subheader("Exportables")
excel_bytes = make_excel_download(result_df, client_name, project_name, currency, contract_value)
st.download_button("Descargar Excel actualizado", data=excel_bytes, file_name=f"valuation_{project_name.lower().replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
pdf_bytes = make_pdf_download(result_df, client_name, project_name, currency, contract_value)
if pdf_bytes:
    st.download_button("Descargar PDF ejecutivo simple", data=pdf_bytes, file_name=f"valuation_{project_name.lower().replace(' ', '_')}.pdf", mime="application/pdf")
else:
    st.info("Para exportar PDF instalá reportlab: pip install reportlab")
