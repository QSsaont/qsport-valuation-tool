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

# -----------------------------------------------------------------------------
# Page config + white label Streamlit UI
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="QSport Valuation Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.set_option("client.showErrorDetails", False)

HIDE_STREAMLIT_UI = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
[data-testid="stToolbar"] {display: none !important;}
[data-testid="stDecoration"] {display: none !important;}
[data-testid="stStatusWidget"] {display: none !important;}
[data-testid="stHeader"] {display: none !important;}
.viewerBadge_container__1QSob {display: none !important;}
.viewerBadge_link__1S137 {display: none !important;}
.block-container {
    padding-top: 2.0rem;
    padding-bottom: 2.5rem;
    max-width: 1500px;
}
section[data-testid="stSidebar"] {
    background: #F3F4F6;
    border-right: 1px solid #E5E7EB;
}
section[data-testid="stSidebar"] > div {
    padding-top: 2rem;
}
h1, h2, h3 {
    letter-spacing: -0.035em;
    color: #111827;
}
p, label, span, div {
    font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}
.qs-header {
    display: flex;
    align-items: center;
    gap: 32px;
    padding: 12px 0 28px 0;
    border-bottom: 1px solid #D1D5DB;
    margin-bottom: 36px;
}
.qs-meli-logo {
    width: 260px;
    min-width: 220px;
    display: flex;
    align-items: center;
}
.qs-divider {
    height: 82px;
    width: 2px;
    background: #111827;
    opacity: 0.85;
}
.qs-title h1 {
    font-size: 48px;
    line-height: 1.05;
    margin: 0 0 14px 0;
    font-weight: 800;
    color: #111827;
}
.qs-title p {
    margin: 0;
    font-size: 16px;
    color: #6B7280;
}
.qs-sidebar-logo {
    margin-bottom: 24px;
}
.qs-sidebar-note {
    color: #6B7280;
    font-size: 14px;
    line-height: 1.5;
    margin-bottom: 20px;
}
.qs-methodology {
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 14px;
    padding: 14px 16px;
    margin-top: 14px;
}
.qs-methodology h4 {
    margin: 0 0 10px 0;
    color: #111827;
    font-size: 15px;
}
.qs-methodology p {
    margin: 4px 0;
    color: #4B5563;
    font-size: 13px;
}
.stDataFrame, [data-testid="stDataFrame"] {
    border: 1px solid #E5E7EB;
    border-radius: 12px;
}
div[data-testid="stMetricValue"] {
    font-size: 1.85rem;
    color: #111827;
}
@media (max-width: 900px) {
    .qs-header {flex-direction: column; align-items: flex-start; gap: 18px;}
    .qs-divider {display:none;}
    .qs-title h1 {font-size: 34px;}
    .qs-meli-logo {width: 220px;}
}
</style>
"""
st.markdown(HIDE_STREAMLIT_UI, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
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
SCORE_OPTIONS = list(SCORE_LABELS.values())
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

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def asset_path(filename: str) -> Path | None:
    candidates = [Path("assets") / filename, Path(filename), Path(".") / filename]
    for path in candidates:
        if path.exists():
            return path
    return None


def image_html(path: Path, width: int | None = None) -> str:
    import base64
    suffix = path.suffix.lower().replace(".", "")
    mime = "jpeg" if suffix in {"jpg", "jpeg"} else "png"
    encoded = base64.b64encode(path.read_bytes()).decode()
    style = f"width:{width}px; max-width:100%; height:auto;" if width else "max-width:100%; height:auto;"
    return f'<img src="data:image/{mime};base64,{encoded}" style="{style}" />'


def default_dataframe() -> pd.DataFrame:
    return pd.DataFrame(DEFAULT_ASSETS, columns=["Asset", *SCORE_COLUMNS])


def clean_score(value) -> float:
    if isinstance(value, str):
        value = value.strip()
        if value in LABEL_TO_SCORE:
            return float(LABEL_TO_SCORE[value])
        if "·" in value:
            value = value.split("·", 1)[0].strip()
        elif "-" in value:
            value = value.split("-", 1)[0].strip()
    try:
        value = float(value)
    except Exception:
        return 1.0
    return min(max(value, 1.0), 5.0)


def score_to_label(value) -> str:
    score = int(round(clean_score(value)))
    return SCORE_LABELS.get(score, SCORE_LABELS[1])


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
        return default_dataframe()
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


def make_display_table(df: pd.DataFrame, currency: str) -> pd.DataFrame:
    display = df.copy()
    for col in SCORE_COLUMNS:
        display[col] = display[col].apply(score_to_label)
    display["Score Final"] = display["Score Final"].map(lambda x: f"{x:.1f}")
    display["Valor Económico"] = display["Valor Económico"].map(lambda x: f"{currency} {x:,.2f} M")
    return display


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
    border = Border(
        left=Side(style="thin", color="E5E7EB"),
        right=Side(style="thin", color="E5E7EB"),
        top=Side(style="thin", color="E5E7EB"),
        bottom=Side(style="thin", color="E5E7EB"),
    )

    ws.merge_cells("A1:I1")
    ws["A1"] = "Sponsorship / Naming Rights Valuation Model"
    ws["A1"].font = Font(bold=True, size=16, color=dark)
    ws["A1"].fill = PatternFill("solid", fgColor=yellow)
    ws["A1"].alignment = Alignment(horizontal="center")

    meta = [
        ["Cliente", client_name],
        ["Proyecto", project_name],
        ["Moneda", currency],
        ["Valor total contrato", contract_value],
        ["Fecha exportación", datetime.now().strftime("%Y-%m-%d %H:%M")],
    ]
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

    widths = [42, 20, 24, 24, 20, 20, 14, 22, 22]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A11"

    ws2["A1"] = "Referencia metodológica"
    ws2["A1"].font = Font(bold=True, size=15, color=dark)
    methodology_rows = [
        ["Score Final", "Clasificación"],
        [">= 4.5", "Core Premium"],
        [">= 3.5", "Strategic Value"],
        [">= 2.5", "Supporting Asset"],
        ["< 2.5", "Low Priority"],
        [],
        ["Escala", "Descripción"],
        ["1", "Muy Bajo"],
        ["2", "Bajo"],
        ["3", "Medio"],
        ["4", "Alto"],
        ["5", "Muy Alto"],
    ]
    for r, row in enumerate(methodology_rows, start=3):
        for c, value in enumerate(row, start=1):
            ws2.cell(r, c).value = value
            ws2.cell(r, c).border = border
            if r in [3, 9]:
                ws2.cell(r, c).font = Font(bold=True, color=white)
                ws2.cell(r, c).fill = PatternFill("solid", fgColor=dark)
    ws2.column_dimensions["A"].width = 25
    ws2.column_dimensions["B"].width = 35

    scenario_rows = [["Escenario", "Multiplicador", f"Valor Contrato {currency} M"]]
    scenario_rows += [[name, mult, round(contract_value * mult, 2)] for name, mult in SCENARIO_MULTIPLIERS.items()]
    for r, row in enumerate(scenario_rows, start=1):
        for c, value in enumerate(row, start=1):
            ws3.cell(r, c).value = value
            ws3.cell(r, c).border = border
            if r == 1:
                ws3.cell(r, c).font = Font(bold=True, color=white)
                ws3.cell(r, c).fill = PatternFill("solid", fgColor=dark)
    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 18
    ws3.column_dimensions["C"].width = 24

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def make_pdf_download(df: pd.DataFrame, client_name: str, project_name: str, currency: str, contract_value: float) -> bytes | None:
    if not REPORTLAB_AVAILABLE:
        return None
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=32, leftMargin=32, topMargin=32, bottomMargin=32)
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph("Sponsorship / Naming Rights Valuation", styles["Title"]))
    story.append(Paragraph(f"Cliente: {client_name} | Proyecto: {project_name}", styles["Normal"]))
    story.append(Paragraph(f"Valor total contrato: {currency} {contract_value:,.2f} M", styles["Normal"]))
    story.append(Spacer(1, 14))

    summary = df.groupby("Clasificación", as_index=False)["Valor Económico"].sum().sort_values("Valor Económico", ascending=False)
    summary_data = [["Clasificación", f"Valor {currency} M"]] + [
        [r["Clasificación"], f"{r['Valor Económico']:,.2f}"] for _, r in summary.iterrows()
    ]
    t = Table(summary_data, colWidths=[260, 140])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#111827")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D1D5DB")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
    ]))
    story.append(t)
    story.append(Spacer(1, 14))

    top = df.sort_values("Valor Económico", ascending=False).head(10)
    table_data = [["Asset", "Score", "Clasificación", f"Valor {currency} M"]]
    for _, r in top.iterrows():
        table_data.append([r["Asset"], f"{r['Score Final']:.2f}", r["Clasificación"], f"{r['Valor Económico']:,.2f}"])
    t2 = Table(table_data, colWidths=[220, 55, 125, 95])
    t2.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#3483FA")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D1D5DB")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
    ]))
    story.append(Paragraph("Top assets por valor económico", styles["Heading2"]))
    story.append(t2)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# -----------------------------------------------------------------------------
# Sidebar
# -----------------------------------------------------------------------------
with st.sidebar:
    qsport_logo = asset_path("logo_qsport.png")
    if qsport_logo:
        st.markdown(f'<div class="qs-sidebar-logo">{image_html(qsport_logo, width=250)}</div>', unsafe_allow_html=True)
    else:
        st.markdown("**QSport · ZETABÉ**")

    st.markdown('<div class="qs-sidebar-note">Modelo desarrollado por QSport ZETABÉ</div>', unsafe_allow_html=True)
    st.divider()

    st.header("Inputs globales")
    client_name = st.text_input("Nombre cliente", "Mercado Livre")
    project_name = st.text_input("Nombre proyecto", "Mercado Livre Arena Pacaembu")
    currency = st.selectbox("Moneda", ["USD", "BRL", "EUR"], index=0)
    contract_value = st.number_input("Valor total contrato (en millones)", min_value=0.0, value=10.0, step=0.5, format="%.2f")
    uploaded_file = st.file_uploader("Cargar Excel de assets", type=["xlsx"])
    expert_mode = st.toggle("Modo experto: editar números manualmente", value=False)

    st.markdown(
        """
        <div class="qs-methodology">
            <h4>Referencia metodológica</h4>
            <p><b>Score Final</b> = promedio de las 5 variables.</p>
            <p><b>≥ 4.5:</b> Core Premium</p>
            <p><b>≥ 3.5:</b> Strategic Value</p>
            <p><b>≥ 2.5:</b> Supporting Asset</p>
            <p><b>&lt; 2.5:</b> Low Priority</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# -----------------------------------------------------------------------------
# Header
# -----------------------------------------------------------------------------
meli_logo = asset_path("logo_meli.png")
if meli_logo:
    logo_markup = image_html(meli_logo, width=250)
else:
    logo_markup = '<div style="background:#FFE600;border-radius:24px;padding:18px 34px;font-weight:800;color:#2D3277;">Mercado Livre</div>'

st.markdown(
    f"""
    <div class="qs-header">
        <div class="qs-meli-logo">{logo_markup}</div>
        <div class="qs-divider"></div>
        <div class="qs-title">
            <h1>Sponsorship / Naming Rights Valuation Tool</h1>
            <p>Modelo de scoring estratégico para valorar assets comerciales y redistribuir el valor económico del contrato.</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Data load
# -----------------------------------------------------------------------------
if uploaded_file:
    input_df = pd.read_excel(uploaded_file)
    base_df = normalize_input(input_df)
else:
    base_df = default_dataframe()

# -----------------------------------------------------------------------------
# Editable table
# -----------------------------------------------------------------------------
st.subheader("Tabla editable de assets")
st.write("Usá los desplegables para puntuar cada asset. El modelo recalcula automáticamente score, clasificación y valor económico.")

if expert_mode:
    edit_df = base_df.copy()
    edited_df = st.data_editor(
        edit_df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Asset": st.column_config.TextColumn("Asset", width="large", required=True),
            **{col: st.column_config.NumberColumn(col, min_value=1, max_value=5, step=1, format="%d") for col in SCORE_COLUMNS},
        },
        hide_index=True,
        key="expert_editor",
    )
else:
    edit_df = base_df.copy()
    for col in SCORE_COLUMNS:
        edit_df[col] = edit_df[col].apply(score_to_label)
    edited_df = st.data_editor(
        edit_df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Asset": st.column_config.TextColumn("Asset", width="large", required=True),
            **{col: st.column_config.SelectboxColumn(col, options=SCORE_OPTIONS, required=True) for col in SCORE_COLUMNS},
        },
        hide_index=True,
        key="dropdown_editor",
    )

result_df = recalculate(edited_df, contract_value)

# -----------------------------------------------------------------------------
# KPIs
# -----------------------------------------------------------------------------
core_value = result_df.loc[result_df["Clasificación"] == "Core Premium", "Valor Económico"].sum()
avg_score = result_df["Score Final"].mean() if not result_df.empty else 0

m1, m2, m3, m4 = st.columns(4)
m1.metric("Valor contrato", f"{currency} {contract_value:,.1f} M")
m2.metric("Assets", f"{len(result_df)}")
m3.metric("Score promedio", f"{avg_score:.2f}")
m4.metric("Core Premium", f"{currency} {core_value:,.1f} M")

# -----------------------------------------------------------------------------
# Output table
# -----------------------------------------------------------------------------
st.subheader("Resumen valorizado")
st.caption("Tabla limpia para lectura ejecutiva. Los puntajes se muestran como etiquetas, no como decimales técnicos.")
st.dataframe(make_display_table(result_df, currency), use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# Charts
# -----------------------------------------------------------------------------
chart_col1, chart_col2 = st.columns(2)
with chart_col1:
    st.subheader("Valor por asset")
    fig_asset = px.bar(
        result_df.sort_values("Valor Económico", ascending=True),
        x="Valor Económico",
        y="Asset",
        orientation="h",
        color="Clasificación",
        color_discrete_map=CLASSIFICATION_COLORS,
        text="Valor Económico",
        labels={"Valor Económico": f"Valor {currency} M"},
    )
    fig_asset.update_traces(texttemplate="%{text:.2f} M", textposition="outside")
    fig_asset.update_layout(height=540, margin=dict(l=10, r=10, t=30, b=10), legend_title_text="Clasificación")
    st.plotly_chart(fig_asset, use_container_width=True)

with chart_col2:
    st.subheader("Valor por clasificación")
    class_df = result_df.groupby("Clasificación", as_index=False)["Valor Económico"].sum()
    class_df["Clasificación"] = pd.Categorical(class_df["Clasificación"], categories=CLASSIFICATION_ORDER, ordered=True)
    class_df = class_df.sort_values("Clasificación")
    fig_class = px.pie(
        class_df,
        names="Clasificación",
        values="Valor Económico",
        hole=0.55,
        color="Clasificación",
        color_discrete_map=CLASSIFICATION_COLORS,
    )
    fig_class.update_layout(height=540, margin=dict(l=10, r=10, t=30, b=10))
    st.plotly_chart(fig_class, use_container_width=True)

# -----------------------------------------------------------------------------
# Scenarios
# -----------------------------------------------------------------------------
st.subheader("Escenarios")
scenario_df = pd.DataFrame([
    {"Escenario": name, "Multiplicador": mult, f"Valor Contrato {currency} M": round(contract_value * mult, 2)}
    for name, mult in SCENARIO_MULTIPLIERS.items()
])
st.dataframe(scenario_df, use_container_width=True, hide_index=True)

fig_scenario = px.bar(
    scenario_df,
    x="Escenario",
    y=f"Valor Contrato {currency} M",
    text=f"Valor Contrato {currency} M",
)
fig_scenario.update_traces(texttemplate="%{text:.2f} M", textposition="outside")
fig_scenario.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10), showlegend=False)
st.plotly_chart(fig_scenario, use_container_width=True)

# -----------------------------------------------------------------------------
# Methodology + exports
# -----------------------------------------------------------------------------
with st.expander("Definición de variables"):
    for var, desc in VARIABLE_DEFINITIONS.items():
        st.markdown(f"**{var}:** {desc}")

with st.expander("Clasificación estratégica"):
    st.table(pd.DataFrame([
        {"Score Final": "≥ 4.5", "Clasificación": "Core Premium"},
        {"Score Final": "≥ 3.5", "Clasificación": "Strategic Value"},
        {"Score Final": "≥ 2.5", "Clasificación": "Supporting Asset"},
        {"Score Final": "< 2.5", "Clasificación": "Low Priority"},
    ]))

st.subheader("Exportables")
excel_bytes = make_excel_download(result_df, client_name, project_name, currency, contract_value)
st.download_button(
    "Descargar Excel actualizado",
    data=excel_bytes,
    file_name=f"valuation_{project_name.lower().replace(' ', '_').replace('/', '-')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

pdf_bytes = make_pdf_download(result_df, client_name, project_name, currency, contract_value)
if pdf_bytes:
    st.download_button(
        "Descargar PDF ejecutivo simple",
        data=pdf_bytes,
        file_name=f"valuation_{project_name.lower().replace(' ', '_').replace('/', '-')}.pdf",
        mime="application/pdf",
    )
else:
    st.info("Para exportar PDF instalá reportlab: pip install reportlab")
