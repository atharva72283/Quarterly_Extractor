import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import os
import base64
import tempfile
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import fitz  # PyMuPDF
import google.generativeai as genai

# ─────────────────────────────────────────
# Page Config
# ─────────────────────────────────────────
st.set_page_config(
    page_title="JM Financial – Quarterly Results Extractor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────
# Custom CSS  – JM Financial Gold Theme
# ─────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #1a1a2e;
    border-right: 2px solid #FCB316;
}
section[data-testid="stSidebar"] * {
    color: #e8e8e8 !important;
}
section[data-testid="stSidebar"] .stTextInput input,
section[data-testid="stSidebar"] .stFileUploader {
    background: #16213e !important;
    border: 1px solid #FCB316 !important;
    color: #fff !important;
    border-radius: 4px;
}

/* Main area */
.main .block-container { padding-top: 2rem; max-width: 1400px; }

/* Header banner */
.jm-header {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    border-left: 5px solid #FCB316;
    padding: 1.2rem 1.8rem;
    border-radius: 6px;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    gap: 1rem;
}
.jm-header h1 { color: #FCB316; font-size: 1.6rem; font-weight: 700; margin: 0; }
.jm-header p { color: #aaa; font-size: 0.85rem; margin: 0; }

/* Step badges */
.step-badge {
    display: inline-block;
    background: #FCB316;
    color: #1a1a2e;
    font-weight: 700;
    font-size: 0.75rem;
    padding: 2px 10px;
    border-radius: 20px;
    margin-bottom: 0.4rem;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 0.05em;
}

/* Card sections */
.card {
    background: #f9f9f9;
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 1.2rem 1.5rem;
    margin-bottom: 1rem;
}

/* Gold buttons */
div.stButton > button {
    background: #FCB316;
    color: #1a1a2e;
    font-weight: 700;
    border: none;
    border-radius: 4px;
    padding: 0.5rem 1.5rem;
    font-family: 'IBM Plex Sans', sans-serif;
    letter-spacing: 0.04em;
    transition: all 0.2s;
}
div.stButton > button:hover {
    background: #e5a110;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(252,179,22,0.35);
}

/* Success / info boxes */
.status-ok {
    background: #e8f5e9; border-left: 4px solid #4caf50;
    padding: 0.6rem 1rem; border-radius: 4px;
    font-size: 0.9rem; margin: 0.5rem 0;
}
.status-warn {
    background: #fff8e1; border-left: 4px solid #FCB316;
    padding: 0.6rem 1rem; border-radius: 4px;
    font-size: 0.9rem; margin: 0.5rem 0;
}

/* Data editor */
.stDataEditor { border: 2px solid #FCB316; border-radius: 6px; }

/* Download button */
div.stDownloadButton > button {
    background: #1a1a2e; color: #FCB316;
    border: 2px solid #FCB316; font-weight: 700;
    border-radius: 4px;
}
div.stDownloadButton > button:hover {
    background: #FCB316; color: #1a1a2e;
}

/* Number metric cards */
.metric-row { display: flex; gap: 0.8rem; flex-wrap: wrap; margin: 0.8rem 0; }
.metric-card {
    background: #1a1a2e; border-radius: 6px;
    padding: 0.7rem 1.2rem; flex: 1; min-width: 130px;
    border-bottom: 3px solid #FCB316;
}
.metric-card .label { color: #aaa; font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.08em; }
.metric-card .value { color: #FCB316; font-size: 1.3rem; font-weight: 700; font-family: 'IBM Plex Mono', monospace; }
.metric-card .sub { color: #888; font-size: 0.72rem; margin-top: 1px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# Header
# ─────────────────────────────────────────
st.markdown("""
<div class="jm-header">
    <div>
        <h1>📊 Quarterly Results Extractor</h1>
        <p>JM Financial · AI-Powered PDF → Excel & Word Automation</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# Sidebar – Credentials & Inputs
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔑 API Configuration")
    api_key = st.text_input(
        "Google Gemini API Key",
        type="password",
        placeholder="AIza...",
        help="Your Gemini 2.5 Flash API key"
    )

    st.markdown("---")
    st.markdown("### 📋 Report Settings")

    company_name = st.text_input(
        "Company Name",
        placeholder="e.g. Reliance Industries Ltd.",
        value=""
    )

    page_number = st.number_input(
        "PDF Page to Analyze",
        min_value=1, max_value=500,
        value=1,
        step=1,
        help="The page number (1-indexed) containing the financial table"
    )

    st.markdown("---")
    st.markdown("### 🖼️ Header Logos *(optional)*")
    left_logo_file  = st.file_uploader("Left logo  (Q4.png)",   type=["png","jpg","jpeg"], key="left_logo")
    right_logo_file = st.file_uploader("Right logo (JM Logo)",  type=["png","jpg","jpeg"], key="right_logo")

    st.markdown("---")
    st.caption("v2.0 · JM Financial Internal Tool")

# ─────────────────────────────────────────
# Session-state init
# ─────────────────────────────────────────
for key in ["df_extracted", "gemini_raw", "step"]:
    if key not in st.session_state:
        st.session_state[key] = None
if "step" not in st.session_state or st.session_state.step is None:
    st.session_state.step = 1

# ─────────────────────────────────────────
# Helper functions
# ─────────────────────────────────────────
COLS_NUMERIC = ['Q4FY2026', 'Q3FY2026', 'Q4FY2025', 'FY2026', 'FY2025']

def extract_pdf_page_as_image(pdf_bytes: bytes, page_idx: int) -> Image.Image:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    if page_idx >= len(doc):
        st.error(f"PDF only has {len(doc)} pages. Please choose a valid page.")
        st.stop()
    page = doc[page_idx]
    pix  = page.get_pixmap(dpi=200)
    return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)


def call_gemini(api_key: str, img: Image.Image) -> str:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
    prompt = """
Extract the financial data from the provided image into a clean CSV format.
Columns MUST exactly match: Particulars, Q4FY2026, Q3FY2026, Q4FY2025, FY2026, FY2025.

Extract or calculate the following rows exactly in this order:
1.  Revenue from Operations
2.  Other Incomes
3.  Total Income         -> (Calculate: Revenue from Operations + Other Incomes)
4.  Total Expenditure    -> (Calculate: Total Expenses - Interest - Depreciation)
5.  EBITDA               -> (Calculate: Total Income - Total Expenditure)
6.  Interest             -> (Finance cost)
7.  Depreciation         -> (Depreciation and amortisation)
8.  PBT                  -> (Profit Before Tax)
9.  Tax                  -> (Total tax expense)
10. Provisions & Contingencies
11. Share of Profit of Associate/JV
12. Exceptional Items
13. PAT                  -> (Profit After Tax / Net Profit for the period)
14. EPS                  -> (Earnings Per Share - Basic/Diluted)

IMPORTANT: Return ONLY the CSV block. No intro, no outro, no explanations.
Use 0 for missing values. Clean numbers of all currency symbols and commas.
Indicate negative numbers with a minus sign (e.g., -100).
"""
    response = model.generate_content([prompt, img])
    return response.text.strip()


def parse_gemini_csv(raw: str) -> pd.DataFrame:
    csv_match = re.search(r"(Particulars,.*)", raw, re.DOTALL)
    clean_csv = (
        csv_match.group(1).replace("```", "").strip()
        if csv_match
        else raw.replace("```csv", "").replace("```", "").strip()
    )
    df = pd.read_csv(io.StringIO(clean_csv), skipinitialspace=True)
    for col in COLS_NUMERIC:
        if col in df.columns:
            df[col] = (
                pd.to_numeric(
                    df[col]
                    .astype(str)
                    .str.replace(",", "", regex=False)
                    .str.replace(r"\(", "-", regex=True)
                    .str.replace(r"\)", "",  regex=True),
                    errors="coerce",
                ).fillna(0)
            )
    # Growth columns
    df["QoQ%"] = ((df["Q4FY2026"] - df["Q3FY2026"]) / df["Q3FY2026"].replace(0, np.nan) * 100).fillna(0).round(0).astype(int)
    df["YoY%"] = ((df["Q4FY2026"] - df["Q4FY2025"]) / df["Q4FY2025"].replace(0, np.nan) * 100).fillna(0).round(0).astype(int)
    df["% Change"] = ((df["FY2026"] - df["FY2025"]) / df["FY2025"].replace(0, np.nan) * 100).fillna(0).round(0).astype(int)
    return df[["Particulars", "Q4FY2026", "Q3FY2026", "Q4FY2025", "QoQ%", "YoY%", "FY2026", "FY2025", "% Change"]]


def get_row(df: pd.DataFrame, label: str) -> pd.Series:
    row = df[df["Particulars"].str.strip() == label]
    return row.iloc[0] if not row.empty else pd.Series({c: 0 for c in df.columns})


def build_template_df(df: pd.DataFrame) -> pd.DataFrame:
    net_sales = get_row(df, "Revenue from Operations")
    other_inc = get_row(df, "Other Incomes")
    tot_inc   = get_row(df, "Total Income")
    tot_exp   = get_row(df, "Total Expenditure")
    ebitda    = get_row(df, "EBITDA")
    interest  = get_row(df, "Interest")
    dep       = get_row(df, "Depreciation")
    pbt       = get_row(df, "PBT")
    tax       = get_row(df, "Tax")
    prov      = get_row(df, "Provisions & Contingencies")
    exc       = get_row(df, "Exceptional Items")
    pat       = get_row(df, "PAT")
    eps       = get_row(df, "EPS")

    template_rows = [
        "Net Sales", "Other Income", "Total Income", "",
        "Total Expenditure (Ex Int & Dep)", "EBIDTA", "Interest",
        "Dep", "Profit before tax", "Tax", "Provisions & Contingencies",
        "Exceptional Item", "Reported Profit", "EBIDTA Margin %", "PAT %", "EPS (Diluted)"
    ]

    all_cols = ["Q4FY2026", "Q3FY2026", "Q4FY2025", "QoQ%", "YoY%", "FY2026", "FY2025", "% Change"]
    mapped = {c: [] for c in all_cols}

    for col in all_cols:
        ns  = net_sales.get(col, 0)
        oi  = other_inc.get(col, 0)
        ti  = tot_inc.get(col, 0)
        te  = tot_exp.get(col, 0)
        eb  = ebitda.get(col, 0)
        ir  = interest.get(col, 0)
        dp  = dep.get(col, 0)
        pb  = pbt.get(col, 0)
        tx  = tax.get(col, 0)
        pv  = prov.get(col, 0)
        ex  = exc.get(col, 0)
        pt  = pat.get(col, 0)
        ep  = eps.get(col, 0)
        eb_m = (eb / ns * 100) if ns != 0 else 0
        pt_m = (pt / ns * 100) if ns != 0 else 0

        mapped[col] = [ns, oi, ti, "", te, eb, ir, dp, pb, tx, pv, ex, pt, eb_m, pt_m, ep]

    tdf = pd.DataFrame({"Particulars": template_rows})
    for col in all_cols:
        tdf[col] = mapped[col]
    return tdf


def generate_paragraphs(df: pd.DataFrame, company: str) -> list[str]:
    ns  = get_row(df, "Revenue from Operations")
    ti  = get_row(df, "Total Income")
    eb  = get_row(df, "EBITDA")
    pat = get_row(df, "PAT")
    eps = get_row(df, "EPS")

    tdf = build_template_df(df)

    def gv(row_label, col):
        r = tdf[tdf["Particulars"] == row_label]
        return float(r.iloc[0][col]) if not r.empty else 0

    ebitda_margin_q4 = gv("EBIDTA Margin %", "Q4FY2026")
    ebitda_margin_q3 = gv("EBIDTA Margin %", "Q3FY2026")
    pat_margin_q4    = gv("PAT %",           "Q4FY2026")
    pat_margin_q3    = gv("PAT %",           "Q3FY2026")
    ebitda_bps = (ebitda_margin_q4 - ebitda_margin_q3) * 100
    pat_bps    = (pat_margin_q4    - pat_margin_q3)    * 100

    fn  = lambda v: f"{float(v):,.2f}"
    fp  = lambda v: f"{abs(float(v)):.0f}"
    ud  = lambda v: "Increased" if float(v) >= 0 else "Decreased"
    udl = lambda v: "increased" if float(v) >= 0 else "decreased"

    p1 = (f"{company} reported net profit of INR {fn(pat['Q4FY2026'])} Crs. in Q4FY2026, "
          f"{ud(pat['QoQ%'])} by ~{fp(pat['QoQ%'])}% QoQ & "
          f"{ud(pat['YoY%'])} by ~{fp(pat['YoY%'])}% YoY "
          f"(Margin percentage {fn(pat_margin_q4)}%).")

    p2 = (f"Company Reported Total Income of INR {fn(ti['Q4FY2026'])} Crs. in Q4FY2026 "
          f"as compared to INR {fn(ti['Q3FY2026'])} Crs. in previous quarter Q3FY2026. "
          f"Total Income has {udl(ti['QoQ%'])} by {fp(ti['QoQ%'])}% QoQ. "
          f"(Total Income of INR {fn(ti['Q4FY2025'])} Crs. in Q4FY2025 "
          f"{udl(ti['YoY%'])} by {fp(ti['YoY%'])}% YoY) "
          f"For FY2026, Company has Reported Total Income of INR {fn(ti['FY2026'])} Crs. "
          f"as against INR {fn(ti['FY2025'])} Crs. in FY2025 "
          f"{udl(ti['% Change'])} by ~{fp(ti['% Change'])}%.")

    p3 = (f"Company Reported EBITDA of INR {fn(eb['Q4FY2026'])} Crs. for Q4FY2026 "
          f"as compared to INR {fn(eb['Q3FY2026'])} in previous quarter Q3FY2026. "
          f"EBITDA has {udl(eb['QoQ%'])} by {fp(eb['QoQ%'])}% QoQ. "
          f"(EBIDTA of INR {fn(eb['Q4FY2025'])} Crs. in Q4FY2025 "
          f"{udl(eb['YoY%'])} by {fp(eb['YoY%'])}% YoY) "
          f"EBIDTA margin stood at {fn(ebitda_margin_q4)}% for the quarter "
          f"{udl(ebitda_bps)} by {int(abs(ebitda_bps))} bps. "
          f"For FY2026, Company has Reported EBIDTA of INR {fn(eb['FY2026'])} Crs. "
          f"as against INR {fn(eb['FY2025'])} Crs. in FY2025 "
          f"{ud(eb['% Change'])} by ~{fp(eb['% Change'])}%.")

    p4 = (f"Company Reported Profit of INR {fn(pat['Q4FY2026'])} Crs. for Q4FY2026 "
          f"as compared to INR {fn(pat['Q3FY2026'])} Crs. in previous quarter Q3FY2026. "
          f"Profit has {udl(pat['QoQ%'])} by {fp(pat['QoQ%'])}% QoQ. "
          f"(Profit of INR {fn(pat['Q4FY2025'])} Crs. in Q4FY2025 "
          f"{udl(pat['YoY%'])} by {fp(pat['YoY%'])}% YoY.) "
          f"Profit margin stood at {fn(pat_margin_q4)}% for the quarter "
          f"{udl(pat_bps)} by {int(abs(pat_bps))} bps. "
          f"For FY2026, Company has Reported Profit of INR {fn(pat['FY2026'])} Crs. "
          f"as against INR {fn(pat['FY2025'])} Crs. in FY2025 "
          f"{ud(pat['% Change'])} by ~{fp(pat['% Change'])}%.")

    p5 = (f"Company reported EPS of INR {fn(eps['Q4FY2026'])} "
          f"{udl(eps['QoQ%'])} by {fp(eps['QoQ%'])}% QoQ & "
          f"for FY2026 EPS stood at INR {fn(eps['FY2026'])}.")

    return [p1, p2, p3, p4, p5]


def build_table_image(tdf: pd.DataFrame) -> bytes:
    """Returns PNG bytes of the styled financial table."""
    df_img = tdf.copy()
    for i in range(len(df_img)):
        row_label = str(df_img.iloc[i, 0]).strip()
        is_eps   = row_label == "EPS (Diluted)"
        is_blank = row_label == ""
        for col in df_img.columns:
            if col == "Particulars":
                continue
            val = tdf.loc[i, col]
            is_pct_col = col in ["QoQ%", "YoY%", "% Change"]
            is_pct_row = row_label in ["EBIDTA Margin %", "PAT %"]
            if is_blank:
                df_img.loc[i, col] = ""
            elif pd.isna(val) or val == 0 or val == "":
                df_img.loc[i, col] = "-"
            else:
                try:
                    vf = float(val)
                    if is_pct_col or is_pct_row:
                        vi = int(round(vf))
                        df_img.loc[i, col] = f"({abs(vi)}%)" if vi < 0 else f"{vi}%"
                    elif is_eps:
                        df_img.loc[i, col] = f"({abs(vf):,.2f})" if vf < 0 else f"{vf:,.2f}"
                    else:
                        vi = int(round(vf))
                        df_img.loc[i, col] = f"({abs(vi):,})" if vi < 0 else f"{vi:,}"
                except Exception:
                    df_img.loc[i, col] = str(val)

    HIGHLIGHT = ["Total Income", "EBIDTA", "Profit before tax", "Reported Profit"]

    fig, ax = plt.subplots(figsize=(15, df_img.shape[0] * 0.34 + 0.5))
    ax.axis("off")
    custom_widths = [0.27] + [0.09] * (len(df_img.columns) - 1)
    table = ax.table(
        cellText=df_img.values,
        colLabels=df_img.columns,
        loc="center",
        cellLoc="right",
        colWidths=custom_widths,
    )
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1, 1.4)

    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_facecolor("#FCB316")
            cell.set_text_props(weight="bold", color="black", ha="center")
        else:
            row_label = str(df_img.iloc[row - 1, 0]).strip()
            if col == 0:
                cell.set_text_props(ha="left")
            if row_label in HIGHLIGHT:
                cell.set_facecolor("#EBEBEB")
                if col == 0:
                    cell.set_text_props(weight="bold")
        cell.set_edgecolor("black")

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight", dpi=220)
    plt.close()
    buf.seek(0)
    return buf.read()


def build_word_doc(
    company: str,
    paragraphs: list[str],
    table_png: bytes,
    left_logo_bytes=None,
    right_logo_bytes=None,
) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.top_margin    = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin   = Inches(0.5)
    section.right_margin  = Inches(0.5)

    # ── Header ──────────────────────────────
    header = section.header
    for para in list(header.paragraphs):
        try:
            p = para._element
            p.getparent().remove(p)
        except Exception:
            pass

    htbl = header.add_table(rows=1, cols=2, width=Inches(7.5))
    htbl.autofit = False
    htbl.columns[0].width = Inches(3.75)
    htbl.columns[1].width = Inches(3.75)

    lp = htbl.cell(0, 0).paragraphs[0]
    lp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if left_logo_bytes:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
            tf.write(left_logo_bytes)
            tf.flush()
            lp.add_run().add_picture(tf.name, width=Inches(2.5))

    rp = htbl.cell(0, 1).paragraphs[0]
    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if right_logo_bytes:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
            tf.write(right_logo_bytes)
            tf.flush()
            rp.add_run().add_picture(tf.name, width=Inches(1.5))

    # ── Page border ─────────────────────────
    sectPr    = section._sectPr
    pgBorders = OxmlElement("w:pgBorders")
    pgBorders.set(qn("w:offsetFrom"), "page")
    for bn in ["top", "left", "bottom", "right"]:
        border = OxmlElement(f"w:{bn}")
        border.set(qn("w:val"),   "single")
        border.set(qn("w:sz"),    "4")
        border.set(qn("w:space"), "24")
        border.set(qn("w:color"), "auto")
        pgBorders.append(border)
    sectPr.append(pgBorders)

    # ── Paragraphs ───────────────────────────
    for i, text in enumerate(paragraphs):
        p    = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run  = p.add_run(text)
        run.font.name = "Segoe UI"
        run.font.size = Pt(11)
        if i == 0:
            run.bold = True

    doc.add_paragraph()

    # ── Table image ──────────────────────────
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
        tf.write(table_png)
        tf.flush()
        doc.add_picture(tf.name, width=Inches(7.5))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def build_excel(df: pd.DataFrame, company: str) -> bytes:
    """Minimal Excel output – relies only on openpyxl (no xlsxwriter needed)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    tdf = build_template_df(df)
    wb  = Workbook()
    ws  = wb.active
    ws.title = "Financials"

    GOLD        = "FCB316"
    LIGHT_GREY  = "EBEBEB"
    thin        = Side(style="thin")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)
    HIGHLIGHT   = {"Total Income", "EBIDTA", "Profit before tax", "Reported Profit"}

    # Title row
    last_col_letter = get_column_letter(len(tdf.columns))
    ws.merge_cells(f"A1:{last_col_letter}1")
    tc = ws["A1"]
    tc.value     = f"{company} – Quarterly Result Update  (Amount in Crs)"
    tc.font      = Font(name="Calibri", size=11, bold=True)
    tc.alignment = Alignment(horizontal="center", vertical="center")
    tc.border    = border

    # Header row (row 2)
    for ci, col_name in enumerate(tdf.columns, start=1):
        cell = ws.cell(row=2, column=ci, value=col_name)
        cell.font      = Font(bold=True, color="000000")
        cell.fill      = PatternFill(start_color=GOLD, end_color=GOLD, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = border

    # Data rows
    NUMBER_FMT  = '#,##0_);[Black](#,##0);"-"'
    PCT_FMT     = '0"%";[Black]\(0"%"\);"-"'

    for ri, row in tdf.iterrows():
        for ci, col_name in enumerate(tdf.columns, start=1):
            val  = row[col_name]
            cell = ws.cell(row=ri + 3, column=ci)
            try:
                cell.value = float(val) if val != "" else None
            except (ValueError, TypeError):
                cell.value = val

            if col_name != "Particulars":
                is_pct = col_name in ["QoQ%", "YoY%", "% Change"] or \
                         str(row["Particulars"]).strip() in ["EBIDTA Margin %", "PAT %"]
                cell.number_format = PCT_FMT if is_pct else NUMBER_FMT
                cell.alignment     = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")

            cell.border = border

            if str(row["Particulars"]).strip() in HIGHLIGHT:
                cell.fill = PatternFill(start_color=LIGHT_GREY, end_color=LIGHT_GREY, fill_type="solid")
                if col_name == "Particulars":
                    cell.font = Font(bold=True)

    # Column widths
    for ci, col_name in enumerate(tdf.columns, start=1):
        cl = get_column_letter(ci)
        if col_name == "Particulars":
            ws.column_dimensions[cl].width = 34
        else:
            ws.column_dimensions[cl].width = 14

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────
# STEP 1 – PDF Upload & Extraction
# ─────────────────────────────────────────
st.markdown('<div class="step-badge">STEP 1</div>', unsafe_allow_html=True)
st.subheader("Upload PDF & Extract Data")

pdf_file = st.file_uploader(
    "Upload the quarterly results PDF",
    type=["pdf"],
    help="The PDF containing the P&L / financial results page"
)

col_btn, col_status = st.columns([1, 3])
with col_btn:
    extract_btn = st.button("⚡ Extract with Gemini", use_container_width=True)

if extract_btn:
    if not api_key:
        st.error("Please enter your Gemini API key in the sidebar.")
    elif not pdf_file:
        st.error("Please upload a PDF file.")
    elif not company_name.strip():
        st.error("Please enter the company name in the sidebar.")
    else:
        with st.spinner(f"Extracting page {page_number} from PDF…"):
            pdf_bytes = pdf_file.read()
            img = extract_pdf_page_as_image(pdf_bytes, page_number - 1)

        with st.spinner("Sending to Gemini 2.5 Flash…"):
            raw = call_gemini(api_key, img)
            st.session_state.gemini_raw = raw

        with st.spinner("Parsing response…"):
            df = parse_gemini_csv(raw)
            st.session_state.df_extracted = df
            st.session_state.step = 2

        st.markdown('<div class="status-ok">✅ Data extracted successfully! Review and edit below.</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────
# STEP 2 – Review & Edit
# ─────────────────────────────────────────
if st.session_state.df_extracted is not None:
    st.markdown("---")
    st.markdown('<div class="step-badge">STEP 2</div>', unsafe_allow_html=True)
    st.subheader("Review & Edit Extracted Data")

    st.caption("Double-click any numeric cell to correct values. Changes are saved automatically.")

    edited_df = st.data_editor(
        st.session_state.df_extracted,
        use_container_width=True,
        num_rows="fixed",
        key="main_editor",
        column_config={
            "Particulars": st.column_config.TextColumn("Particulars", width="medium"),
            "Q4FY2026": st.column_config.NumberColumn("Q4FY2026", format="%.2f"),
            "Q3FY2026": st.column_config.NumberColumn("Q3FY2026", format="%.2f"),
            "Q4FY2025": st.column_config.NumberColumn("Q4FY2025", format="%.2f"),
            "QoQ%":     st.column_config.NumberColumn("QoQ%",     format="%d"),
            "YoY%":     st.column_config.NumberColumn("YoY%",     format="%d"),
            "FY2026":   st.column_config.NumberColumn("FY2026",   format="%.2f"),
            "FY2025":   st.column_config.NumberColumn("FY2025",   format="%.2f"),
            "% Change": st.column_config.NumberColumn("% Change", format="%d"),
        }
    )

    # Quick KPI strip
    try:
        pat_row = get_row(edited_df, "PAT")
        eb_row  = get_row(edited_df, "EBITDA")
        rev_row = get_row(edited_df, "Revenue from Operations")
        eb_mg   = (eb_row["Q4FY2026"] / rev_row["Q4FY2026"] * 100) if rev_row["Q4FY2026"] else 0
        pt_mg   = (pat_row["Q4FY2026"] / rev_row["Q4FY2026"] * 100) if rev_row["Q4FY2026"] else 0

        st.markdown(f"""
        <div class="metric-row">
          <div class="metric-card">
            <div class="label">Revenue (Q4FY26)</div>
            <div class="value">₹{rev_row['Q4FY2026']:,.0f}</div>
            <div class="sub">Crs.</div>
          </div>
          <div class="metric-card">
            <div class="label">EBITDA (Q4FY26)</div>
            <div class="value">₹{eb_row['Q4FY2026']:,.0f}</div>
            <div class="sub">{eb_mg:.1f}% margin</div>
          </div>
          <div class="metric-card">
            <div class="label">PAT (Q4FY26)</div>
            <div class="value">₹{pat_row['Q4FY2026']:,.0f}</div>
            <div class="sub">{pt_mg:.1f}% margin</div>
          </div>
          <div class="metric-card">
            <div class="label">PAT QoQ</div>
            <div class="value">{pat_row['QoQ%']:+.0f}%</div>
            <div class="sub">vs Q3FY26</div>
          </div>
          <div class="metric-card">
            <div class="label">PAT YoY</div>
            <div class="value">{pat_row['YoY%']:+.0f}%</div>
            <div class="sub">vs Q4FY25</div>
          </div>
        </div>
        """, unsafe_allow_html=True)
    except Exception:
        pass

    # Recalculate growth cols whenever editor changes
    for col in COLS_NUMERIC:
        if col not in edited_df.columns:
            edited_df[col] = 0
    edited_df["QoQ%"]    = ((edited_df["Q4FY2026"] - edited_df["Q3FY2026"]) / edited_df["Q3FY2026"].replace(0, np.nan) * 100).fillna(0).round(0).astype(int)
    edited_df["YoY%"]    = ((edited_df["Q4FY2026"] - edited_df["Q4FY2025"]) / edited_df["Q4FY2025"].replace(0, np.nan) * 100).fillna(0).round(0).astype(int)
    edited_df["% Change"]= ((edited_df["FY2026"]   - edited_df["FY2025"])   / edited_df["FY2025"].replace(0, np.nan) * 100).fillna(0).round(0).astype(int)

    # ─────────────────────────────────────────
    # STEP 3 – Generate Outputs
    # ─────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="step-badge">STEP 3</div>', unsafe_allow_html=True)
    st.subheader("Generate & Download Outputs")

    tab_excel, tab_word = st.tabs(["📗  Excel Output", "📝  Word Report"])

    # ── Excel ────────────────────────────────
    with tab_excel:
        st.markdown("Generates a styled Excel file matching the JM Financial quarterly template.")
        if st.button("Build Excel", key="build_excel"):
            with st.spinner("Building Excel…"):
                xl_bytes = build_excel(edited_df, company_name)
            st.download_button(
                label="⬇️  Download Excel",
                data=xl_bytes,
                file_name=f"{company_name} Q4_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ── Word ─────────────────────────────────
    with tab_word:
        st.markdown("Generates the styled Word document with auto-written earnings commentary and coloured table.")

        # Show generated paragraphs for review
        paras = generate_paragraphs(edited_df, company_name)
        st.markdown("**Generated Earnings Summary**")
        for i, p in enumerate(paras):
            st.text_area(f"Paragraph {i+1}", value=p, height=90, key=f"para_{i}")

        st.caption("Edit any paragraph above before generating the Word document.")

        if st.button("Build Word Document", key="build_word"):
            # Collect possibly-edited paras
            final_paras = [st.session_state.get(f"para_{i}", p) for i, p in enumerate(paras)]

            left_logo_bytes  = left_logo_file.read()  if left_logo_file  else None
            right_logo_bytes = right_logo_file.read() if right_logo_file else None

            with st.spinner("Rendering table image…"):
                tdf       = build_template_df(edited_df)
                table_png = build_table_image(tdf)

            with st.spinner("Building Word document…"):
                docx_bytes = build_word_doc(
                    company_name, final_paras, table_png,
                    left_logo_bytes, right_logo_bytes
                )

            st.download_button(
                label="⬇️  Download Word Document",
                data=docx_bytes,
                file_name=f"{company_name} Q4_Result.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

# ─────────────────────────────────────────
# Footer
# ─────────────────────────────────────────
st.markdown("---")
st.caption("JM Financial · Internal Research Automation · For internal use only")
