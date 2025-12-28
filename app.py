import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors

from datetime import datetime
from dateutil.relativedelta import relativedelta
import re

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="Forecast Alpha & Beta",
    layout="wide"
)

# =========================
# HEADER + FAKE CAROUSEL
# =========================
st.markdown(
    """
<style>
/* ================================
   GLOBAL
   ================================ */
html,
body {
    scroll-behavior: smooth;
}

/* ================================
   HEADER (ALL DEVICES)
   ================================ */
#custom-header {
    position: fixed;              /* HEADER IKUT SCROLL */
    top: 0;
    left: 0;
    width: 100%;
    z-index: 9999;

    background-color: #ED5A0E;
    border-bottom: 1px solid rgba(0, 0, 0, 0.08);
    padding: 18px 42px;

    box-sizing: border-box;       /* ⬅️ PENTING */
}

/* ================================
   NAV CONTAINER
   ================================ */
.nav {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 16px;
}

/* ================================
   TITLE
   ================================ */
.nav-title {
    font-size: 24px;
    font-weight: 900;
    background: linear-gradient(
    90deg,
    #f5f5f0,
    #ffffff,
    #f5f5f0
);

    background-size: 200%;
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    animation: titleGlow 6s linear infinite;
}

@keyframes titleGlow {
    from { background-position: 0%; }
    to   { background-position: 200%; }
}

/* ================================
   MENU BUTTON
   ================================ */
.nav-menu {
    display: flex;
    gap: 12px;
    flex-wrap: nowrap;
}

.nav-menu a {
    padding: 10px 22px;
    border-radius: 999px;
    font-weight: 700;
    color: white !important;
    text-decoration: none;
    background: linear-gradient(
        90deg,
        #1f77b4,
        #3ba3ec,
        #1f77b4
    );
    white-space: nowrap;
}

/* ================================
   FILE UPLOADER
   ================================ */
div[data-testid="stFileUploader"] {
    display: flex;
    justify-content: center;
}

div[data-testid="stFileUploader"] > div {
    width: 500px;
}

/* ================================
   TABLET (769px – 1050px)
   ================================ */
@media (min-width: 769px) and (max-width: 1050px) {

    #custom-header {
        padding: 18px 28px;
    }

    .nav {
        flex-direction: column;     /* ⬅️ INI KUNCI */
        align-items: center;
        text-align: center;
    }

    .nav-title {
        font-size: 22px;
    }

    .nav-menu {
        flex-wrap: wrap;
        justify-content: center;
    }

    .nav-menu a {
        font-size: 14px;
        padding: 9px 18px;
    }

    div[data-testid="stFileUploader"] > div {
        max-width: 420px;
        width: 100%;
    }
}

/* ================================
   MOBILE (≤768px)
   ================================ */
@media (max-width: 768px) {

    #custom-header {
        padding: 14px 18px;

    }

    .nav {
        flex-direction: column;
        align-items: center;
        gap: 10px;
    }

    .nav-title {
        font-size: 20px;
        text-align: center;
    }

    .nav-menu {
        flex-wrap: wrap;
        justify-content: center;
        gap: 8px;
        margin-top: 20px;
    }

    .nav-menu a {
        font-size: 13px;
        padding: 8px 14px;
    }

    div[data-testid="stFileUploader"] > div {
        max-width: 320px;
        width: 100%;
    }

    h1 {
        font-size: 22px !important;
        text-align: center;
    }

    h2,
    h3 {
        font-size: 18px !important;
    }

    .stDataFrame,
    .stPlotlyChart {
        overflow-x: auto;
    }
}

/* ================================
   FIX CONTENT KETUTUP HEADER
   ================================ */
section.main {
    padding-top: 200px;   /* ⬅️ INI PALING PENTING */
}

/* ================================
   DESKTOP BESAR (≥1051px)
   FIX NAV MENU RUSAK
   ================================ */
@media (min-width: 1051px) {

    .nav {
        flex-wrap: wrap;          /* ⬅️ BIAR TIDAK KETIMPA */
    }

    .nav-menu {
        flex-wrap: wrap;          /* ⬅️ INI KUNCINYA */
        justify-content: flex-end;
        margin-top: 70px;
    }

    .nav-title {
        
        margin-top: 70px;
    }
}


</style>




<div id="custom-header">
    <div class="nav">
        <div class="nav-title">📈 Forecast Alpha & Beta</div>
        <div class="nav-menu">
            <a href="#input">Input Data</a>
            <a href="#forecast">Forecast</a>
            <a href="#error">Evaluasi Error</a>
            <a href="#export">Export</a>
        </div>
    </div>
</div>

""",
    unsafe_allow_html=True
)

# SPACER WAJIB
st.markdown('<div style="height:90px"></div>', unsafe_allow_html=True)
# =========================
# FILE UPLOADER (NATIVE)
# =========================
uploaded_file = st.file_uploader(
    "",
    type=["xlsx"],
    label_visibility="collapsed"
)

st.markdown('<div id="input"></div>', unsafe_allow_html=True)
st.title("📈 Forecast dengan Alpha & Beta (0.1 – 0.9)")


# =========================
# IMPORT EXCEL (AUTO HEADER DETECT)
# =========================
st.subheader("📥 Import Data Excel (Opsional)")

df_excel = None

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)

        header_row = None
        for i in range(min(10, len(df_raw))):
            row_text = " ".join(df_raw.iloc[i].astype(str).str.lower())
            if "bulan" in row_text and ("harga" in row_text or "close" in row_text):
                header_row = i
                break

        if header_row is None:
            st.error("❌ Header kolom tidak ditemukan")
            st.dataframe(df_raw.head(10))
        else:
            df_excel = pd.read_excel(uploaded_file, header=header_row)
            df_excel.columns = df_excel.columns.astype(str).str.strip().str.lower()

            column_map = {
                "bulan": "Bulan",
                "month": "Bulan",
                "tanggal": "Bulan",
                "periode": "Bulan",
                "harga penutupan": "Harga Penutupan",
                "harga": "Harga Penutupan",
                "close": "Harga Penutupan",
                "closing price": "Harga Penutupan",
            }

            df_excel = df_excel.rename(columns=column_map)

            if "Bulan" not in df_excel.columns or "Harga Penutupan" not in df_excel.columns:
                st.error("❌ Kolom tidak sesuai format")
                st.write(list(df_excel.columns))
            else:
                df_excel = df_excel[["Bulan", "Harga Penutupan"]]
                st.success("✅ Data Excel berhasil dibaca")
                st.dataframe(df_excel, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Gagal membaca file: {e}")

# =========================
# DATAFRAME UTAMA
# =========================
df_input = pd.DataFrame({
    "Bulan": ["" for _ in range(200)],
    "Harga Penutupan": [np.nan] * 200,
    "Level": [np.nan] * 200,
    "Trend": [np.nan] * 200,
    "Forecast": [np.nan] * 200,
    "Error": [np.nan] * 200,
    "Absolute Error": [np.nan] * 200,
    "Squared Error": [np.nan] * 200,
    "MAPE (%)": [np.nan] * 200,
})

# =========================
# ISI DATA DARI EXCEL
# =========================
if df_excel is not None:
    max_len = min(len(df_excel), len(df_input))
    df_input.loc[: max_len - 1, "Bulan"] = df_excel.loc[: max_len - 1, "Bulan"]
    df_input.loc[: max_len - 1, "Harga Penutupan"] = df_excel.loc[: max_len - 1, "Harga Penutupan"]


# =========================
# TABEL ATAS (INPUT)
# =========================
df_input_view = df_input[["Bulan","Harga Penutupan"]]

df_input_view = st.data_editor(
    df_input_view,
    num_rows="dynamic",
    use_container_width=True,
    key="editor_input"
)

df_input["Bulan"] = df_input_view["Bulan"]
df_input["Harga Penutupan"] = df_input_view["Harga Penutupan"]

# =========================
# PERHITUNGAN HOLT
# =========================
df_input.loc[0,"Level"] = df_input.loc[0,"Harga Penutupan"]
df_input.loc[0,"Trend"] = df_input.loc[1,"Harga Penutupan"] - df_input.loc[0,"Harga Penutupan"]

alpha_values = [round(i,1) for i in np.arange(0.1,1.0,0.1)]
beta_values  = [round(i,1) for i in np.arange(0.1,1.0,0.1)]
periode_opsi = list(range(1, 101))

c1,c2,c3 = st.columns(3)
with c1:
    selected_alpha = st.selectbox("Pilih Alpha", alpha_values)
with c2:
    selected_beta = st.selectbox("Pilih Beta", beta_values)
with c3:
    selected_period = st.selectbox("Periode Forecast (bulan ke depan)", periode_opsi)

for i in range(1,len(df_input)):
    if not np.isnan(df_input.loc[i,"Harga Penutupan"]):
        df_input.loc[i,"Level"] = (
            selected_alpha*df_input.loc[i,"Harga Penutupan"] +
            (1-selected_alpha)*(df_input.loc[i-1,"Level"]+df_input.loc[i-1,"Trend"])
        )
        df_input.loc[i,"Trend"] = (
            selected_beta*(df_input.loc[i,"Level"]-df_input.loc[i-1,"Level"]) +
            (1-selected_beta)*df_input.loc[i-1,"Trend"]
        )

last_actual = df_input["Harga Penutupan"].last_valid_index()

# =========================
# PARSER BULAN
# =========================
bulan_map = {
    "jan":1,"januari":1,"feb":2,"februari":2,"mar":3,"maret":3,
    "apr":4,"april":4,"mei":5,"may":5,"jun":6,"juni":6,
    "jul":7,"juli":7,"agt":8,"agu":8,"agustus":8,"aug":8,
    "sep":9,"september":9,"okt":10,"oct":10,"oktober":10,
    "nov":11,"november":11,"des":12,"dec":12,"desember":12
}

def parse_bulan(text):
    if not isinstance(text,str) or text.strip()=="":
        return None
    t = re.sub(r"[^a-z0-9 ]"," ",text.lower())
    parts = t.split()
    bulan = tahun = None
    for p in parts:
        if p in bulan_map: bulan = bulan_map[p]
        elif p.isdigit() and len(p)==4: tahun = int(p)
        elif p.isdigit() and len(p)==2: tahun = 2000 + int(p)
    if bulan and tahun:
        return datetime(tahun, bulan, 1)
    return None

# =========================
# AUTO BULAN FORECAST
# =========================
if last_actual is not None:
    last_text = df_input.loc[last_actual,"Bulan"]
    last_date = parse_bulan(last_text)

    if last_date:
        for h in range(1, selected_period+1):
            idx = last_actual + h
            if idx < len(df_input):
                next_date = last_date + relativedelta(months=h)
                df_input.loc[idx,"Bulan"] = next_date.strftime("%b-%y")

# =========================
# FORECAST
# =========================
if last_actual is not None:
    for i in range(1,last_actual+2):
        df_input.loc[i,"Forecast"] = df_input.loc[i-1,"Level"]+df_input.loc[i-1,"Trend"]

    L_last = df_input.loc[last_actual,"Level"]
    T_last = df_input.loc[last_actual,"Trend"]

    h = 1
    for i in range(last_actual+1, min(last_actual+1+selected_period, len(df_input))):
        df_input.loc[i,"Forecast"] = L_last + h*T_last
        h += 1

# =========================
# ERROR & METRIK
# =========================
for i in range(len(df_input)-1):
    df_input.loc[i,"Error"] = (
        df_input.loc[i,"Harga Penutupan"] -
        df_input.loc[i+1,"Forecast"]
    )

df_input["Error"] = df_input["Error"].bfill()
df_input["Absolute Error"] = df_input["Error"].abs()
df_input["Squared Error"] = df_input["Error"] ** 2
df_input["MAPE (%)"] = (df_input["Absolute Error"]/df_input["Harga Penutupan"])*100
df_input["MAPE (%)"] = df_input["MAPE (%)"].bfill()

# =========================
# TABEL BAWAH
# =========================
st.data_editor(
    df_input,
    num_rows="dynamic",
    use_container_width=True,
    key="editor_output"
)

st.markdown('<div id="forecast"></div>', unsafe_allow_html=True)
st.subheader("📊 Grafik Forecast")

# kalikan data ke ribuan SEBELUM melt
df_input["Harga Penutupan"] = df_input["Harga Penutupan"] * 1000
df_input["Forecast"] = df_input["Forecast"] * 1000

df_plot = df_input.melt(
    id_vars="Bulan",
    value_vars=["Harga Penutupan","Forecast"],
    var_name="Jenis",
    value_name="Harga"
)

fig = px.line(
    df_plot,
    x="Bulan",
    y="Harga",
    color="Jenis",
    markers=True,
    color_discrete_map={
        "Harga Penutupan": "blue",
        "Forecast": "orange"
    }
)

fig.update_layout(xaxis_title="Tahun")

fig.update_yaxes(
    title="Harga (USD)",
    tickformat=",.0f",
    showexponent="none",
    exponentformat="none"
)

st.plotly_chart(fig, use_container_width=True)

# =========================
# RINGKASAN ERROR
# =========================
st.markdown('<div id="error"></div>', unsafe_allow_html=True)
st.subheader("📊 Tabel Ringkasan Error")


mean_error = df_input["Error"].mean()
mae = df_input["Absolute Error"].mean()
mse = df_input["Squared Error"].mean()
rmse = np.sqrt(mse)
mape_total = df_input["MAPE (%)"].mean()

df_summary = pd.DataFrame({
    "Metode": ["Mean Error", "MAE", "MSE", "RMSE", "MAPE"],
    "Nilai": [mean_error, mae, mse, rmse, mape_total]
})

st.dataframe(df_summary, use_container_width=True)

# =========================
# EXPORT
# =========================
st.markdown('<div id="export"></div>', unsafe_allow_html=True)
st.subheader("⬇️ Export Data & Grafik")

img_bytes = fig.to_image(format="png")

# =========================
# EXPORT EXCEL
# =========================
excel_buffer = BytesIO()
wb = Workbook()
ws = wb.active
ws.title = "Hasil Forecast"

ws.append(["Alpha", selected_alpha])
ws.append(["Beta", selected_beta])
ws.append(["Periode Forecast (bulan)", selected_period])
ws.append([])

ws.append(["No"] + list(df_input.columns))
for i, row in df_input.iterrows():
    ws.append([i + 1] + list(row))

img = XLImage(BytesIO(img_bytes))
img.anchor = "L2"
ws.add_image(img)

wb.save(excel_buffer)
excel_buffer.seek(0)

# ⬇️ WRAPPER EXCEL
st.markdown('<div class="export-excel">', unsafe_allow_html=True)
st.download_button(
    "📥 Download Excel",
    excel_buffer,
    "forecast_holt.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.markdown('</div>', unsafe_allow_html=True)


# =========================
# EXPORT PDF
# =========================
pdf_buffer = BytesIO()
doc = SimpleDocTemplate(pdf_buffer, pagesize=A4)

table_data = [
    ["Alpha", selected_alpha],
    ["Beta", selected_beta],
    ["Periode Forecast (bulan)", selected_period],
    []
]
table_data.append(["No"] + list(df_input.columns))

for i, row in df_input.iterrows():
    table_data.append([i + 1] + list(row))

table = Table(table_data, repeatRows=1)
table.setStyle(TableStyle([
    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey)
]))

doc.build([
    table,
    Image(BytesIO(img_bytes), width=400, height=250)
])

pdf_buffer.seek(0)

# ⬇️ WRAPPER PDF
st.markdown('<div class="export-pdf">', unsafe_allow_html=True)
st.download_button(
    "📄 Download PDF",
    pdf_buffer,
    "forecast_holt.pdf",
    "application/pdf"
)
st.markdown('</div>', unsafe_allow_html=True)

