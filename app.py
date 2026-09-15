import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import re
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib.units import cm


# =========================================================
# KONFIGURASI
# =========================================================

st.set_page_config(
    page_title="Dashboard Produktivitas 2026",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)




# =========================================================
# NAMA SHEET
# =========================================================

TRAIN_SHEET = "Pelatihan Produktivitas"
BIM_SHEET = "Bimbingan Konsultasi"

MONTHS = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember"
]


# =========================================================
# BACA EXCEL
# =========================================================

@st.cache_data(show_spinner=False)
def load_excel(file_bytes):

    xls = pd.ExcelFile(
        BytesIO(file_bytes),
        engine="openpyxl"
    )

    sheets = {}

    for sheet in xls.sheet_names:

        sheets[sheet] = pd.read_excel(
            xls,
            sheet_name=sheet,
            header=None
        )

    return sheets


# =========================================================
# KOLOM UNIK
# =========================================================

def make_unique(cols):

    seen = {}
    output = []

    for col in cols:

        col = (
            str(col).strip()
            if pd.notna(col)
            else "Kolom"
        )

        seen[col] = (
            seen.get(col, 0) + 1
        )

        if seen[col] == 1:

            output.append(col)

        else:

            output.append(
                f"{col}_{seen[col]}"
            )

    return output


# =========================================================
# NORMALISASI TEKS
# =========================================================

def normalisasi_teks(value):

    if pd.isna(value):
        return pd.NA

    text = str(value).strip()

    if text == "":
        return pd.NA

    return text


# =========================================================
# STANDARDISASI PROVINSI
# =========================================================

def standardisasi_provinsi(value):

    if pd.isna(value):
        return "Tidak Diisi"

    text = str(value).strip().upper()

    mapping = {

        "DKI JAKARTA": "DKI Jakarta",

        "JAWA BARAT": "Jawa Barat",
        "JAWA TENGAH": "Jawa Tengah",
        "JAWA TIMUR": "Jawa Timur",

        "BANTEN": "Banten",

        "SUMATERA BARAT": "Sumatera Barat",
        "SUMATERA UTARA": "Sumatera Utara",
        "SUMATERA SELATAN": "Sumatera Selatan",

        "KALIMANTAN TIMUR": "Kalimantan Timur",
        "KALIMANTAN BARAT": "Kalimantan Barat",
        "KALIMANTAN SELATAN": "Kalimantan Selatan",
        "KALIMANTAN TENGAH": "Kalimantan Tengah",
        "KALIMANTAN UTARA": "Kalimantan Utara",

        "ACEH": "Aceh",

        "SULAWESI SELATAN": "Sulawesi Selatan",
        "SULSEL": "Sulawesi Selatan",

        "SULAWESI TENGGARA": "Sulawesi Tenggara",
        "SULAWESI UTARA": "Sulawesi Utara",
        "SULAWESI TENGAH": "Sulawesi Tengah",

        "PAPUA BARAT DAYA": "Papua Barat Daya",
        "PAPUA": "Papua"
    }

    return mapping.get(
        text,
        str(value).strip().title()
    )


# =========================================================
# STANDARDISASI METODE
# =========================================================

def standardisasi_metode(value):

    if pd.isna(value):
        return "Tidak Diisi"

    text = str(value).strip().lower()

    if text == "online":
        return "Online"

    if text == "offline":
        return "Offline"

    if text == "hybrid":
        return "Hybrid"

    return str(value).strip().title()


# =========================================================
# JENIS PELATIHAN
# =========================================================

def standardisasi_jenis_pelatihan(value):

    if pd.isna(value):
        return "Tidak Diisi"

    text = str(value).strip().upper()

    if text == "NON BOARDING":
        return "Non Boarding"

    if text == "BOARDING":
        return "Boarding"

    return str(value).strip().title()


# =========================================================
# STATUS LULUS
# =========================================================

def standardisasi_status_lulus(value):

    if pd.isna(value):
        return "Tidak Diisi"

    text = str(value).strip().lower()

    if text in [
        "lulus",
        "100"
    ]:
        return "Lulus"

    if text in [
        "tidak lulus",
        "0"
    ]:
        return "Tidak Lulus"

    return str(value).strip().title()


# =========================================================
# STATUS SELESAI
# =========================================================

def standardisasi_status_selesai(value):

    if pd.isna(value):
        return "Tidak Diisi"

    text = str(value).strip().lower()

    if text in [
        "selesai",
        "ya",
        "1"
    ]:
        return "Selesai"

    if text in [
        "belum selesai",
        "tidak",
        "0"
    ]:
        return "Belum Selesai"

    return str(value).strip().title()


# =========================================================
# KEJURUAN
# =========================================================

def standardisasi_kejuruan(value):

    if pd.isna(value):
        return "Tidak Diisi"

    text = str(value).strip().lower()

    if text == "produktivitas":
        return "Produktivitas"

    return str(value).strip().title()


# =========================================================
# KATEGORI BIDANG USAHA
# =========================================================

def kategori_bidang_usaha(value):

    if pd.isna(value):
        return "Tidak Diisi"

    text = str(value).strip().lower()

    if text == "":
        return "Tidak Diisi"


    # MAKANAN & MINUMAN

    if any(
        k in text
        for k in [

            "kuliner",
            "bakery",
            "keripik",
            "snack",
            "dimsum",
            "kue",
            "catering",
            "bawang goreng",
            "jus",
            "bubur",
            "jamu",
            "peyek",
            "tape",
            "bandeng presto",
            "sambal",
            "seblak",
            "warung makan",
            "minuman",
            "ayam gepuk",
            "donat",
            "abon",
            "crispy",
            "tepung",
            "makanan"
        ]
    ):

        return "Makanan & Minuman"


    # TEKSTIL

    if any(
        k in text
        for k in [

            "menjahit",
            "jahit",
            "fashion",
            "konveksi",
            "pakaian",
            "tekstil",
            "bordir"
        ]
    ):

        return "Tekstil & Fashion"


    # OTOMOTIF

    if any(
        k in text
        for k in [

            "bengkel",
            "otomotif",
            "motor",
            "mobil",
            "kendaraan"
        ]
    ):

        return "Otomotif"


    # PERTANIAN & PERIKANAN

    if any(
        k in text
        for k in [

            "budidaya",
            "pertanian",
            "peternakan",
            "perikanan",
            "ikan",
            "perkebunan"
        ]
    ):

        return "Pertanian & Perikanan"


    # PERDAGANGAN

    if any(
        k in text
        for k in [

            "dagang",
            "toko",
            "agen",
            "perdagangan",
            "distributor"
        ]
    ):

        return "Perdagangan"


    # JASA

    if any(
        k in text
        for k in [

            "jasa",
            "expedisi",
            "ekspedisi",
            "service"
        ]
    ):

        return "Jasa"


    return "Lainnya"


# =========================================================
# CLEANING PELATIHAN
# =========================================================

def clean_training(raw):

    df = raw.copy()

    df.columns = make_unique(
        df.iloc[0]
    )

    df = df.iloc[1:].copy()

    df = df.dropna(
        how="all"
    )


    # Hanya peserta

    if "Nama Peserta" in df.columns:

        df = df[
            df["Nama Peserta"].notna()
        ].copy()


    # Tanggal

    for col in [

        "Tanggal Mulai Pelatihan",
        "Tanggal Selesai Pelatihan",
        "Tanggal Lahir"

    ]:

        if col in df.columns:

            df[col] = pd.to_datetime(
                df[col],
                errors="coerce",
                dayfirst=True
            )


    # Numeric

    for col in [

        "Durasi (JP)",
        "Absensi"

    ]:

        if col in df.columns:

            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            )


    # Bulan & Tahun

    if "Tanggal Mulai Pelatihan" in df.columns:

        bulan = {

            1: "Januari",
            2: "Februari",
            3: "Maret",
            4: "April",
            5: "Mei",
            6: "Juni",
            7: "Juli",
            8: "Agustus",
            9: "September",
            10: "Oktober",
            11: "November",
            12: "Desember"
        }

        df["Bulan Pelatihan"] = (
            df["Tanggal Mulai Pelatihan"]
            .dt.month
            .map(bulan)
        )

        df["Tahun Pelatihan"] = (
            df["Tanggal Mulai Pelatihan"]
            .dt.year
        )


    # Standardisasi

    if "Provinsi" in df.columns:

        df["Provinsi"] = (
            df["Provinsi"]
            .apply(standardisasi_provinsi)
        )


    if "Metode Pelatihan" in df.columns:

        df["Metode Pelatihan"] = (
            df["Metode Pelatihan"]
            .apply(standardisasi_metode)
        )


    if "Jenis Pelatihan" in df.columns:

        df["Jenis Pelatihan"] = (
            df["Jenis Pelatihan"]
            .apply(standardisasi_jenis_pelatihan)
        )


    if "Status KeLulusan" in df.columns:

        df["Status KeLulusan"] = (
            df["Status KeLulusan"]
            .apply(standardisasi_status_lulus)
        )


    if "Status Selesai Pelatihan" in df.columns:

        df["Status Selesai Pelatihan"] = (
            df["Status Selesai Pelatihan"]
            .apply(standardisasi_status_selesai)
        )


    if "Kejuruan" in df.columns:

        df["Kejuruan"] = (
            df["Kejuruan"]
            .apply(standardisasi_kejuruan)
        )


    # Normalisasi teks

    for col in [

        "Kab./Kota",
        "Nama Lembaga",
        "Judul Program Pelatihan",
        "Jenis Kelamin",
        "Pendidikan Terakhir (setara dengan)"

    ]:

        if col in df.columns:

            df[col] = (
                df[col]
                .apply(normalisasi_teks)
            )


    return df.reset_index(
        drop=True
    )


# =========================================================
# CLEANING BIMBINGAN
# =========================================================

def clean_bim(raw):

    df = raw.copy()

    header = 0


    # Cari header

    for i, row in df.iterrows():

        values = (
            row.astype(str)
            .str.upper()
            .tolist()
        )

        if "NAMA PERUSAHAAN" in values:

            header = i

            break


    df.columns = make_unique(
        df.iloc[header]
    )

    df = df.iloc[
        header + 1:
    ].copy()

    df = df.dropna(
        how="all"
    )


    # Hanya perusahaan

    if "NAMA PERUSAHAAN" in df.columns:

        df = df[
            df["NAMA PERUSAHAAN"].notna()
        ].copy()

        df = df[
            df["NAMA PERUSAHAAN"]
            .astype(str)
            .str.upper()
            != "CONTOH PENGISIAN"
        ]


    # Tanggal pelaksanaan

    if "WAKTU PELAKSANAAN" in df.columns:

        df["Tanggal Pelaksanaan"] = pd.to_datetime(
            df["WAKTU PELAKSANAAN"],
            errors="coerce",
            dayfirst=True
        )


        # Contoh:
        # 20/7/2026 s/d 12/08/2026

        mask = (
            df["Tanggal Pelaksanaan"]
            .isna()
        )

        if mask.any():

            extracted = (
                df.loc[
                    mask,
                    "WAKTU PELAKSANAAN"
                ]
                .astype(str)
                .str.extract(
                    r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})"
                )[0]
            )

            df.loc[
                mask,
                "Tanggal Pelaksanaan"
            ] = pd.to_datetime(
                extracted,
                errors="coerce",
                dayfirst=True
            )


    # Tanggal data

    if (
        "DATA PER TANGGAL (tanggal/bulan/tahun)"
        in df.columns
    ):

        df["Tanggal Data"] = pd.to_datetime(
            df[
                "DATA PER TANGGAL (tanggal/bulan/tahun)"
            ],
            errors="coerce",
            dayfirst=True
        )


    # Bulan & tahun

    if "Tanggal Pelaksanaan" in df.columns:

        bulan = {

            1: "Januari",
            2: "Februari",
            3: "Maret",
            4: "April",
            5: "Mei",
            6: "Juni",
            7: "Juli",
            8: "Agustus",
            9: "September",
            10: "Oktober",
            11: "November",
            12: "Desember"
        }

        df["Bulan Bimbingan"] = (
            df["Tanggal Pelaksanaan"]
            .dt.month
            .map(bulan)
        )

        df["Tahun Bimbingan"] = (
            df["Tanggal Pelaksanaan"]
            .dt.year
        )


    # Normalisasi teks

    for col in [

        "NAMA KABUPATEN/KOTA",
        "BIDANG USAHA",
        "PELAKSANA",
        "AKTIVITAS PENINGKATAN PRODUKTIVITAS",
        "HASIL PENINGKATAN PRODUKTIVITAS"

    ]:

        if col in df.columns:

            df[col] = (
                df[col]
                .apply(normalisasi_teks)
            )


    # Bidang usaha

    if "BIDANG USAHA" in df.columns:

        df["BIDANG USAHA ASLI"] = (
            df["BIDANG USAHA"]
        )

        df["Bidang Usaha Kategori"] = (
            df["BIDANG USAHA"]
            .apply(kategori_bidang_usaha)
        )


    return df.reset_index(
        drop=True
    )


# =========================================================
# FILTER
# =========================================================

def selectbox_filter(
    df,
    col,
    label,
    key=None
):

    if col not in df.columns:
        return df

    values = sorted([
        str(x)
        for x in df[col]
        .dropna()
        .unique()
    ])

    selected = st.sidebar.selectbox(
        label,
        ["Semua"] + values,
        key=key
    )

    if selected == "Semua":
        return df

    return df[
        df[col].astype(str)
        == selected
    ]


# =========================================================
# FORMAT ANGKA
# =========================================================

def format_number(value):

    try:

        return f"{int(value):,}".replace(
            ",",
            "."
        )

    except:

        return "0"


# =========================================================
# PERSENTASE
# =========================================================

def percentage(part, total):

    if total == 0:
        return 0

    return (
        part / total
    ) * 100


# =========================================================
# DATA BULANAN
# =========================================================

def monthly_data(
    df,
    month_column,
    value_name
):

    if month_column not in df.columns:

        return pd.DataFrame({
            "Bulan": MONTHS,
            value_name: [0] * 12
        })


    x = (
        df[month_column]
        .value_counts()
        .reindex(MONTHS)
        .fillna(0)
        .reset_index()
    )

    x.columns = [
        "Bulan",
        value_name
    ]

    return x


# =========================================================
# STYLE PLOTLY
# =========================================================

def style_chart(
    fig,
    height=370
):

    fig.update_layout(

        height=height,

        margin=dict(
            l=25,
            r=25,
            t=60,
            b=35
        ),

        paper_bgcolor="white",

        plot_bgcolor="white",

        font=dict(
            family="Arial",
            color="#17324d"
        ),

        title_font=dict(
            size=18,
            color="#104d7b"
        ),

        legend=dict(
            orientation="h",
            y=1.08,
            x=0
        )
    )


    fig.update_xaxes(
        showgrid=True,
        gridcolor="#e6eef5",
        zeroline=False
    )

    fig.update_yaxes(
        showgrid=True,
        gridcolor="#e6eef5",
        zeroline=False
    )

    return fig


# =========================================================
# INSIGHT PELATIHAN
# =========================================================

def insights_training(df):

    insights = []

    if df.empty:

        return [
            "Tidak ada data sesuai filter."
        ]


    total = len(df)


    # Provinsi

    if "Provinsi" in df.columns:

        x = (
            df["Provinsi"]
            .fillna("Tidak Diisi")
            .value_counts()
        )

        if not x.empty:

            nama = x.index[0]
            jumlah = int(x.iloc[0])
            persen = percentage(
                jumlah,
                total
            )

            insights.append(
                (
                    "🏆 **Provinsi Dominan** — "
                    f"{nama} memiliki peserta terbanyak, "
                    f"yaitu **{format_number(jumlah)} peserta "
                    f"({persen:.1f}%)**."
                )
            )


    # Bulan

    if "Bulan Pelatihan" in df.columns:

        x = (
            df["Bulan Pelatihan"]
            .value_counts()
            .reindex(MONTHS)
            .fillna(0)
        )

        if x.max() > 0:

            bulan = x.idxmax()
            jumlah = int(x.max())

            insights.append(
                (
                    "📈 **Bulan Tertinggi** — "
                    f"Jumlah peserta tertinggi terjadi pada "
                    f"**{bulan}**, sebanyak "
                    f"**{format_number(jumlah)} peserta**."
                )
            )


    # Kelulusan

    if "Status KeLulusan" in df.columns:

        lulus = (
            df["Status KeLulusan"]
            .astype(str)
            .str.lower()
            == "lulus"
        ).sum()

        persen = percentage(
            lulus,
            total
        )

        insights.append(
            (
                "🎓 **Kelulusan** — "
                f"Sebanyak **{format_number(lulus)} peserta** "
                f"berstatus lulus dengan tingkat kelulusan "
                f"**{persen:.1f}%**."
            )
        )


    # Metode

    if "Metode Pelatihan" in df.columns:

        x = (
            df["Metode Pelatihan"]
            .fillna("Tidak Diisi")
            .value_counts()
        )

        if not x.empty:

            metode = x.index[0]
            jumlah = int(x.iloc[0])
            persen = percentage(
                jumlah,
                total
            )

            insights.append(
                (
                    "⚙️ **Metode Dominan** — "
                    f"Metode **{metode}** digunakan oleh "
                    f"**{format_number(jumlah)} peserta "
                    f"({persen:.1f}%)**."
                )
            )


    return insights


# =========================================================
# INSIGHT BIMBINGAN
# =========================================================

def insights_bim(df):

    insights = []

    if df.empty:

        return [
            "Tidak ada data sesuai filter."
        ]


    total = len(df)


    # Bidang usaha

    if "Bidang Usaha Kategori" in df.columns:

        x = (
            df["Bidang Usaha Kategori"]
            .fillna("Tidak Diisi")
            .value_counts()
        )

        if not x.empty:

            bidang = x.index[0]
            jumlah = int(x.iloc[0])
            persen = percentage(
                jumlah,
                total
            )

            insights.append(
                (
                    "🏭 **Bidang Usaha Dominan** — "
                    f"{bidang} menjadi kategori dengan "
                    f"bimbingan terbanyak, yaitu "
                    f"**{format_number(jumlah)} kegiatan "
                    f"({persen:.1f}%)**."
                )
            )


    # Wilayah

    if "NAMA KABUPATEN/KOTA" in df.columns:

        x = (
            df["NAMA KABUPATEN/KOTA"]
            .fillna("Tidak Diisi")
            .value_counts()
        )

        if not x.empty:

            wilayah = x.index[0]
            jumlah = int(x.iloc[0])

            insights.append(
                (
                    "📍 **Wilayah Dominan** — "
                    f"{wilayah} menjadi wilayah dengan "
                    f"kegiatan bimbingan terbanyak, "
                    f"yaitu **{format_number(jumlah)} kegiatan**."
                )
            )


    # Bulan

    if "Bulan Bimbingan" in df.columns:

        x = (
            df["Bulan Bimbingan"]
            .value_counts()
            .reindex(MONTHS)
            .fillna(0)
        )

        if x.max() > 0:

            bulan = x.idxmax()
            jumlah = int(x.max())

            insights.append(
                (
                    "📅 **Bulan Tertinggi** — "
                    f"Aktivitas bimbingan tertinggi terjadi "
                    f"pada **{bulan}**, sebanyak "
                    f"**{format_number(jumlah)} kegiatan**."
                )
            )


    return insights


# =========================================================
# EXPORT PDF HASIL DASHBOARD
# =========================================================

def _pdf_text(text):
    """Bersihkan markdown sederhana dari insight sebelum dimasukkan ke PDF."""
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", str(text))
    return text.replace("—", "-").replace("📌", "").strip()


def _save_chart_png(fig):
    """Simpan matplotlib figure ke memory agar bisa dimasukkan ke PDF."""
    buf = BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _matplotlib_training_charts(df):
    charts = []
    # Tren peserta
    x = monthly_data(df, "Bulan Pelatihan", "Peserta")
    if not x.empty:
        fig, ax = plt.subplots(figsize=(8.0, 3.8))
        ax.plot(x["Bulan"], x["Peserta"], marker="o")
        ax.set_title("Peserta per Bulan")
        ax.set_xlabel("Bulan")
        ax.set_ylabel("Peserta")
        ax.tick_params(axis="x", rotation=45)
        charts.append(_save_chart_png(fig))

    # Top provinsi
    if "Provinsi" in df.columns:
        x = df["Provinsi"].fillna("Tidak Diisi").value_counts().head(10).sort_values()
        if not x.empty:
            fig, ax = plt.subplots(figsize=(8.0, 4.2))
            ax.barh(x.index.astype(str), x.values)
            ax.set_title("Top 10 Provinsi")
            ax.set_xlabel("Peserta")
            charts.append(_save_chart_png(fig))

    # Kelulusan
    if "Status KeLulusan" in df.columns:
        x = df["Status KeLulusan"].fillna("Tidak Diisi").value_counts()
        if not x.empty:
            fig, ax = plt.subplots(figsize=(6.5, 4.0))
            ax.pie(x.values, labels=x.index.astype(str), autopct="%1.1f%%")
            ax.set_title("Status Kelulusan")
            charts.append(_save_chart_png(fig))

    # Metode
    if "Metode Pelatihan" in df.columns:
        x = df["Metode Pelatihan"].fillna("Tidak Diisi").value_counts()
        if not x.empty:
            fig, ax = plt.subplots(figsize=(7.0, 4.0))
            ax.bar(x.index.astype(str), x.values)
            ax.set_title("Metode Pelatihan")
            ax.set_ylabel("Peserta")
            ax.tick_params(axis="x", rotation=25)
            charts.append(_save_chart_png(fig))
    return charts


def _matplotlib_bim_charts(df):
    charts = []
    # Tren bimbingan
    x = monthly_data(df, "Bulan Bimbingan", "Kegiatan")
    if not x.empty:
        fig, ax = plt.subplots(figsize=(8.0, 3.8))
        ax.plot(x["Bulan"], x["Kegiatan"], marker="o")
        ax.set_title("Tren Bimbingan Konsultasi")
        ax.set_xlabel("Bulan")
        ax.set_ylabel("Kegiatan")
        ax.tick_params(axis="x", rotation=45)
        charts.append(_save_chart_png(fig))

    # Bidang usaha
    if "Bidang Usaha Kategori" in df.columns:
        x = df["Bidang Usaha Kategori"].fillna("Tidak Diisi").value_counts().sort_values()
        if not x.empty:
            fig, ax = plt.subplots(figsize=(8.0, 4.2))
            ax.barh(x.index.astype(str), x.values)
            ax.set_title("Kegiatan Berdasarkan Bidang Usaha")
            ax.set_xlabel("Kegiatan")
            charts.append(_save_chart_png(fig))

    # Wilayah
    for col in ["NAMA KABUPATEN/KOTA", "Kabupaten/Kota", "NAMA KABUPATEN / KOTA"]:
        if col in df.columns:
            x = df[col].fillna("Tidak Diisi").value_counts().head(10).sort_values()
            if not x.empty:
                fig, ax = plt.subplots(figsize=(8.0, 4.2))
                ax.barh(x.index.astype(str), x.values)
                ax.set_title("Top 10 Wilayah")
                ax.set_xlabel("Kegiatan")
                charts.append(_save_chart_png(fig))
            break
    return charts


def create_dashboard_pdf(training_df=None, bim_df=None, title="Hasil Analisis Dashboard Produktivitas 2026"):
    """Membuat PDF visual dari hasil dashboard sesuai filter aktif."""
    output = BytesIO()
    doc = SimpleDocTemplate(
        output, pagesize=A4,
        rightMargin=1.4*cm, leftMargin=1.4*cm,
        topMargin=1.3*cm, bottomMargin=1.3*cm
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="PdfTitle", parent=styles["Title"], alignment=TA_CENTER, fontSize=17, leading=21, spaceAfter=8))
    styles.add(ParagraphStyle(name="PdfSub", parent=styles["Normal"], alignment=TA_CENTER, fontSize=9, textColor=colors.grey, spaceAfter=12))
    styles.add(ParagraphStyle(name="PdfHead", parent=styles["Heading2"], fontSize=13, leading=16, spaceBefore=8, spaceAfter=7))
    styles.add(ParagraphStyle(name="PdfBody", parent=styles["BodyText"], fontSize=9, leading=13, spaceAfter=5))
    styles.add(ParagraphStyle(name="PdfSmall", parent=styles["BodyText"], fontSize=8, leading=11, spaceAfter=4))

    story = [Paragraph(title, styles["PdfTitle"]),
             Paragraph("Laporan visual berdasarkan data dan filter yang sedang aktif pada dashboard.", styles["PdfSub"])]

    def add_kpis(items):
        data = [[Paragraph(f"<b>{k}</b><br/>{v}", styles["PdfSmall"]) for k, v in items]]
        table = Table(data, colWidths=[(A4[0]-2.8*cm)/len(items)]*len(items))
        table.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.5, colors.lightgrey),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ALIGN", (0,0), (-1,-1), "CENTER"),
            ("TOPPADDING", (0,0), (-1,-1), 8),
            ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ]))
        story.extend([table, Spacer(1, 8)])

    if training_df is not None and not training_df.empty:
        df = training_df
        story.append(Paragraph("🎓 Pelatihan Produktivitas", styles["PdfHead"]))
        add_kpis([
            ("Total Peserta", format_number(len(df))),
            ("Total Program", format_number(df["Judul Program Pelatihan"].nunique() if "Judul Program Pelatihan" in df.columns else 0)),
            ("Total Provinsi", format_number(df["Provinsi"].nunique() if "Provinsi" in df.columns else 0)),
            ("Total Lembaga", format_number(df["Nama Lembaga"].nunique() if "Nama Lembaga" in df.columns else 0)),
        ])
        story.append(Paragraph("Analisis Grafik", styles["PdfHead"]))
        charts = _matplotlib_training_charts(df)
        for i, chart in enumerate(charts):
            story.append(Image(chart, width=16.5*cm, height=8.0*cm))
            if i < len(charts)-1: story.append(Spacer(1, 5))
        story.append(Paragraph("Insight Otomatis", styles["PdfHead"]))
        for text in insights_training(df):
            story.append(Paragraph("• " + _pdf_text(text), styles["PdfBody"]))
        story.append(Paragraph("Kesimpulan", styles["PdfHead"]))
        story.append(Paragraph("Analisis pelatihan menggambarkan jumlah peserta, program, pemerataan provinsi dan lembaga, serta pola kelulusan dan metode pelatihan berdasarkan filter yang dipilih.", styles["PdfBody"]))

    if bim_df is not None and not bim_df.empty:
        if training_df is not None and not training_df.empty:
            story.append(PageBreak())
        df = bim_df
        story.append(Paragraph("🤝 Bimbingan Konsultasi", styles["PdfHead"]))
        add_kpis([
            ("Total Kegiatan", format_number(len(df))),
            ("Total Perusahaan", format_number(df["NAMA PERUSAHAAN"].nunique() if "NAMA PERUSAHAAN" in df.columns else 0)),
            ("Total Wilayah", format_number(df["NAMA KABUPATEN/KOTA"].nunique() if "NAMA KABUPATEN/KOTA" in df.columns else 0)),
            ("Bidang Usaha", format_number(df["Bidang Usaha Kategori"].nunique() if "Bidang Usaha Kategori" in df.columns else 0)),
        ])
        story.append(Paragraph("Analisis Grafik", styles["PdfHead"]))
        charts = _matplotlib_bim_charts(df)
        for i, chart in enumerate(charts):
            story.append(Image(chart, width=16.5*cm, height=8.0*cm))
            if i < len(charts)-1: story.append(Spacer(1, 5))
        story.append(Paragraph("Insight Otomatis", styles["PdfHead"]))
        for text in insights_bim(df):
            story.append(Paragraph("• " + _pdf_text(text), styles["PdfBody"]))
        story.append(Paragraph("Kesimpulan", styles["PdfHead"]))
        story.append(Paragraph("Analisis bimbingan memberikan gambaran mengenai konsentrasi kegiatan berdasarkan bidang usaha, wilayah, dan waktu pelaksanaan untuk mendukung monitoring dan evaluasi pendampingan.", styles["PdfBody"]))

    if not story:
        story.append(Paragraph("Tidak ada data sesuai filter.", styles["PdfBody"]))
    doc.build(story)
    output.seek(0)
    return output.getvalue()


# =========================================================
# EXPORT HASIL ANALISIS
# =========================================================

def create_analysis_excel(
    training_df=None,
    bim_df=None
):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:


        # =================================================
        # PELATIHAN
        # =================================================

        if (
            training_df is not None
            and not training_df.empty
        ):

            # Data filter

            training_df.to_excel(
                writer,
                sheet_name="Data Pelatihan",
                index=False
            )


            # KPI

            total_peserta = len(
                training_df
            )

            total_program = (
                training_df[
                    "Judul Program Pelatihan"
                ].nunique()
                if "Judul Program Pelatihan"
                in training_df.columns
                else 0
            )

            total_provinsi = (
                training_df[
                    "Provinsi"
                ].nunique()
                if "Provinsi"
                in training_df.columns
                else 0
            )

            total_lembaga = (
                training_df[
                    "Nama Lembaga"
                ].nunique()
                if "Nama Lembaga"
                in training_df.columns
                else 0
            )


            kpi = pd.DataFrame({

                "Indikator": [

                    "Total Peserta",
                    "Total Program",
                    "Total Provinsi",
                    "Total Lembaga"

                ],

                "Nilai": [

                    total_peserta,
                    total_program,
                    total_provinsi,
                    total_lembaga

                ]

            })


            kpi.to_excel(
                writer,
                sheet_name="KPI Pelatihan",
                index=False
            )


            # Peserta per bulan

            bulanan = monthly_data(
                training_df,
                "Bulan Pelatihan",
                "Jumlah Peserta"
            )

            bulanan.to_excel(
                writer,
                sheet_name="Peserta per Bulan",
                index=False
            )


            # Provinsi

            if "Provinsi" in training_df.columns:

                provinsi = (
                    training_df["Provinsi"]
                    .fillna("Tidak Diisi")
                    .value_counts()
                    .reset_index()
                )

                provinsi.columns = [
                    "Provinsi",
                    "Jumlah Peserta"
                ]

                provinsi.to_excel(
                    writer,
                    sheet_name="Peserta per Provinsi",
                    index=False
                )


            # Kelulusan

            if "Status KeLulusan" in training_df.columns:

                status = (
                    training_df["Status KeLulusan"]
                    .fillna("Tidak Diisi")
                    .value_counts()
                    .reset_index()
                )

                status.columns = [
                    "Status Kelulusan",
                    "Jumlah"
                ]

                status.to_excel(
                    writer,
                    sheet_name="Status Kelulusan",
                    index=False
                )


            # Metode

            if "Metode Pelatihan" in training_df.columns:

                metode = (
                    training_df["Metode Pelatihan"]
                    .fillna("Tidak Diisi")
                    .value_counts()
                    .reset_index()
                )

                metode.columns = [
                    "Metode Pelatihan",
                    "Jumlah"
                ]

                metode.to_excel(
                    writer,
                    sheet_name="Metode Pelatihan",
                    index=False
                )


            # Insight

            insight_text = insights_training(
                training_df
            )

            insight_df = pd.DataFrame({

                "Insight": [
                    x.replace("**", "")
                    for x in insight_text
                ]

            })

            insight_df.to_excel(
                writer,
                sheet_name="Insight Analisis",
                index=False
            )


        # =================================================
        # BIMBINGAN
        # =================================================

        if (
            bim_df is not None
            and not bim_df.empty
        ):

            bim_df.to_excel(
                writer,
                sheet_name="Data Bimbingan",
                index=False
            )


            total_kegiatan = len(
                bim_df
            )

            total_perusahaan = (
                bim_df[
                    "NAMA PERUSAHAAN"
                ].nunique()
                if "NAMA PERUSAHAAN"
                in bim_df.columns
                else 0
            )

            total_wilayah = (
                bim_df[
                    "NAMA KABUPATEN/KOTA"
                ].nunique()
                if "NAMA KABUPATEN/KOTA"
                in bim_df.columns
                else 0
            )

            total_bidang = (
                bim_df[
                    "Bidang Usaha Kategori"
                ].nunique()
                if "Bidang Usaha Kategori"
                in bim_df.columns
                else 0
            )


            kpi = pd.DataFrame({

                "Indikator": [

                    "Total Kegiatan",
                    "Total Perusahaan",
                    "Total Wilayah",
                    "Kategori Bidang Usaha"

                ],

                "Nilai": [

                    total_kegiatan,
                    total_perusahaan,
                    total_wilayah,
                    total_bidang

                ]

            })


            kpi.to_excel(
                writer,
                sheet_name="KPI Bimbingan",
                index=False
            )


            # Bulanan

            bulanan = monthly_data(
                bim_df,
                "Bulan Bimbingan",
                "Jumlah Kegiatan"
            )

            bulanan.to_excel(
                writer,
                sheet_name="Bimbingan per Bulan",
                index=False
            )


            # Bidang usaha

            if "Bidang Usaha Kategori" in bim_df.columns:

                bidang = (
                    bim_df[
                        "Bidang Usaha Kategori"
                    ]
                    .fillna("Tidak Diisi")
                    .value_counts()
                    .reset_index()
                )

                bidang.columns = [
                    "Bidang Usaha",
                    "Jumlah Kegiatan"
                ]

                bidang.to_excel(
                    writer,
                    sheet_name="Bidang Usaha",
                    index=False
                )


            # Wilayah

            if "NAMA KABUPATEN/KOTA" in bim_df.columns:

                wilayah = (
                    bim_df[
                        "NAMA KABUPATEN/KOTA"
                    ]
                    .fillna("Tidak Diisi")
                    .value_counts()
                    .reset_index()
                )

                wilayah.columns = [
                    "Wilayah",
                    "Jumlah Kegiatan"
                ]

                wilayah.to_excel(
                    writer,
                    sheet_name="Wilayah",
                    index=False
                )


            # Insight

            insight_text = insights_bim(
                bim_df
            )

            insight_df = pd.DataFrame({

                "Insight": [
                    x.replace("**", "")
                    for x in insight_text
                ]

            })

            insight_df.to_excel(
                writer,
                sheet_name="Insight Analisis",
                index=False
            )


    return output.getvalue()


# =========================================================
# HEADER
# =========================================================

# =========================================================
# HEADER DASHBOARD
# =========================================================

st.title("📊 Dashboard Analisis Produktivitas 2026")

st.caption(
    "Pelatihan Produktivitas & Bimbingan Konsultasi"
)

st.divider()


# =========================================================
# UPLOAD
# =========================================================

uploaded = st.file_uploader(
    "📁 Upload File Excel",
    type=["xlsx"],
    help="Upload file Excel data produktivitas."
)


if not uploaded:

    st.info(
        "👋 **Selamat Datang!**\n\n"
        "Silakan upload file Excel untuk memulai "
        "analisis data Pelatihan Produktivitas "
        "dan Bimbingan Konsultasi.\n\n"
        "Sistem akan otomatis membaca, membersihkan, "
        "menstandarkan, memfilter, menganalisis, "
        "dan menampilkan data."
    )

    st.stop()


# =========================================================
# PROSES DATA
# =========================================================

with st.spinner(
    "⏳ Membaca dan membersihkan data..."
):

    sheets = load_excel(
        uploaded.getvalue()
    )


    train = (

        clean_training(
            sheets[TRAIN_SHEET]
        )

        if TRAIN_SHEET in sheets

        else pd.DataFrame()
    )


    bim = (

        clean_bim(
            sheets[BIM_SHEET]
        )

        if BIM_SHEET in sheets

        else pd.DataFrame()
    )


st.success(
    f"✅ Data berhasil diproses — "
    f"{format_number(len(train))} data pelatihan dan "
    f"{format_number(len(bim))} data bimbingan."
)


# =========================================================
# SIDEBAR
# =========================================================

st.sidebar.markdown(
    "## 📊 Dashboard"
)

st.sidebar.caption(
    "Analisis Produktivitas 2026"
)

st.sidebar.divider()


page = st.sidebar.radio(
    "Menu",
    [
        "🏠 Overview",
        "🎓 Pelatihan Produktivitas",
        "🤝 Bimbingan Konsultasi",
        "📋 Data Detail"
    ]
)


st.sidebar.divider()


# =========================================================
# OVERVIEW
# =========================================================

if page == "🏠 Overview":

    st.header("🏠 Overview Produktivitas 2026")
    st.caption("Ringkasan pelaksanaan Pelatihan Produktivitas dan Bimbingan Konsultasi")
    st.divider()


    total_program = (

        train[
            "Judul Program Pelatihan"
        ].nunique()

        if "Judul Program Pelatihan"
        in train.columns

        else 0
    )


    total_perusahaan = (

        bim[
            "NAMA PERUSAHAAN"
        ].nunique()

        if "NAMA PERUSAHAAN"
        in bim.columns

        else 0
    )


    a, b, c, d = st.columns(4)


    with a:

        st.metric(
            "👥 Total Peserta",
            format_number(
                len(train)
            )
        )


    with b:

        st.metric(
            "📚 Total Program",
            format_number(
                total_program
            )
        )


    with c:

        st.metric(
            "🤝 Total Bimbingan",
            format_number(
                len(bim)
            )
        )


    with d:

        st.metric(
            "🏢 Total Perusahaan",
            format_number(
                total_perusahaan
            )
        )


    st.markdown(
        "### 📈 Gambaran Aktivitas"
    )


    col1, col2 = st.columns(2)


    # -----------------------------------------------------
    # PESERTA
    # -----------------------------------------------------

    with col1:

        x = monthly_data(
            train,
            "Bulan Pelatihan",
            "Peserta"
        )

        fig = px.line(
            x,
            x="Bulan",
            y="Peserta",
            markers=True,
            text="Peserta",
            title="🎓 Tren Peserta Pelatihan"
        )

        fig.update_traces(
            line_width=4,
            textposition="top center"
        )

        st.plotly_chart(
            style_chart(fig),
            use_container_width=True
        )


    # -----------------------------------------------------
    # BIMBINGAN
    # -----------------------------------------------------

    with col2:

        x = monthly_data(
            bim,
            "Bulan Bimbingan",
            "Kegiatan"
        )

        fig = px.bar(
            x,
            x="Bulan",
            y="Kegiatan",
            text="Kegiatan",
            title="🤝 Tren Bimbingan Konsultasi"
        )

        fig.update_traces(
            textposition="outside"
        )

        st.plotly_chart(
            style_chart(fig),
            use_container_width=True
        )


    # -----------------------------------------------------
    # INSIGHT
    # -----------------------------------------------------

    st.markdown(
        "### 💡 Temuan Utama"
    )


    ta = insights_training(
        train
    )

    ba = insights_bim(
        bim
    )


    col1, col2 = st.columns(2)


    with col1:

        st.markdown(
            "#### 🎓 Pelatihan"
        )

        for text in ta:

            st.info(text)


    with col2:

        st.markdown(
            "#### 🤝 Bimbingan"
        )

        for text in ba:

            st.info(text)


    # -----------------------------------------------------
    # KESIMPULAN
    # -----------------------------------------------------

    st.subheader("🎯 Kesimpulan Analisis")
    st.info(
        "Dashboard memberikan gambaran mengenai jumlah peserta, program pelatihan, "
        "wilayah, kegiatan bimbingan, perusahaan, bidang usaha, serta perkembangan "
        "aktivitas sepanjang periode data yang tersedia.\n\n"
        "Informasi tersebut dapat digunakan sebagai dasar monitoring, evaluasi, "
        "dan penyusunan laporan pelaksanaan program produktivitas."
    )


# =========================================================
# PELATIHAN
# =========================================================

elif page == "🎓 Pelatihan Produktivitas":

    st.header("🎓 Pelatihan Produktivitas")
    st.caption("Analisis peserta, program, provinsi, kelulusan, dan metode pelatihan")
    st.divider()


    st.sidebar.markdown(
        "### 🔎 Filter Pelatihan"
    )


    f = train.copy()


    # Filter

    for col, label, key in [

        (
            "Tahun Pelatihan",
            "Tahun",
            "train_year"
        ),

        (
            "Provinsi",
            "Provinsi",
            "train_province"
        ),

        (
            "Metode Pelatihan",
            "Metode",
            "train_method"
        ),

        (
            "Jenis Pelatihan",
            "Jenis Pelatihan",
            "train_type"
        ),

        (
            "Status KeLulusan",
            "Status Kelulusan",
            "train_status"
        )

    ]:

        f = selectbox_filter(
            f,
            col,
            label,
            key
        )


    st.caption(
        f"📌 Menampilkan **{format_number(len(f))} peserta** "
        "sesuai filter."
    )


    # -----------------------------------------------------
    # KPI
    # -----------------------------------------------------

    a, b, c, d = st.columns(4)


    with a:

        st.metric(
            "👥 Total Peserta",
            format_number(
                len(f)
            )
        )


    with b:

        total_program = (

            f[
                "Judul Program Pelatihan"
            ].nunique()

            if "Judul Program Pelatihan"
            in f.columns

            else 0
        )

        st.metric(
            "📚 Total Program",
            format_number(
                total_program
            )
        )


    with c:

        total_provinsi = (

            f[
                "Provinsi"
            ].nunique()

            if "Provinsi"
            in f.columns

            else 0
        )

        st.metric(
            "📍 Total Provinsi",
            format_number(
                total_provinsi
            )
        )


    with d:

        total_lembaga = (

            f[
                "Nama Lembaga"
            ].nunique()

            if "Nama Lembaga"
            in f.columns

            else 0
        )

        st.metric(
            "🏢 Total Lembaga",
            format_number(
                total_lembaga
            )
        )


    # -----------------------------------------------------
    # ANALISIS
    # -----------------------------------------------------

    st.markdown(
        "### 📊 Analisis Pelatihan"
    )


    col1, col2 = st.columns(2)


    # Tren peserta

    with col1:

        x = monthly_data(
            f,
            "Bulan Pelatihan",
            "Peserta"
        )

        fig = px.line(
            x,
            x="Bulan",
            y="Peserta",
            markers=True,
            text="Peserta",
            title="📈 Peserta per Bulan"
        )

        fig.update_traces(
            line_width=4,
            textposition="top center"
        )

        st.plotly_chart(
            style_chart(fig),
            use_container_width=True
        )


    # Top provinsi

    with col2:

        if "Provinsi" in f.columns:

            x = (
                f["Provinsi"]
                .fillna("Tidak Diisi")
                .value_counts()
                .head(10)
                .sort_values()
                .reset_index()
            )

            x.columns = [
                "Provinsi",
                "Peserta"
            ]

            fig = px.bar(
                x,
                x="Peserta",
                y="Provinsi",
                orientation="h",
                text="Peserta",
                title="🏆 Top 10 Provinsi"
            )

            fig.update_traces(
                textposition="outside"
            )

            st.plotly_chart(
                style_chart(fig),
                use_container_width=True
            )


    # -----------------------------------------------------
    # STATUS + METODE
    # -----------------------------------------------------

    col1, col2 = st.columns(2)


    with col1:

        if "Status KeLulusan" in f.columns:

            x = (
                f[
                    "Status KeLulusan"
                ]
                .fillna("Tidak Diisi")
                .value_counts()
                .reset_index()
            )

            x.columns = [
                "Status",
                "Jumlah"
            ]

            fig = px.pie(
                x,
                names="Status",
                values="Jumlah",
                hole=0.55,
                title="🎓 Status Kelulusan"
            )

            fig.update_traces(
                textinfo="percent+label"
            )

            st.plotly_chart(
                style_chart(fig, 350),
                use_container_width=True
            )


    with col2:

        if "Metode Pelatihan" in f.columns:

            x = (
                f[
                    "Metode Pelatihan"
                ]
                .fillna("Tidak Diisi")
                .value_counts()
                .reset_index()
            )

            x.columns = [
                "Metode",
                "Jumlah"
            ]

            fig = px.bar(
                x,
                x="Metode",
                y="Jumlah",
                text="Jumlah",
                title="⚙️ Metode Pelatihan"
            )

            fig.update_traces(
                textposition="outside"
            )

            st.plotly_chart(
                style_chart(fig, 350),
                use_container_width=True
            )


    # -----------------------------------------------------
    # INSIGHT
    # -----------------------------------------------------

    st.markdown(
        "### 💡 Insight Otomatis"
    )


    if f.empty:

        st.warning(
            "Tidak ada data sesuai filter."
        )

    else:

        for text in insights_training(f):

            st.info(text)


    # -----------------------------------------------------
    # DOWNLOAD HASIL ANALISIS
    # -----------------------------------------------------

    st.subheader("📥 Download Hasil Analisis")
    st.caption(
        "Download PDF berisi hasil akhir dashboard sesuai filter aktif: KPI, grafik, "
        "insight, dan kesimpulan analisis."
    )

    if not f.empty:
        pdf_training = create_dashboard_pdf(training_df=f)
        st.download_button(
            label="📄 Download Hasil Dashboard Pelatihan (PDF)",
            data=pdf_training,
            file_name="Hasil_Dashboard_Pelatihan_Produktivitas_2026.pdf",
            mime="application/pdf",
            use_container_width=True
        )


# =========================================================
# BIMBINGAN
# =========================================================

elif page == "🤝 Bimbingan Konsultasi":

    st.header("🤝 Bimbingan Konsultasi")
    st.caption("Analisis kegiatan, perusahaan, wilayah, bidang usaha, dan pelaksana")
    st.divider()


    st.sidebar.markdown(
        "### 🔎 Filter Bimbingan"
    )


    f = bim.copy()


    for col, label, key in [

        (
            "Bidang Usaha Kategori",
            "Bidang Usaha",
            "bim_business"
        ),

        (
            "NAMA KABUPATEN/KOTA",
            "Kabupaten/Kota",
            "bim_region"
        ),

        (
            "PELAKSANA",
            "Pelaksana",
            "bim_executor"
        ),

        (
            "Tahun Bimbingan",
            "Tahun",
            "bim_year"
        )

    ]:

        f = selectbox_filter(
            f,
            col,
            label,
            key
        )


    st.caption(
        f"📌 Menampilkan **{format_number(len(f))} kegiatan** "
        "sesuai filter."
    )


    # -----------------------------------------------------
    # KPI
    # -----------------------------------------------------

    a, b, c, d = st.columns(4)


    with a:

        st.metric(
            "📋 Total Kegiatan",
            format_number(
                len(f)
            )
        )


    with b:

        total_perusahaan = (

            f[
                "NAMA PERUSAHAAN"
            ].nunique()

            if "NAMA PERUSAHAAN"
            in f.columns

            else 0
        )

        st.metric(
            "🏢 Total Perusahaan",
            format_number(
                total_perusahaan
            )
        )


    with c:

        total_wilayah = (

            f[
                "NAMA KABUPATEN/KOTA"
            ].nunique()

            if "NAMA KABUPATEN/KOTA"
            in f.columns

            else 0
        )

        st.metric(
            "📍 Total Wilayah",
            format_number(
                total_wilayah
            )
        )


    with d:

        total_bidang = (

            f[
                "Bidang Usaha Kategori"
            ].nunique()

            if "Bidang Usaha Kategori"
            in f.columns

            else 0
        )

        st.metric(
            "🏭 Bidang Usaha",
            format_number(
                total_bidang
            )
        )


    # -----------------------------------------------------
    # GRAFIK
    # -----------------------------------------------------

    st.markdown(
        "### 📊 Analisis Bimbingan"
    )


    col1, col2 = st.columns(2)


    # Bidang usaha

    with col1:

        if "Bidang Usaha Kategori" in f.columns:

            x = (
                f[
                    "Bidang Usaha Kategori"
                ]
                .fillna("Tidak Diisi")
                .value_counts()
                .sort_values()
                .reset_index()
            )

            x.columns = [
                "Bidang Usaha",
                "Jumlah"
            ]

            fig = px.bar(
                x,
                x="Jumlah",
                y="Bidang Usaha",
                orientation="h",
                text="Jumlah",
                title="🏭 Bimbingan per Bidang Usaha"
            )

            fig.update_traces(
                textposition="outside"
            )

            st.plotly_chart(
                style_chart(fig),
                use_container_width=True
            )


    # Wilayah

    with col2:

        if "NAMA KABUPATEN/KOTA" in f.columns:

            x = (
                f[
                    "NAMA KABUPATEN/KOTA"
                ]
                .fillna("Tidak Diisi")
                .value_counts()
                .head(10)
                .sort_values()
                .reset_index()
            )

            x.columns = [
                "Wilayah",
                "Jumlah"
            ]

            fig = px.bar(
                x,
                x="Jumlah",
                y="Wilayah",
                orientation="h",
                text="Jumlah",
                title="📍 Top 10 Wilayah"
            )

            fig.update_traces(
                textposition="outside"
            )

            st.plotly_chart(
                style_chart(fig),
                use_container_width=True
            )


    # -----------------------------------------------------
    # BULAN
    # -----------------------------------------------------

    x = monthly_data(
        f,
        "Bulan Bimbingan",
        "Kegiatan"
    )

    fig = px.line(
        x,
        x="Bulan",
        y="Kegiatan",
        markers=True,
        text="Kegiatan",
        title="📅 Jumlah Bimbingan per Bulan"
    )

    fig.update_traces(
        line_width=4,
        textposition="top center"
    )

    st.plotly_chart(
        style_chart(fig),
        use_container_width=True
    )


    # -----------------------------------------------------
    # INSIGHT
    # -----------------------------------------------------

    st.markdown(
        "### 💡 Insight Otomatis"
    )


    if f.empty:

        st.warning(
            "Tidak ada data sesuai filter."
        )

    else:

        for text in insights_bim(f):

            st.info(text)


    # -----------------------------------------------------
    # KESIMPULAN
    # -----------------------------------------------------

    st.subheader("🎯 Kesimpulan Analisis Bimbingan")
    st.info(
        "Analisis bimbingan memberikan gambaran mengenai konsentrasi kegiatan "
        "berdasarkan bidang usaha, wilayah, dan waktu pelaksanaan.\n\n"
        "Data tersebut dapat digunakan untuk monitoring pemerataan pendampingan "
        "serta identifikasi wilayah atau bidang usaha yang masih membutuhkan "
        "perhatian lebih lanjut."
    )


    # -----------------------------------------------------
    # DOWNLOAD
    # -----------------------------------------------------

    st.subheader("📥 Download Hasil Analisis")
    st.caption(
        "Download PDF berisi hasil akhir dashboard sesuai filter aktif: KPI, grafik, "
        "insight, dan kesimpulan analisis."
    )

    if not f.empty:
        pdf_bim = create_dashboard_pdf(bim_df=f)
        st.download_button(
            label="📄 Download Hasil Dashboard Bimbingan (PDF)",
            data=pdf_bim,
            file_name="Hasil_Dashboard_Bimbingan_Konsultasi_2026.pdf",
            mime="application/pdf",
            use_container_width=True
        )


# =========================================================
# DATA DETAIL
# =========================================================

else:

    st.header("📋 Data Detail & Hasil Cleaning")
    st.caption("Data setelah proses cleaning dan standardisasi")
    st.divider()


    tipe = st.radio(
        "Pilih data",
        [
            "Pelatihan",
            "Bimbingan"
        ],
        horizontal=True
    )


    df = (
        train
        if tipe == "Pelatihan"
        else bim
    )


    a, b = st.columns(2)


    with a:

        st.metric(
            "📊 Jumlah Baris",
            format_number(
                len(df)
            )
        )


    with b:

        st.metric(
            "🧹 Jumlah Kolom",
            format_number(
                len(df.columns)
            )
        )


    st.dataframe(
        df,
        use_container_width=True,
        height=550
    )


    # =====================================================
    # DOWNLOAD DATA CLEANING
    # =====================================================

    output = BytesIO()


    df.to_excel(
        output,
        index=False,
        engine="openpyxl"
    )


    st.download_button(

        label="⬇️ Download Data Hasil Cleaning",

        data=output.getvalue(),

        file_name=(
            f"data_{tipe.lower()}_clean.xlsx"
        ),

        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),

        use_container_width=True
    )


# =========================================================
# FOOTER
# =========================================================

st.divider()

st.caption(
    "📊 Dashboard Analisis Produktivitas 2026 • Data mengikuti file Excel yang diunggah "
    "• Cleaning dan standardisasi dilakukan otomatis"
)