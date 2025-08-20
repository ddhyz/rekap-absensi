import streamlit as st
import pandas as pd

# ================= CONFIG ==================
st.set_page_config(page_title="Rekap Absensi PT. QUANTUM", layout="wide")

# ================= CSS CUSTOM ==================
st.markdown("""
<style>
    /* Background */
    .main {
        background: linear-gradient(135deg, #f5f7fa 0%, #e4edf5 100%);
    }
    /* Judul */
    .title {
        text-align:center; 
        font-weight:700; 
        font-size:1.8rem; 
        color:#0d6efd;
        margin-bottom:20px;
    }
    /* Card Statistik */
    .stat-card {
        padding:20px; 
        border-radius:15px; 
        text-align:center; 
        color:white; 
        box-shadow:0 6px 18px rgba(0,0,0,0.1);
        font-weight:600;
        transition: all .3s ease;
    }
    .stat-card:hover { transform: translateY(-5px); }
    .telat { background: linear-gradient(135deg, #ffc107 0%, #ffca2c 100%); }
    .tidak-hadir { background: linear-gradient(135deg, #6c757d 0%, #5c636a 100%); }
    .jumlah { background: linear-gradient(135deg, #198754 0%, #157347 100%); }
    /* Upload box */
    .uploadbox {
        border:2px dashed #0d6efd; 
        padding:30px; 
        border-radius:15px; 
        background:white;
        text-align:center;
        margin-bottom:25px;
    }
    /* Download Button */
    .stDownloadButton button {
        border-radius:12px; 
        padding:12px 25px;
        font-weight:600;
        background: linear-gradient(135deg, #6f42c1 0%, #5a32a3 100%);
        color:white;
        border:none;
    }
    .stDownloadButton button:hover {
        transform: translateY(-2px);
        box-shadow:0 6px 20px rgba(111,66,193,.4);
    }
</style>
""", unsafe_allow_html=True)

# ================= HEADER ==================
st.markdown("<div class='title'>📑 Rekap Absensi PT. QUANTUM</div>", unsafe_allow_html=True)

# ================= FILE UPLOAD ==================
st.markdown("<div class='uploadbox'>📂 Upload File Excel Absensi</div>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # --- Dummy logika absensi (samain sesuai kebutuhanmu) ---
        if "Jam Masuk" in df.columns:
            df["Telat"] = df["Jam Masuk"] > "09:00:00"
        else:
            df["Telat"] = False

        if "Keterangan" in df.columns:
            df_tidak_hadir = df[df["Keterangan"] == "Tidak Hadir"]
        else:
            df_tidak_hadir = pd.DataFrame()

        df_telat = df[df["Telat"]]
        df_jumlah_absen = df.groupby("Nama").size().reset_index(name="Jumlah Hadir")
        df_tidak_hadir_lebih3 = df_tidak_hadir.groupby("Nama").size().reset_index(name="Jumlah Tidak Hadir")
        df_tidak_hadir_lebih3 = df_tidak_hadir_lebih3[df_tidak_hadir_lebih3["Jumlah Tidak Hadir"] > 3]

        # --- Statistik Karyawan ---
        jumlah_karyawan_telat = len(df_telat["Nama"].unique())
        jumlah_karyawan_tidak_hadir = len(df_tidak_hadir["Nama"].unique())
        jumlah_total_karyawan = len(df["Nama"].unique())

        # ================= KOTAK STATISTIK ==================
        col1, col2, col3 = st.columns(3)
        col1.markdown(f"<div class='stat-card telat'>⏰<br><h3>{jumlah_karyawan_telat}</h3><p>Karyawan Telat</p></div>", unsafe_allow_html=True)
        col2.markdown(f"<div class='stat-card tidak-hadir'>🚫<br><h3>{jumlah_karyawan_tidak_hadir}</h3><p>Tidak Hadir</p></div>", unsafe_allow_html=True)
        col3.markdown(f"<div class='stat-card jumlah'>👥<br><h3>{jumlah_total_karyawan}</h3><p>Total Karyawan</p></div>", unsafe_allow_html=True)

        st.markdown("---")

        # ================= TAB HASIL ==================
        tab1, tab2, tab3, tab4 = st.tabs([
            "⏰ Telat", 
            "🚫 Tidak Hadir", 
            "📊 Jumlah Absen", 
            "⚠️ Tidak Hadir > 3 Hari"
        ])
        
        with tab1:
            st.dataframe(df_telat)
        with tab2:
            st.dataframe(df_tidak_hadir)
        with tab3:
            st.dataframe(df_jumlah_absen)
        with tab4:
            if not df_tidak_hadir_lebih3.empty:
                st.dataframe(df_tidak_hadir_lebih3)
                st.download_button("📄 Download Surat Panggilan", "Isi Surat Dummy", "panggilan.docx")
            else:
                st.info("Tidak ada karyawan dengan ketidakhadiran > 3 hari ✅")

        # ================= FOOTER ==================
        st.markdown("<p style='text-align:center; margin-top:40px; color:#6c757d;'>© 2025 PT. QUANTUM - Sistem Rekap Absensi Karyawan</p>", unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Gagal membaca file: {e}")

