import streamlit as st
import pandas as pd

st.set_page_config(page_title="Rekap Absensi PT. QUANTUM", layout="wide")

# Judul
st.title("📑 Rekap Absensi PT. QUANTUM")
st.write("Upload file absensi karyawan, lalu sistem akan otomatis menampilkan rekap 📊")

# Upload file
uploaded_file = st.file_uploader("📂 Upload File Excel Absensi", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # --- Dummy contoh logika absensi ---
        df["Telat"] = df["Jam Masuk"] > "09:00:00"

        df_telat = df[df["Telat"]]
        df_tidak_hadir = df[df["Keterangan"] == "Tidak Hadir"]
        df_jumlah_absen = df.groupby("Nama").size().reset_index(name="Jumlah Hadir")
        df_tidak_hadir_lebih3 = df_tidak_hadir.groupby("Nama").size().reset_index(name="Jumlah Tidak Hadir")
        df_tidak_hadir_lebih3 = df_tidak_hadir_lebih3[df_tidak_hadir_lebih3["Jumlah Tidak Hadir"] > 3]

        # --- Statistik ringkas ---
        jumlah_karyawan_telat = len(df_telat["Nama"].unique())
        jumlah_karyawan_tidak_hadir = len(df_tidak_hadir["Nama"].unique())
        jumlah_total_karyawan = len(df["Nama"].unique())

        # --- Tampilkan 3 kotak statistik ---
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("⏰ Karyawan Telat", jumlah_karyawan_telat)
        with col2:
            st.metric("🚫 Tidak Hadir", jumlah_karyawan_tidak_hadir)
        with col3:
            st.metric("👥 Total Karyawan", jumlah_total_karyawan)

        st.markdown("---")

        # --- Tab untuk tabel ---
        tab1, tab2, tab3, tab4 = st.tabs(["Telat", "Tidak Hadir", "Jumlah Absen", "Tidak Hadir > 3 Hari"])

        with tab1:
            st.subheader("⏰ Daftar Karyawan Telat")
            st.dataframe(df_telat)

        with tab2:
            st.subheader("🚫 Daftar Tidak Hadir")
            st.dataframe(df_tidak_hadir)

        with tab3:
            st.subheader("📊 Jumlah Absen per Karyawan")
            st.dataframe(df_jumlah_absen)

        with tab4:
            st.subheader("⚠️ Karyawan Tidak Hadir Lebih dari 3 Hari")
            if not df_tidak_hadir_lebih3.empty:
                st.dataframe(df_tidak_hadir_lebih3)
                st.download_button("📄 Download Surat Panggilan",
                                   data="Surat Panggilan Dummy",
                                   file_name="panggilan.docx")
            else:
                st.info("Tidak ada karyawan dengan ketidakhadiran lebih dari 3 hari ✅")

        st.markdown("---")
        st.success("✅ Data berhasil diproses!")

    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
