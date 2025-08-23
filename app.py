# ===============================
# app_streamlit.py (Highlight ID=1000 + Surat Panggilan)
# ===============================
import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
import re
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# --- Konfigurasi upload ---
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- Helper ---
def clean_id(id_value):
    if pd.isna(id_value):
        return ""
    id_str = str(id_value).strip()
    if id_str.endswith('.0'):
        id_str = id_str[:-2]
    return id_str

def sort_nicely(l):
    def convert(text):
        return int(text) if text.isdigit() else text.lower()
    def alphanum_key(key):
        return [convert(c) for c in re.split('([0-9]+)', key)]
    return sorted(l, key=alphanum_key)

# --- Streamlit UI ---
st.set_page_config(page_title="Rekap Absensi PT. QUANTUM", layout="wide")
st.title("📑 Rekap Absensi PT. QUANTUM")

uploaded_file = st.file_uploader("Upload File Absensi (.xlsx/.xls)", type=["xlsx", "xls"])

if uploaded_file:
    # Simpan sementara
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = os.path.join(UPLOAD_FOLDER, f"{timestamp}_{uploaded_file.name}")
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # === PROSES EXCEL ===
    df = pd.read_excel(file_path)
    max_cols = len(df.columns)
    column_mapping = {}
    column_names = ['Perusahaan', 'Nama', 'ID', 'Tgl/Waktu', 'Mesin_ID', 'Kolom6', 'Status', 'Kolom8']
    for i in range(min(max_cols, len(column_names))):
        column_mapping[column_names[i]] = df.iloc[:, i]
    df_fix = pd.DataFrame(column_mapping)

    # ID unik
    semua_id_dari_file = [clean_id(idv) for idv in df.iloc[:,2] if clean_id(idv) != ""]
    semua_id_unik = sort_nicely(list(set(semua_id_dari_file)))

    # Normalisasi
    df_fix["Nama"] = df_fix["Nama"].astype(str).str.strip()
    df_fix["ID"] = df_fix["ID"].apply(clean_id)
    df_fix = df_fix[df_fix["Nama"].notna() & (df_fix["Nama"] != "nan") & (df_fix["Nama"] != "")]
    df_fix = df_fix[df_fix["ID"].notna() & (df_fix["ID"] != "nan") & (df_fix["ID"] != "")]
    df_fix["Tgl/Waktu"] = pd.to_datetime(df_fix["Tgl/Waktu"], dayfirst=True, errors='coerce')
    df_fix = df_fix.dropna(subset=["Tgl/Waktu"])
    df_fix["Tanggal_Saja"] = df_fix["Tgl/Waktu"].dt.date
    df_fix = df_fix.drop_duplicates(subset=["ID", "Tanggal_Saja"])

    # Rentang tanggal kerja (tanpa Minggu)
    if not df_fix["Tgl/Waktu"].empty:
        tanggal_awal = df_fix["Tgl/Waktu"].dt.date.min()
        tanggal_akhir = df_fix["Tgl/Waktu"].dt.date.max()
        semua_tanggal = [tgl for tgl in pd.date_range(tanggal_awal, tanggal_akhir).date if pd.Timestamp(tgl).weekday() != 6]
    else:
        semua_tanggal = []

    # Rekap tidak hadir + jumlah hadir
    rekap_tidak_hadir = []
    jumlah_absen_total = []
    id_to_nama = dict(zip(df_fix["ID"], df_fix["Nama"]))
    for id_karyawan in semua_id_unik:
        nama_karyawan = id_to_nama.get(id_karyawan, "Unknown")
        data_id = df_fix[df_fix["ID"]==id_karyawan]
        hadir_tanggal = set(data_id["Tgl/Waktu"].dt.date) if not data_id.empty else set()
        tidak_hadir_tanggal = [tgl for tgl in semua_tanggal if tgl not in hadir_tanggal]
        for tgl in tidak_hadir_tanggal:
            rekap_tidak_hadir.append({"ID": id_karyawan, "Nama": nama_karyawan, "Tanggal Tidak Hadir": tgl})
        hadir_per_tanggal = len(hadir_tanggal)
        jumlah_absen_total.append({"ID": id_karyawan, "Nama": nama_karyawan, "Jumlah Absen Total": hadir_per_tanggal})
    df_tidak_hadir = pd.DataFrame(rekap_tidak_hadir)
    df_jumlah_absen = pd.DataFrame(jumlah_absen_total)

    if not df_tidak_hadir.empty:
        jumlah_tidak_hadir = df_tidak_hadir.groupby("ID").size().reset_index(name="Jumlah Tidak Hadir")
        df_jumlah_absen = pd.merge(df_jumlah_absen, jumlah_tidak_hadir, on="ID", how="left")
    else:
        df_jumlah_absen["Jumlah Tidak Hadir"] = 0

    df_jumlah_absen["Jumlah Tidak Hadir"] = df_jumlah_absen["Jumlah Tidak Hadir"].fillna(0).astype(int)

    # --- TAMPILKAN DI STREAMLIT ---
    st.subheader("📌 Rekap Tidak Hadir")
    st.dataframe(df_tidak_hadir if not df_tidak_hadir.empty else pd.DataFrame([{"Info":"Tidak ada data karyawan tidak hadir"}]))

    # --- Highlight di Streamlit (ID=1000 hijau) ---
    def highlight_id(val):
        if str(val) == "1000":
            return "background-color: lightgreen; color: black; font-weight: bold;"
        return ""
    st.subheader("📌 Jumlah Kehadiran (Highlight ID 1000)")
    st.dataframe(df_jumlah_absen.style.applymap(highlight_id, subset=["ID"]))

    # === Simpan ke Excel + highlight ===
    hasil_rekap_path = os.path.join(UPLOAD_FOLDER, f"hasil_rekap_{uploaded_file.name}")
    with pd.ExcelWriter(hasil_rekap_path) as writer:
        if not df_tidak_hadir.empty:
            df_tidak_hadir.to_excel(writer, sheet_name="Karyawan Tidak Hadir", index=False)
        df_jumlah_absen.to_excel(writer, sheet_name="Jumlah Kehadiran", index=False)

    wb = load_workbook(hasil_rekap_path)
    ws = wb["Jumlah Kehadiran"]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            if str(cell.value) == "1000":
                cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                cell.font = Font(color="000000", bold=True)
    wb.save(hasil_rekap_path)

    with open(hasil_rekap_path, "rb") as f:
        st.download_button("📥 Download Rekap Excel", f, file_name=os.path.basename(hasil_rekap_path))

    # --- Surat Panggilan untuk yang tidak hadir > 3 ---
    df_tidak_hadir_lebih3 = df_jumlah_absen[df_jumlah_absen["Jumlah Tidak Hadir"]>3].copy()
    if not df_tidak_hadir_lebih3.empty:
        st.subheader("📌 Surat Panggilan")
        hari_list = ["Senin","Selasa","Rabu","Kamis","Jumat","Sabtu","Minggu"]
        for _, row in df_tidak_hadir_lebih3.iterrows():
            spg_filename = f"surat_panggilan_{row['ID']}_{uploaded_file.name.rsplit('.',1)[0]}.docx"
            spg_path = os.path.join(UPLOAD_FOLDER, spg_filename)
            template_path = os.path.join("templates", "template_surat_panggilan.docx")
            doc = DocxTemplate(template_path)

            df_absen_id = df_tidak_hadir[df_tidak_hadir["ID"]==row['ID']]
            semua_tgl = df_absen_id["Tanggal Tidak Hadir"].apply(lambda x: x.strftime("%d-%m-%Y")).tolist()
            tanggal_terakhir = ", ".join(semua_tgl)
            jumlah_hari = len(semua_tgl)

            tanggal_surat = date.today()
            nama_hari = hari_list[tanggal_surat.weekday()]

            context = {
                "NAMA": row['Nama'],
                "ID": row['ID'],
                "JUMLAH_HARI": jumlah_hari,
                "TANGGAL_ABSEN": tanggal_terakhir,
                "TANGGAL_SURAT": f"{nama_hari}, {tanggal_surat.strftime('%d-%m-%Y')}"
            }

            doc.render(context)
            doc.save(spg_path)

            with open(spg_path, "rb") as f:
                st.download_button(f"📥 Download Surat Panggilan untuk {row['Nama']}", f, file_name=spg_filename)

