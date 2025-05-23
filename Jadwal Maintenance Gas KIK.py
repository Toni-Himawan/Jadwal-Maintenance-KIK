import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io

st.set_page_config(page_title="Jadwal Maintenance & Shift Operator", layout="centered")

st.title("📋 Jadwal Maintenance dan Shift Operator")

uploaded_file = st.file_uploader("Unggah file Excel (.xlsx) berisi dua sheet:", type="xlsx")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Jadwal Shift')
    return output.getvalue()

def highlight_shift(val):
    if val == "Pagi":
        return 'background-color: #d4edda'  # hijau muda
    elif val == "Siang":
        return 'background-color: #fff3cd'  # kuning muda
    elif val == "Malam":
        return 'background-color: #f8d7da'  # merah muda
    elif val == "Libur":
        return 'background-color: #e2e3e5'  # abu muda
    return ''

def highlight_today(row):
    today = datetime.now().strftime('%d %B %Y')
    if row['Tanggal'] == today:
        return ['background-color: #cce5ff; font-weight: bold'] * len(row)
    return [''] * len(row)

hari_dict = {
    'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu',
    'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu', 'Sunday': 'Minggu'
}
bulan_dict = {
    1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
    5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
    9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
}

# Fungsi untuk membuat laporan PDF
def create_pdf(sheet1_filtered, sheet2_filtered, bulan_pilihan_nama, minggu_pilihan, file_name="laporan_jadwal_shift.pdf"):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)

    # Menambahkan judul laporan
    c.setFont("Helvetica-Bold", 14)
    c.drawString(200, 750, "Laporan Jadwal Shift dan Maintenance")
    c.setFont("Helvetica", 10)

    # Menambahkan tanggal laporan
    today = datetime.now().strftime('%d %B %Y')
    c.drawString(30, 730, f"Tanggal Laporan: {today}")
    
    # Menambahkan bagian Jadwal Maintenance
    c.setFont("Helvetica-Bold", 12)
    c.drawString(30, 710, f"Jadwal Maintenance Bulan: {bulan_pilihan_nama} - Minggu Ke: {minggu_pilihan}")
    c.setFont("Helvetica", 10)

    # Menambahkan tabel Jadwal Maintenance
    y_position = 690
    for col in sheet1_filtered.columns:
        c.drawString(30, y_position, col)
        y_position -= 20

    for index, row in sheet1_filtered.iterrows():
        y_position -= 20
        for value in row:
            c.drawString(100, y_position, str(value))
            y_position += 40

    # Menambahkan bagian Jadwal Shift Operator
    y_position -= 60
    c.setFont("Helvetica-Bold", 12)
    c.drawString(30, y_position, "Jadwal Shift Operator")
    c.setFont("Helvetica", 10)

    y_position -= 20
    for col in sheet2_filtered.columns:
        c.drawString(30, y_position, col)
        y_position -= 20

    for index, row in sheet2_filtered.iterrows():
        y_position -= 20
        for value in row:
            c.drawString(100, y_position, str(value))
            y_position += 40

    c.save()

    buffer.seek(0)
    return buffer

if uploaded_file:
    try:
        sheet1 = pd.read_excel(uploaded_file, sheet_name=0)
        sheet2 = pd.read_excel(uploaded_file, sheet_name=1)

        sheet1['Tanggal'] = pd.to_datetime(sheet1['Tanggal'])
        try:
            sheet1['Hari'] = pd.to_datetime(sheet1['Hari']).dt.day_name().map(hari_dict)
        except:
            pass

        sheet2['Tanggal'] = pd.to_datetime(sheet2['Tanggal'])
        try:
            sheet2['Hari'] = pd.to_datetime(sheet2['Hari']).dt.day_name().map(hari_dict)
        except:
            sheet2['Hari'] = sheet2['Hari'].map(hari_dict).fillna(sheet2['Hari'])

        # Menampilkan Jadwal Maintenance
        st.header("🛠️ Jadwal Maintenance")
        kolom_maintenance = ['Tanggal', 'Hari', 'Kegiatan', 'Jumlah Titik/Item', 'Minggu Ke']
        if all(col in sheet1.columns for col in kolom_maintenance):
            bulan_list = sheet1['Tanggal'].dt.month.unique()
            bulan_nama_list = [bulan_dict.get(b, str(b)) for b in bulan_list]
            bulan_pilihan_nama = st.selectbox("Pilih Bulan Maintenance", options=bulan_nama_list)
            bulan_pilihan = [k for k,v in bulan_dict.items() if v == bulan_pilihan_nama][0]

            filtered = sheet1[sheet1['Tanggal'].dt.month == bulan_pilihan]
            minggu_list = sorted(filtered['Minggu Ke'].unique())
            minggu_pilihan = st.selectbox("Pilih Minggu Ke", options=minggu_list)

            sheet1_filtered = filtered[filtered['Minggu Ke'] == minggu_pilihan]
            sheet1_filtered['Tanggal'] = sheet1_filtered['Tanggal'].dt.strftime('%d %B %Y')
            if 'Bulan' in sheet1_filtered.columns:
                sheet1_filtered = sheet1_filtered.drop(columns=['Bulan'])

            st.subheader(f"Jadwal Maintenance {bulan_pilihan_nama}")
            st.dataframe(sheet1_filtered)
        else:
            st.warning("Kolom di sheet pertama tidak sesuai format yang diharapkan.")

        # Menampilkan Jadwal Shift Operator
        hari_ini = datetime.now()
        tanggal_indo = f"{hari_ini.day} {bulan_dict[hari_ini.month]} {hari_ini.year}"
        st.header("👷 Jadwal Shift Operator")
        st.markdown(f"📅 Tanggal hari ini: **{tanggal_indo}**")

        kolom_shift = ['Tanggal', 'Hari', 'Huda', 'Supriyanto', 'Anta']
        if all(col in sheet2.columns for col in kolom_shift):
            hari_valid = list(hari_dict.values())
            df_shift_filtered = sheet2[sheet2['Hari'].isin(hari_valid)].copy()
            df_shift_filtered['Tanggal'] = df_shift_filtered['Tanggal'].dt.strftime('%d %B %Y')

            tab1, tab2 = st.tabs(["📆 Lihat Semua Jadwal", "🔍 Filter Operator"])

            with tab1:
                st.subheader("Jadwal Lengkap Shift")
                styled_df = df_shift_filtered.style \
                    .applymap(highlight_shift, subset=['Huda', 'Supriyanto', 'Anta']) \
                    .apply(highlight_today, axis=1)
                st.dataframe(styled_df, use_container_width=True)

                df_xlsx = to_excel(df_shift_filtered)
                st.download_button("⬇️ Unduh Jadwal Shift (Excel)", df_xlsx, "jadwal_shift.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                # Laporan PDF
                pdf_buffer = create_pdf(sheet1_filtered, df_shift_filtered, bulan_pilihan_nama, minggu_pilihan)
                st.download_button("⬇️ Unduh Laporan Shift dan Maintenance (PDF)", pdf_buffer, "laporan_shift_maintenance.pdf", "application/pdf")

            with tab2:
                st.subheader("Filter Jadwal per Operator")
                operator_pilihan = st.selectbox("Pilih Operator", ['Huda', 'Supriyanto', 'Anta'])
                df_filtered = df_shift_filtered[['Tanggal', 'Hari', operator_pilihan]].rename(columns={operator_pilihan: 'Shift'})
                st.dataframe(df_filtered.style.applymap(highlight_shift, subset=['Shift']), use_container_width=True)
        else:
            st.warning("Kolom di sheet kedua tidak sesuai format yang diharapkan.")

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
else:
    st.info("Silakan unggah file Excel yang berisi dua sheet: (1) Jadwal Maintenance dan (2) Jadwal Shift Operator.")
