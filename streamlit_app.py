import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Fungsi untuk mengekstrak 6 digit pertama Nomor SP2D dari kolom Keterangan
def extract_sp2d_number(description):
    match = re.search(r'\b\d{6}\b', str(description))  # Pastikan description adalah string
    return match.group(0) if match else None

# Fungsi untuk melakukan vouching
def perform_vouching(rk_df, sp2d_df):
    # Inisialisasi list untuk menyimpan hasil
    matched_rk = []
    unmatched_rk = []
    matched_sp2d = []
    unmatched_sp2d = []

    # Membuat dictionary untuk SP2D agar pencarian lebih cepat
    sp2d_dict = {}
    for _, row in sp2d_df.iterrows():
        no_sp2d = str(row.get('NoSP2D', ''))  # Konversi ke string, default '' jika tidak ada
        if len(no_sp2d) >= 6:  # Pastikan panjang >= 6
            sp2d_dict[no_sp2d[:6]] = row

    # Proses vouching untuk setiap baris di RK
    progress_bar = st.progress(0)
    total_rows = len(rk_df)
    for i, rk_row in rk_df.iterrows():
        sp2d_number = extract_sp2d_number(rk_row.get('Keterangan', ''))
        if sp2d_number and sp2d_number in sp2d_dict:
            sp2d_row = sp2d_dict[sp2d_number]
            if rk_row.get('Jumlah') == sp2d_row.get('Jumlah'):
                matched_rk.append({
                    'Tanggal': rk_row.get('Tanggal'),
                    'Keterangan': rk_row.get('Keterangan'),
                    'Jumlah (Rp)': rk_row.get('Jumlah'),
                    'NO SP2D': sp2d_row.get('NoSP2D')
                })
                matched_sp2d.append(sp2d_row)
            else:
                unmatched_rk.append({
                    'Tanggal': rk_row.get('Tanggal'),
                    'Keterangan': rk_row.get('Keterangan'),
                    'Jumlah (Rp)': rk_row.get('Jumlah'),
                    'NO SP2D': sp2d_number
                })
        else:
            unmatched_rk.append({
                'Tanggal': rk_row.get('Tanggal'),
                'Keterangan': rk_row.get('Keterangan'),
                'Jumlah (Rp)': rk_row.get('Jumlah'),
                'NO SP2D': sp2d_number
            })

        # Update progress bar
        progress_bar.progress((i + 1) / total_rows)

    # Menentukan transaksi SP2D yang tidak cocok
    sp2d_numbers_in_rk = {row['NO SP2D'] for row in matched_rk}
    for _, sp2d_row in sp2d_df.iterrows():
        no_sp2d = str(sp2d_row.get('NoSP2D', ''))
        if no_sp2d[:6] not in sp2d_numbers_in_rk:
            unmatched_sp2d.append(sp2d_row)

    # Mengonversi hasil ke DataFrame
    matched_rk_df = pd.DataFrame(matched_rk)
    unmatched_rk_df = pd.DataFrame(unmatched_rk)
    matched_sp2d_df = pd.DataFrame(matched_sp2d)
    unmatched_sp2d_df = pd.DataFrame(unmatched_sp2d)

    return matched_rk_df, unmatched_rk_df, matched_sp2d_df, unmatched_sp2d_df

# Antarmuka Streamlit
st.title("Aplikasi Vouching SP2D vs Rekening Koran")

# Upload file
rk_file = st.file_uploader("Upload File Rekening Koran (Excel)", type=["xlsx"])
sp2d_file = st.file_uploader("Upload File SP2D (Excel)", type=["xlsx"])

if rk_file and sp2d_file:
    try:
        # Baca file Excel
        rk_df = pd.read_excel(rk_file)
        sp2d_df = pd.read_excel(sp2d_file)

        # Validasi kolom
        required_columns_rk = {'Tanggal', 'Keterangan', 'Jumlah'}
        required_columns_sp2d = {'SKPD', 'NoSP2D', 'TglSP2D', 'Jumlah'}

        if not required_columns_rk.issubset(rk_df.columns):
            st.error(f"File Rekening Koran tidak memiliki kolom yang diperlukan: {required_columns_rk}")
            st.stop()

        if not required_columns_sp2d.issubset(sp2d_df.columns):
            st.error(f"File SP2D tidak memiliki kolom yang diperlukan: {required_columns_sp2d}")
            st.stop()

        # Membersihkan data
        sp2d_df['NoSP2D'] = sp2d_df['NoSP2D'].astype(str).str.strip()  # Konversi ke string dan hapus spasi
        sp2d_df = sp2d_df[sp2d_df['NoSP2D'].str.len() >= 6]  # Hanya simpan baris dengan NoSP2D >= 6 karakter
        rk_df['Keterangan'] = rk_df['Keterangan'].astype(str).str.strip()  # Konversi ke string dan hapus spasi

        # Tombol untuk memulai vouching
        if st.button("Mulai Vouching"):
            matched_rk_df, unmatched_rk_df, matched_sp2d_df, unmatched_sp2d_df = perform_vouching(rk_df, sp2d_df)

            # Tampilkan hasil
            st.subheader("Hasil Vouching")
            st.write("Rekening Koran Matched:")
            st.dataframe(matched_rk_df)
            st.write("Rekening Koran Unmatched:")
            st.dataframe(unmatched_rk_df)
            st.write("SP2D Matched:")
            st.dataframe(matched_sp2d_df)
            st.write("SP2D Unmatched:")
            st.dataframe(unmatched_sp2d_df)

            # Simpan hasil ke file Excel
            excel_data = save_to_excel(matched_rk_df, unmatched_rk_df, matched_sp2d_df, unmatched_sp2d_df)
            st.download_button(
                label="Unduh Hasil Vouching",
                data=excel_data,
                file_name="hasil_vouching.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
