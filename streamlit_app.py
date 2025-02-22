import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Fungsi untuk mengekstrak 6 digit pertama Nomor SP2D dari kolom Keterangan
def extract_sp2d_number(description):
    match = re.search(r'\b\d{6}\b', description)
    return match.group(0) if match else None

# Fungsi untuk melakukan vouching
def perform_vouching(rk_df, sp2d_df):
    # Inisialisasi list untuk menyimpan hasil
    matched_rk = []
    unmatched_rk = []
    matched_sp2d = []
    unmatched_sp2d = []

    # Membuat dictionary untuk SP2D agar pencarian lebih cepat
    sp2d_dict = {row['NoSP2D'][:6]: row for _, row in sp2d_df.iterrows()}

    # Proses vouching untuk setiap baris di RK
    progress_bar = st.progress(0)
    total_rows = len(rk_df)
    for i, rk_row in rk_df.iterrows():
        sp2d_number = extract_sp2d_number(rk_row['Keterangan'])
        if sp2d_number and sp2d_number in sp2d_dict:
            sp2d_row = sp2d_dict[sp2d_number]
            if rk_row['Jumlah'] == sp2d_row['Jumlah']:
                matched_rk.append({
                    'Tanggal': rk_row['Tanggal'],
                    'Keterangan': rk_row['Keterangan'],
                    'Jumlah (Rp)': rk_row['Jumlah'],
                    'NO SP2D': sp2d_row['NoSP2D']
                })
                matched_sp2d.append(sp2d_row)
            else:
                unmatched_rk.append({
                    'Tanggal': rk_row['Tanggal'],
                    'Keterangan': rk_row['Keterangan'],
                    'Jumlah (Rp)': rk_row['Jumlah'],
                    'NO SP2D': sp2d_number
                })
        else:
            unmatched_rk.append({
                'Tanggal': rk_row['Tanggal'],
                'Keterangan': rk_row['Keterangan'],
                'Jumlah (Rp)': rk_row['Jumlah'],
                'NO SP2D': sp2d_number
            })

        # Update progress bar
        progress_bar.progress((i + 1) / total_rows)

    # Menentukan transaksi SP2D yang tidak cocok
    sp2d_numbers_in_rk = {row['NO SP2D'] for row in matched_rk}
    for _, sp2d_row in sp2d_df.iterrows():
        if sp2d_row['NoSP2D'][:6] not in sp2d_numbers_in_rk:
            unmatched_sp2d.append(sp2d_row)

    # Mengonversi hasil ke DataFrame
    matched_rk_df = pd.DataFrame(matched_rk)
    unmatched_rk_df = pd.DataFrame(unmatched_rk)
    matched_sp2d_df = pd.DataFrame(matched_sp2d)
    unmatched_sp2d_df = pd.DataFrame(unmatched_sp2d)

    return matched_rk_df, unmatched_rk_df, matched_sp2d_df, unmatched_sp2d_df

# Fungsi untuk menyimpan hasil ke file Excel
def save_to_excel(matched_rk_df, unmatched_rk_df, matched_sp2d_df, unmatched_sp2d_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        matched_rk_df.to_excel(writer, sheet_name='RK_Matched', index=False)
        unmatched_rk_df.to_excel(writer, sheet_name='RK_Unmatched', index=False)
        matched_sp2d_df.to_excel(writer, sheet_name='SP2D_Matched', index=False)
        unmatched_sp2d_df.to_excel(writer, sheet_name='SP2D_Unmatched', index=False)
    output.seek(0)
    return output

# Antarmuka Streamlit
st.title("Aplikasi Vouching SP2D vs Rekening Koran")

# Upload file
rk_file = st.file_uploader("Upload File Rekening Koran (Excel)", type=["xlsx"])
sp2d_file = st.file_uploader("Upload File SP2D (Excel)", type=["xlsx"])

if rk_file and sp2d_file:
    # Baca file Excel
    rk_df = pd.read_excel(rk_file)
    sp2d_df = pd.read_excel(sp2d_file)

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
