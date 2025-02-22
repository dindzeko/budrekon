import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO

# Fungsi untuk mengekstrak 6 digit pertama Nomor SP2D dari kolom Keterangan
def extract_sp2d_number(description):
    description = str(description)
    for word in description.split():
        if word.isdigit() and len(word) >= 6:
            return word[:6]
    return None

# Fungsi untuk melakukan vouching utama
def perform_vouching(rk_df, sp2d_df):
    # Membuat dictionary untuk SP2D agar pencarian lebih cepat
    sp2d_dict = {str(row['NoSP2D'])[:6]: row for _, row in sp2d_df.iterrows()}

    # Inisialisasi hasil
    matched_rk = []
    unmatched_rk = []

    # Proses vouching menggunakan vectorization
    progress_bar = st.progress(0)
    total_rows = len(rk_df)
    progress_step = max(1, total_rows // 100)  # Update setiap 1%

    for i, (_, rk_row) in enumerate(rk_df.iterrows()):
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
        if i % progress_step == 0:
            progress_bar.progress(i / total_rows)

    return matched_rk, unmatched_rk

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
        sp2d_df['NoSP2D'] = sp2d_df['NoSP2D'].astype(str).str.strip()
        sp2d_df['TglSP2D'] = pd.to_datetime(sp2d_df['TglSP2D'])
        rk_df['Keterangan'] = rk_df['Keterangan'].astype(str).str.strip()
        rk_df['Tanggal'] = pd.to_datetime(rk_df['Tanggal'])

        # Tombol untuk memulai vouching
        if st.button("Mulai Vouching"):
            matched_rk, unmatched_rk = perform_vouching(rk_df, sp2d_df)

            # Tampilkan hasil
            st.subheader("Hasil Vouching")
            st.write("Rekening Koran Matched:")
            st.dataframe(matched_rk)
            st.write("Rekening Koran Unmatched:")
            st.dataframe(unmatched_rk)

            # Simpan hasil ke file Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                pd.DataFrame(matched_rk).to_excel(writer, sheet_name='Matched', index=False)
                pd.DataFrame(unmatched_rk).to_excel(writer, sheet_name='Unmatched', index=False)
            output.seek(0)

            st.download_button(
                label="Unduh Hasil Vouching",
                data=output,
                file_name="hasil_vouching.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
