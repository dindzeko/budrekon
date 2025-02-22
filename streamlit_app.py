import streamlit as st
import pandas as pd
import re
from datetime import timedelta
from io import BytesIO

# Fungsi untuk mengekstrak 6 digit pertama Nomor SP2D dari kolom Keterangan
def extract_sp2d_number(description):
    match = re.search(r'\b\d{6}\b', str(description))  # Pastikan description adalah string
    return match.group(0) if match else None

# Fungsi untuk melakukan vouching utama
def perform_vouching(rk_df, sp2d_df):
    matched_rk = []
    unmatched_rk = []
    matched_sp2d = []
    unmatched_sp2d = []

    # Membuat dictionary untuk SP2D agar pencarian lebih cepat
    sp2d_dict = {str(row['NoSP2D'])[:6]: row for _, row in sp2d_df.iterrows()}

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
        no_sp2d = str(sp2d_row.get('NoSP2D', ''))[:6]
        if no_sp2d not in sp2d_numbers_in_rk:
            unmatched_sp2d.append(sp2d_row)

    return matched_rk, unmatched_rk, matched_sp2d, unmatched_sp2d

# Fungsi untuk mencari alternatif berdasarkan jumlah, tanggal, dan SKPD
def find_alternative_matches(unmatched_rk, sp2d_df):
    alternative_matches = []
    for rk_row in unmatched_rk:
        rk_amount = rk_row['Jumlah (Rp)']
        rk_date = pd.to_datetime(rk_row['Tanggal'])
        rk_description = str(rk_row['Keterangan']).lower()

        # Cari SP2D yang sesuai dengan kriteria alternatif
        for _, sp2d_row in sp2d_df.iterrows():
            sp2d_amount = sp2d_row['Jumlah']
            sp2d_date = pd.to_datetime(sp2d_row['TglSP2D'])
            sp2d_skpd = str(sp2d_row['SKPD']).lower()

            # Kriteria alternatif: jumlah sama, tanggal selisih <= 1 hari, dan SKPD ada di uraian RK
            if (rk_amount == sp2d_amount and
                abs((rk_date - sp2d_date).days) <= 1 and
                sp2d_skpd in rk_description):
                alternative_matches.append({
                    'Tanggal': rk_row['Tanggal'],
                    'Keterangan': rk_row['Keterangan'],
                    'Jumlah (Rp)': rk_row['Jumlah (Rp)'],
                    'NO SP2D': sp2d_row['NoSP2D'],
                    'Status': 'Alternative Match'
                })
                break

    return alternative_matches

# Fungsi untuk menambahkan hasil vouching ke file RK
def add_vouching_to_rk(rk_df, matched_rk, unmatched_rk, alternative_matches):
    rk_df['Keterangan_Vouching'] = "Unmatched"  # Default: Unmatched

    # Isi kolom Keterangan_Vouching untuk transaksi matched
    for row in matched_rk:
        mask = (rk_df['Keterangan'].str.contains(str(row['NO SP2D']), na=False)) & \
               (rk_df['Jumlah'] == row['Jumlah (Rp)'])
        rk_df.loc[mask, 'Keterangan_Vouching'] = row['NO SP2D']

    # Isi kolom Keterangan_Vouching untuk transaksi alternatif
    for row in alternative_matches:
        mask = (rk_df['Keterangan'].str.contains(str(row['NO SP2D']), na=False)) & \
               (rk_df['Jumlah'] == row['Jumlah (Rp)'])
        rk_df.loc[mask, 'Keterangan_Vouching'] = f"Alternative Match: {row['NO SP2D']}"

    return rk_df

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
            matched_rk, unmatched_rk, matched_sp2d, unmatched_sp2d = perform_vouching(rk_df, sp2d_df)

            # Cari alternatif untuk transaksi unmatched
            alternative_matches = find_alternative_matches(unmatched_rk, sp2d_df)

            # Tampilkan hasil
            st.subheader("Hasil Vouching")
            st.write("Rekening Koran Matched:")
            st.dataframe(matched_rk)
            st.write("Rekening Koran Unmatched:")
            st.dataframe(unmatched_rk)
            st.write("Rekening Koran Alternative Matches:")
            st.dataframe(alternative_matches)

            # Tambahkan hasil vouching ke file RK
            rk_df = add_vouching_to_rk(rk_df, matched_rk, unmatched_rk, alternative_matches)

            # Simpan hasil ke file Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                rk_df.to_excel(writer, sheet_name='Rekening Koran', index=False)
                sp2d_df.to_excel(writer, sheet_name='SP2D', index=False)
            output.seek(0)

            st.download_button(
                label="Unduh Hasil Vouching",
                data=output,
                file_name="hasil_vouching.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
