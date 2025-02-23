import streamlit as st
import pandas as pd
from io import BytesIO

# Fungsi vouching (contoh sederhana, perlu dilengkapi sesuai logika Anda)
def vouching(rk_df, sp2d_df):
    # Pastikan nama kolom sesuai (lowercase untuk konsistensi)
    rk_df.columns = rk_df.columns.str.lower()
    sp2d_df.columns = sp2d_df.columns.str.lower()

    # Rename kolom 'tglsp2d' menjadi 'tanggal' di sp2d_df agar sesuai dengan rk_df
    sp2d_df = sp2d_df.rename(columns={'tglsp2d': 'tanggal'})

    # Merge berdasarkan jumlah dan tanggal
    merged = pd.merge(rk_df, sp2d_df, on=['jumlah', 'tanggal'], how='left', suffixes=('_rk', '_sp2d'))
    # Tambahkan kolom status
    merged['status'] = merged['nosp2d'].apply(lambda x: 'Matched' if pd.notnull(x) else 'Unmatched')

    unmatched_sp2d = sp2d_df[~sp2d_df['jumlah'].isin(merged['jumlah'])]

    return merged, unmatched_sp2d


def to_excel(df_list, sheet_names):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, sheet_name in zip(df_list, sheet_names):
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

st.title("Aplikasi Vouching SP2D vs Rekening Koran (Enhanced)")

rk_file = st.file_uploader("Upload Rekening Koran", type="xlsx")
sp2d_file = st.file_uploader("Upload SP2D", type="xlsx")

if rk_file and sp2d_file:
    try:
        rk_df = pd.read_excel(rk_file)
        sp2d_df = pd.read_excel(sp2d_file)

        # Validasi kolom
        required_rk = {'tanggal', 'keterangan', 'jumlah'}
        required_sp2d = {'skpd', 'nosp2d', 'tanggal', 'jumlah'}  # Ubah tglsp2d menjadi tanggal

        # Konversi nama kolom menjadi lowercase sebelum validasi
        rk_df.columns = rk_df.columns.str.lower()
        sp2d_df.columns = sp2d_df.columns.str.lower()

        if not required_rk.issubset(rk_df.columns):
            st.error(f"Kolom Rekening Koran tidak valid! Harus ada: {required_rk}")
            st.stop()

        # Periksa apakah kolom 'tglsp2d' ada di sp2d_df, jika ada ubah namanya menjadi 'tanggal'
        if 'tglsp2d' in sp2d_df.columns:
            sp2d_df = sp2d_df.rename(columns={'tglsp2d': 'tanggal'})

        if not required_sp2d.issubset(sp2d_df.columns):
            st.error(f"Kolom SP2D tidak valid! Harus ada: {required_sp2d}")
            st.stop()

        # Proses vouching (pencocokan) data
        merged_df, unmatched_sp2d_df = vouching(rk_df, sp2d_df) # Panggil fungsi vouching

        # Membuat file excel hasil vouching
        excel_file = to_excel([merged_df, unmatched_sp2d_df], ['Matched Data', 'Unmatched SP2D'])

        st.download_button(label='Download Hasil Vouching', data=excel_file, file_name='hasil_vouching.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
