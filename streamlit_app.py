import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# Fungsi untuk ekstraksi SP2D dengan regex yang lebih robust
def extract_sp2d_number(description):
    # Mencari 6 digit angka yang mungkin merupakan SP2D
    match = re.search(r'(?<!\d)\d{6}(?!\d)', str(description))
    return match.group(0) if match else None

# Fungsi utama untuk proses vouching dengan optimasi
@st.cache_data
def perform_vouching(rk_df, sp2d_df):
    # Preprocessing data
    rk_df = rk_df.copy()
    sp2d_df = sp2d_df.copy()
    
    # Normalisasi nama kolom
    rk_df.columns = rk_df.columns.str.strip().str.lower()
    sp2d_df.columns = sp2d_df.columns.str.strip().str.lower()
    
    # Ekstraksi nomor SP2D
    rk_df['nosp2d_6digits'] = rk_df['keterangan'].apply(extract_sp2d_number)
    sp2d_df['nosp2d_6digits'] = sp2d_df['nosp2d'].astype(str).str[:6]
    
    # Konversi tipe data
    numeric_cols = ['jumlah']
    for col in numeric_cols:
        rk_df[col] = pd.to_numeric(rk_df[col], errors='coerce')
        sp2d_df[col] = pd.to_numeric(sp2d_df[col], errors='coerce')
    
    # Handle missing values
    rk_df['nosp2d_6digits'] = rk_df['nosp2d_6digits'].fillna('')
    sp2d_df['nosp2d_6digits'] = sp2d_df['nosp2d_6digits'].fillna('')
    
    # Konversi kolom tanggal
    rk_df['tanggal'] = pd.to_datetime(rk_df['tanggal'], errors='coerce')
    sp2d_df['tglsp2d'] = pd.to_datetime(sp2d_df['tglsp2d'], errors='coerce')
    
    # Proses matching utama berdasarkan nomor SP2D
    merged = rk_df.merge(
        sp2d_df[['nosp2d_6digits', 'jumlah', 'tglsp2d', 'skpd', 'nosp2d']],
        left_on=['nosp2d_6digits', 'jumlah'],
        right_on=['nosp2d_6digits', 'jumlah'],
        how='left',
        suffixes=('', '_SP2D')
    )
    
    # Tambahkan status awal berdasarkan nomor SP2D
    merged['status'] = merged['nosp2d'].notna().map({True: 'Matched by SP2D', False: 'Unmatched'})
    
    # Pisahkan data unmatched untuk pencocokan alternatif
    unmatched_mask = merged['status'] == 'Unmatched'
    unmatched_rk = merged[unmatched_mask].copy()
    
    # Proses matching alternatif berdasarkan jumlah dan tanggal
    unmatched_rk = unmatched_rk.merge(
        sp2d_df[['jumlah', 'tglsp2d', 'skpd', 'nosp2d']],
        left_on=['jumlah', 'tanggal'],
        right_on=['jumlah', 'tglsp2d'],
        how='left',
        suffixes=('', '_alt')
    )
    
    # Update status untuk transaksi yang matched berdasarkan jumlah dan tanggal
    unmatched_rk['status'] = unmatched_rk['nosp2d_alt'].notna().map({True: 'Matched by Amount and Date', False: 'Unmatched'})
    
    # Gabungkan hasil matched dari kedua tahap
    matched_by_amount_date = unmatched_rk[unmatched_rk['status'] == 'Matched by Amount and Date']
    unmatched_final = unmatched_rk[unmatched_rk['status'] == 'Unmatched']
    
    # Gabungkan semua hasil matched
    matched_final = pd.concat([
        merged[merged['status'] == 'Matched by SP2D'],
        matched_by_amount_date
    ])
    
    # Identifikasi SP2D yang tidak terpakai
    used_sp2d_keys = set(matched_final['nosp2d']).union(set(matched_by_amount_date['nosp2d_alt']))
    unmatched_sp2d = sp2d_df[~sp2d_df['nosp2d'].isin(used_sp2d_keys)]
    
    # Gabungkan semua data RK dengan status
    all_rk = pd.concat([matched_final, unmatched_final])
    
    return all_rk, unmatched_sp2d

# Fungsi untuk membuat file Excel
def to_excel(df_list, sheet_names):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, sheet_name in zip(df_list, sheet_names):
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    processed_data = output.getvalue()
    return processed_data

# UI
st.title("Aplikasi Vouching SP2D vs Rekening Koran")

# Upload file
rk_file = st.file_uploader("Upload Rekening Koran", type="xlsx")
sp2d_file = st.file_uploader("Upload SP2D", type="xlsx")

if rk_file and sp2d_file:
    try:
        # Membaca file
        rk_df = pd.read_excel(rk_file)
        sp2d_df = pd.read_excel(sp2d_file)
        
        # Debugging: Tampilkan nama kolom
        st.write("Nama kolom Rekening Koran:", rk_df.columns.tolist())
        st.write("Nama kolom SP2D:", sp2d_df.columns.tolist())
        
        # Validasi kolom
        required_rk = {'tanggal', 'keterangan', 'jumlah'}
        required_sp2d = {'skpd', 'nosp2d', 'tglsp2d', 'jumlah'}
        
        if not required_rk.issubset(rk_df.columns.str.lower()):
            st.error(f"File Rekening Koran tidak memiliki kolom yang diperlukan: {required_rk}")
            st.stop()
            
        if not required_sp2d.issubset(sp2d_df.columns.str.lower()):
            st.error(f"File SP2D tidak memiliki kolom yang diperlukan: {required_sp2d}")
            st.stop()
        
        # Debugging: Tampilkan isi DataFrame
        st.write("Data Rekening Koran:")
        st.dataframe(rk_df)
        st.write("Data SP2D:")
        st.dataframe(sp2d_df)
        
        # Proses vouching
        with st.spinner('Memproses data...'):
            all_rk, unmatched_sp2d = perform_vouching(rk_df, sp2d_df)
        
        # Tampilkan statistik
        st.subheader("Statistik")
        cols = st.columns(3)
        cols[0].metric("RK Matched", len(all_rk[all_rk['status'].str.contains('Matched')]))
        cols[1].metric("RK Unmatched", len(all_rk[all_rk['status'] == 'Unmatched']))
        cols[2].metric("SP2D Unmatched", len(unmatched_sp2d))
        
        # Buat file Excel untuk di-download
        df_list = [all_rk, unmatched_sp2d]
        sheet_names = ['Rekening Koran (All)', 'SP2D Unmatched']
        excel_data = to_excel(df_list, sheet_names)
        
        # Tombol download
        st.download_button(
            label="Download Hasil Vouching",
            data=excel_data,
            file_name=f"vouching_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")
        st.stop()
