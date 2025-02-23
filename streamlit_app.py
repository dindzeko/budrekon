import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# Fungsi untuk membersihkan format angka dengan separator
def preprocess_jumlah(series):
    """Fungsi untuk membersihkan format angka dengan separator"""
    series = series.astype(str)
    series = series.str.replace(r'[.]', '', regex=True)  # Hapus separator ribuan
    series = series.str.replace(',', '.', regex=False)   # Ganti desimal koma dengan titik
    return pd.to_numeric(series, errors='coerce')

# Fungsi untuk mengekstrak nomor SP2D (6 digit pertama)
def extract_sp2d_number(description):
    match = re.search(r'(?<!\d)\d{6}(?!\d)', str(description))
    return match.group(0) if match else None

# Fungsi utama untuk melakukan vouching
@st.cache_data
def perform_vouching(rk_df, sp2d_df):
    # Preprocessing data
    rk_df = rk_df.copy()
    sp2d_df = sp2d_df.copy()
    
    # Normalisasi kolom
    rk_df.columns = rk_df.columns.str.strip().str.lower()
    sp2d_df.columns = sp2d_df.columns.str.strip().str.lower()
    
    # Preprocessing jumlah
    numeric_cols = ['jumlah']
    for col in numeric_cols:
        rk_df[col] = preprocess_jumlah(rk_df[col])
        sp2d_df[col] = preprocess_jumlah(sp2d_df[col])
    
    # Ekstraksi SP2D
    rk_df['nosp2d_6digits'] = rk_df['keterangan'].apply(extract_sp2d_number)
    sp2d_df['nosp2d_6digits'] = sp2d_df['nosp2d'].astype(str).str[:6]
    
    # Konversi tanggal
    rk_df['tanggal'] = pd.to_datetime(rk_df['tanggal'], errors='coerce')
    sp2d_df['tglsp2d'] = pd.to_datetime(sp2d_df['tglsp2d'], errors='coerce')
    
    # Membuat kunci
    rk_df['key'] = rk_df['nosp2d_6digits'] + '_' + rk_df['jumlah'].astype(str)
    sp2d_df['key'] = sp2d_df['nosp2d_6digits'] + '_' + sp2d_df['jumlah'].astype(str)
    
    # Vouching pertama (kunci SP2D + jumlah)
    merged = rk_df.merge(
        sp2d_df[['key', 'nosp2d', 'tglsp2d', 'skpd']],
        on='key',
        how='left',
        suffixes=('', '_SP2D')
    )
    merged['status'] = merged['nosp2d'].notna().map({True: 'Matched', False: 'Unmatched'})
    
    # Identifikasi data belum terhubung
    used_sp2d = set(merged.loc[merged['status'] == 'Matched', 'key'])
    unmatched_sp2d = sp2d_df[~sp2d_df['key'].isin(used_sp2d)]
    unmatched_rk = merged[merged['status'] == 'Unmatched'].copy()
    
    # Vouching kedua (jumlah + tanggal)
    if not unmatched_rk.empty and not unmatched_sp2d.empty:
        second_merge = unmatched_rk.merge(
            unmatched_sp2d,
            left_on=['jumlah', 'tanggal'],
            right_on=['jumlah', 'tglsp2d'],
            how='inner',
            suffixes=('', '_y')
        )
        
        if not second_merge.empty:
            # Update data hasil merge kedua
            merged.loc[second_merge.index, 'nosp2d'] = second_merge['nosp2d_y']
            merged.loc[second_merge.index, 'tglsp2d'] = second_merge['tglsp2d_y']
            merged.loc[second_merge.index, 'skpd'] = second_merge['skpd_y']
            merged.loc[second_merge.index, 'status'] = 'Matched (Secondary)'
            
            # Update daftar SP2D yang digunakan
            used_sp2d.update(second_merge['key_y'])
            unmatched_sp2d = sp2d_df[~sp2d_df['key'].isin(used_sp2d)]
    
    return merged, unmatched_sp2d

# Fungsi untuk menyimpan DataFrame ke file Excel
def to_excel(df_list, sheet_names):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, sheet_name in zip(df_list, sheet_names):
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# Streamlit App
st.title("Aplikasi Vouching SP2D vs Rekening Koran (Enhanced)")

# Upload file
rk_file = st.file_uploader("Upload Rekening Koran", type="xlsx")
sp2d_file = st.file_uploader("Upload SP2D", type="xlsx")

if rk_file and sp2d_file:
    try:
        # Baca file Excel
        rk_df = pd.read_excel(rk_file)
        sp2d_df = pd.read_excel(sp2d_file)
        
        # Validasi kolom
        required_rk = {'tanggal', 'keterangan', 'jumlah'}
        required_sp2d = {'skpd', 'nosp2d', 'tglsp2d', 'jumlah'}
        
        if not required_rk.issubset(rk_df.columns.str.lower()):
            st.error(f"Kolom Rekening Koran tidak valid! Harus ada: {required_rk}")
            st.stop()
            
        if not required_sp2d.issubset(sp2d_df.columns.str.lower()):
            st.error(f"Kolom SP2D tidak valid! Harus ada: {required_sp2d}")
            st.stop()
        
        # Proses vouching
        with st.spinner('Memproses data...'):
            merged_rk, unmatched_sp2d = perform_vouching(rk_df, sp2d_df)
        
        # Statistik
        st.subheader("Statistik")
        cols = st.columns(4)
        cols[0].metric("Total RK", len(merged_rk))
        cols[1].metric("Matched (Primary)", len(merged_rk[merged_rk['status'] == 'Matched']))
        cols[2].metric("Matched (Secondary)", len(merged_rk[merged_rk['status'] == 'Matched (Secondary)']))
        cols[3].metric("Unmatched SP2D", len(unmatched_sp2d))
        
        # Download hasil
        df_list = [merged_rk, unmatched_sp2d]
        sheet_names = ['Hasil Vouching', 'SP2D Belum Terpakai']
        excel_data = to_excel(df_list, sheet_names)
        
        st.download_button(
            label="Download Hasil",
            data=excel_data,
            file_name=f"vouching_enhanced_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error: {str(e)}")
        st.stop()
