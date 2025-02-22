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
    
    # Ekstraksi nomor SP2D
    rk_df['NoSP2D_6digits'] = rk_df['Keterangan'].apply(extract_sp2d_number)
    sp2d_df['NoSP2D_6digits'] = sp2d_df['NoSP2D'].astype(str).str[:6]
    
    # Konversi tipe data
    numeric_cols = ['Jumlah']
    for col in numeric_cols:
        rk_df[col] = pd.to_numeric(rk_df[col], errors='coerce')
        sp2d_df[col] = pd.to_numeric(sp2d_df[col], errors='coerce')
    
    # Handle missing values
    rk_df['NoSP2D_6digits'] = rk_df['NoSP2D_6digits'].fillna('')
    sp2d_df['NoSP2D_6digits'] = sp2d_df['NoSP2D_6digits'].fillna('')
    
    # Membuat kunci unik
    rk_df['key'] = rk_df['NoSP2D_6digits'] + '_' + rk_df['Jumlah'].astype(str)
    sp2d_df['key'] = sp2d_df['NoSP2D_6digits'] + '_' + sp2d_df['Jumlah'].astype(str)
    
    # Proses matching
    merged = rk_df.merge(
        sp2d_df[['key', 'NoSP2D', 'TglSP2D', 'SKPD']],
        on='key',
        how='left',
        suffixes=('', '_SP2D')
    )
    
    # Klasifikasi data
    matched_mask = merged['NoSP2D'].notna()
    matched_rk = merged[matched_mask]
    unmatched_rk = merged[~matched_mask]
    
    # Identifikasi SP2D yang tidak terpakai
    used_sp2d_keys = set(matched_rk['key'])
    unmatched_sp2d = sp2d_df[~sp2d_df['key'].isin(used_sp2d_keys)]
    
    return matched_rk, unmatched_rk, sp2d_df, unmatched_sp2d

# Fungsi untuk pencarian alternatif dengan vectorization
def find_alternative_matches(unmatched_rk, sp2d_df):
    # Konversi ke datetime
    unmatched_rk['Tanggal'] = pd.to_datetime(unmatched_rk['Tanggal'], errors='coerce')
    sp2d_df['TglSP2D'] = pd.to_datetime(sp2d_df['TglSP2D'], errors='coerce')
    
    # Merge berdasarkan jumlah
    merged = unmatched_rk.merge(
        sp2d_df,
        on='Jumlah',
        suffixes=('_RK', '_SP2D')
    )
    
    # Filter berdasarkan tanggal dan SKPD
    merged['date_diff'] = (merged['Tanggal'] - merged['TglSP2D']).abs().dt.days
    merged['skpd_match'] = merged.apply(
        lambda x: str(x['SKPD']).lower() in str(x['Keterangan']).lower(),
        axis=1
    )
    
    # Kriteria alternatif
    alt_matches = merged[
        (merged['date_diff'] <= 1) &
        (merged['skpd_match'])
    ].drop_duplicates(subset=['Jumlah', 'Tanggal'], keep='first')
    
    return alt_matches

# Fungsi untuk membuat output Excel
def create_output(matched_rk, unmatched_rk, matched_sp2d, unmatched_sp2d, alt_matches):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # RK sheets
        matched_rk[['Tanggal', 'Keterangan', 'Jumlah', 'NoSP2D']].to_excel(
            writer, sheet_name='RK Matched', index=False)
        
        unmatched_rk[['Tanggal', 'Keterangan', 'Jumlah']].to_excel(
            writer, sheet_name='RK Unmatched', index=False)
        
        # SP2D sheets
        matched_sp2d.to_excel(writer, sheet_name='SP2D Matched', index=False)
        unmatched_sp2d.to_excel(writer, sheet_name='SP2D Unmatched', index=False)
        
        # Alternative matches
        if not alt_matches.empty:
            alt_matches[['Tanggal', 'Keterangan', 'Jumlah', 'NoSP2D']].to_excel(
                writer, sheet_name='Alternative Matches', index=False)
    
    output.seek(0)
    return output

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
        
        # Validasi kolom
        required_rk = {'Tanggal', 'Keterangan', 'Jumlah'}
        required_sp2d = {'SKPD', 'NoSP2D', 'TglSP2D', 'Jumlah'}
        
        if not required_rk.issubset(rk_df.columns):
            st.error(f"Kolom RK harus mengandung: {required_rk}")
            st.stop()
            
        if not required_sp2d.issubset(sp2d_df.columns):
            st.error(f"Kolom SP2D harus mengandung: {required_sp2d}")
            st.stop()
        
        # Proses vouching
        with st.spinner('Memproses data...'):
            matched_rk, unmatched_rk, sp2d_full, unmatched_sp2d = perform_vouching(rk_df, sp2d_df)
            alt_matches = find_alternative_matches(unmatched_rk, sp2d_df)
        
        # Tampilkan statistik
        st.subheader("Statistik")
        cols = st.columns(3)
        cols[0].metric("RK Matched", len(matched_rk))
        cols[1].metric("RK Unmatched", len(unmatched_rk))
        cols[2].metric("Alternative Matches", len(alt_matches))
        
        # Download hasil
        output = create_output(matched_rk, unmatched_rk, 
                             sp2d_full[sp2d_full['key'].isin(matched_rk['key'])], 
                             unmatched_sp2d, alt_matches)
        
        st.download_button(
            label="Unduh Hasil Lengkap",
            data=output,
            file_name=f"vouching_result_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")
        st.stop()
