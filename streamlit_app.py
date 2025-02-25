import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

def preprocess_jumlah(series):
    """Membersihkan format angka dengan separator"""
    series = series.astype(str)
    series = series.str.replace(r'[.,]', '', regex=True)
    return pd.to_numeric(series, errors='coerce')

def extract_sp2d_number(description):
    """Ekstrak 6 digit pertama nomor SP2D dari keterangan"""
    matches = re.findall(r'\b\d{6}\b', str(description))
    return matches[0] if matches else None

def extract_skpd_code(description):
    """Ekstrak kode SKPD dari keterangan RK"""
    parts = str(description).split('/')
    if len(parts) >= 6:
        return parts[5].strip().upper()
    return None

def clean_skpd_name(name):
    """Bersihkan nama SKPD dari angka dan prefix"""
    name = re.sub(r'\d+\s*', '', str(name)).strip().upper()
    name = re.sub(r'(KECAMATAN|KELURAHAN|Badan|Dinas)\s*', '', name)
    return name.strip()

@st.cache_data
def perform_vouching(rk_df, sp2d_df):
    # Normalisasi data
    rk_df = rk_df.copy()
    sp2d_df = sp2d_df.copy()
    
    # Normalisasi kolom
    rk_df.columns = rk_df.columns.str.strip().str.lower()
    sp2d_df.columns = sp2d_df.columns.str.strip().str.lower()
    
    # Preprocessing jumlah
    rk_df['jumlah'] = preprocess_jumlah(rk_df['jumlah'])
    sp2d_df['jumlah'] = preprocess_jumlah(sp2d_df['jumlah'])
    
    # Ekstraksi data
    rk_df['nosp2d_6digits'] = rk_df['keterangan'].apply(extract_sp2d_number)
    rk_df['skpd_code'] = rk_df['keterangan'].apply(extract_skpd_code)
    
    sp2d_df['nosp2d_6digits'] = sp2d_df['nosp2d'].astype(str).str[:6]
    sp2d_df['skpd_code'] = sp2d_df['skpd'].apply(clean_skpd_name)
    
    # Konversi tanggal
    rk_df['tanggal'] = pd.to_datetime(rk_df['tanggal'], errors='coerce')
    sp2d_df['tglsp2d'] = pd.to_datetime(sp2d_df['tglsp2d'], errors='coerce')
    
    # Primary Matching: SP2D + Jumlah
    rk_df['key'] = rk_df['nosp2d_6digits'] + '_' + rk_df['jumlah'].astype(str)
    sp2d_df['key'] = sp2d_df['nosp2d_6digits'] + '_' + sp2d_df['jumlah'].astype(str)
    
    merged = rk_df.merge(
        sp2d_df[['key', 'nosp2d', 'tglsp2d', 'skpd', 'skpd_code']],
        on='key',
        how='left',
        suffixes=('', '_sp2d')
    )
    
    # Update SKPD dengan data dari SP2D jika ada
    merged['skpd'] = merged['skpd_sp2d'].combine_first(merged['skpd'])
    merged['status'] = merged['nosp2d'].notna().map({True: 'Matched', False: 'Unmatched'})
    
    # Secondary Matching: Jumlah + Tanggal + SKPD
    unmatched_rk = merged[merged['status'] == 'Unmatched'].copy()
    remaining_sp2d = sp2d_df[~sp2d_df['key'].isin(merged[merged['status'] == 'Matched']['key'])]
    
    if not unmatched_rk.empty and not remaining_sp2d.empty:
        secondary_merge = unmatched_rk.merge(
            remaining_sp2d,
            left_on=['jumlah', 'tanggal', 'skpd_code'],
            right_on=['jumlah', 'tglsp2d', 'skpd_code'],
            how='inner',
            suffixes=('', '_y')
        )
        
        # Perbaikan: Gunakan kolom 'skpd_y' dari hasil merge
        if not secondary_merge.empty:
            merged.loc[secondary_merge.index, 'nosp2d'] = secondary_merge['nosp2d_y']
            merged.loc[secondary_merge.index, 'tglsp2d'] = secondary_merge['tglsp2d_y']
            merged.loc[secondary_merge.index, 'skpd'] = secondary_merge['skpd_y']  # Diperbaiki
            merged.loc[secondary_merge.index, 'status'] = 'Matched (Secondary)'
    
    return merged, remaining_sp2d

def to_excel(df_list, sheet_names):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, sheet_name in zip(df_list, sheet_names):
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# UI Streamlit
st.title("🔄 Aplikasi Vouching SP2D - Rekening Koran (Enhanced)")

# Upload file
col1, col2 = st.columns(2)
with col1:
    rk_file = st.file_uploader("Upload Rekening Koran", type="xlsx")
with col2:
    sp2d_file = st.file_uploader("Upload Data SP2D", type="xlsx")

if rk_file and sp2d_file:
    try:
        # Load data
        rk_df = pd.read_excel(rk_file)
        sp2d_df = pd.read_excel(sp2d_file)
        
        # Validasi kolom
        required_rk = {'tanggal', 'keterangan', 'jumlah'}
        required_sp2d = {'nosp2d', 'tglsp2d', 'jumlah', 'skpd'}
        
        if not required_rk.issubset(rk_df.columns.str.lower()):
            st.error(f"Kolom RK harus mengandung: {required_rk}")
            st.stop()
            
        if not required_sp2d.issubset(sp2d_df.columns.str.lower()):
            st.error(f"Kolom SP2D harus mengandung: {required_sp2d}")
            st.stop()
            
        # Proses vouching
        with st.spinner('🔍 Memproses data...'):
            result_df, unmatched_sp2d = perform_vouching(rk_df, sp2d_df)
        
        # Tampilkan statistik
        st.subheader("📊 Hasil Vouching")
        cols = st.columns(4)
        cols[0].metric("Total Transaksi", len(result_df))
        cols[1].metric("Terekoniliasi (Primer)", 
                      len(result_df[result_df['status'] == 'Matched']),
                      help="Match berdasarkan SP2D + Jumlah")
        cols[2].metric("Terekoniliasi (Sekunder)", 
                      len(result_df[result_df['status'] == 'Matched (Secondary)']),
                      help="Match berdasarkan Jumlah + Tanggal + SKPD")
        cols[3].metric("SP2D Belum Terpakai", 
                      len(unmatched_sp2d),
                      help="SP2D yang tidak memiliki transaksi terkait")
        
        # Tampilkan preview
        with st.expander("🔎 Lihat Detail Hasil"):
            st.dataframe(result_df)
        
        # Download hasil
        excel_data = to_excel(
            [result_df, unmatched_sp2d],
            ['Hasil Vouching', 'SP2D Belum Terpakai']
        )
        
        st.download_button(
            label="📥 Download Hasil Lengkap",
            data=excel_data,
            file_name=f"vouching_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.stop()
