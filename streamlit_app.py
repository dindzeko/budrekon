import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

def preprocess_jumlah(series):
    series = series.astype(str).str.replace(r'[.]', '', regex=True)
    series = series.str.replace(',', '.', regex=False)
    return pd.to_numeric(series, errors='coerce')

def extract_sp2d_number(description):
    match = re.search(r'(?<!\d)\d{6}(?!\d)', str(description))
    return match.group(0) if match else None

@st.cache_data
def perform_vouching(rk_df, sp2d_df):
    rk_df = rk_df.copy()
    sp2d_df = sp2d_df.copy()
    
    rk_df.columns = rk_df.columns.str.strip().str.lower()
    sp2d_df.columns = sp2d_df.columns.str.strip().str.lower()
    
    numeric_cols = ['jumlah']
    for col in numeric_cols:
        rk_df[col] = preprocess_jumlah(rk_df[col])
        sp2d_df[col] = preprocess_jumlah(sp2d_df[col])
    
    rk_df['nosp2d_6digits'] = rk_df['keterangan'].apply(extract_sp2d_number)
    sp2d_df['nosp2d_6digits'] = sp2d_df['nosp2d'].astype(str).str[:6]
    
    rk_df['tanggal'] = pd.to_datetime(rk_df['tanggal'], format='%Y-%m-%d', errors='coerce')
    sp2d_df['tglsp2d'] = pd.to_datetime(sp2d_df['tglsp2d'], format='%d/%m/%Y', errors='coerce')
    
    rk_df['key'] = rk_df['nosp2d_6digits'] + '_' + rk_df['jumlah'].astype(str)
    sp2d_df['key'] = sp2d_df['nosp2d_6digits'] + '_' + sp2d_df['jumlah'].astype(str)
    
    merged = rk_df.merge(sp2d_df[['key', 'nosp2d', 'tglsp2d', 'skpd']],
                          on='key', how='left', suffixes=('', '_SP2D'))
    merged['status'] = merged['nosp2d'].notna().map({True: 'Matched', False: 'Unmatched'})
    
    used_sp2d = set(merged.loc[merged['status'] == 'Matched', 'key'])
    unmatched_sp2d = sp2d_df[~sp2d_df['key'].isin(used_sp2d)]
    unmatched_rk = merged[merged['status'] == 'Unmatched'].copy()
    
    if not unmatched_rk.empty and not unmatched_sp2d.empty:
        unmatched_rk = unmatched_rk.sort_values('tanggal')
        unmatched_sp2d = unmatched_sp2d.sort_values('tglsp2d')
        
        second_merge = pd.merge_asof(
            unmatched_rk,
            unmatched_sp2d,
            left_on='tanggal',
            right_on='tglsp2d',
            by='jumlah',
            direction='nearest'
        )
        
        if not second_merge.empty:
            merged.loc[second_merge.index, ['nosp2d', 'tglsp2d', 'skpd', 'status']] = \
                second_merge[['nosp2d', 'tglsp2d', 'skpd']].assign(status='Matched (Secondary)')
            
            used_sp2d.update(second_merge['key'])
            unmatched_sp2d = sp2d_df[~sp2d_df['key'].isin(used_sp2d)]
    
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
        
        required_rk = {'tanggal', 'keterangan', 'jumlah'}
        required_sp2d = {'skpd', 'nosp2d', 'tglsp2d', 'jumlah'}
        
        if not required_rk.issubset(rk_df.columns.str.lower()):
            st.error(f"Kolom Rekening Koran tidak valid! Harus ada: {required_rk}")
            st.stop()
        
        if not required_sp2d.issubset(sp2d_df.columns.str.lower()):
            st.error(f"Kolom SP2D tidak valid! Harus ada: {required_sp2d}")
            st.stop()
        
        with st.spinner('Memproses data...'):
            merged_rk, unmatched_sp2d = perform_vouching(rk_df, sp2d_df)
        
        st.subheader("Statistik")
        cols = st.columns(4)
        cols[0].metric("Total RK", len(merged_rk))
        cols[1].metric("Matched (Primary)", len(merged_rk[merged_rk['status'] == 'Matched']))
        cols[2].metric("Matched (Secondary)", len(merged_rk[merged_rk['status'] == 'Matched (Secondary)']))
        cols[3].metric("Unmatched SP2D", len(unmatched_sp2d))
        
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
