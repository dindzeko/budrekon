import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

def preprocess_jumlah(series):
    series = series.astype(str).str.replace(r'[^\d]', '', regex=True)
    return pd.to_numeric(series, errors='coerce')

def extract_sp2d_number(description):
    if pd.isna(description):
        return None
    matches = re.findall(r'\b\d{6}\b', str(description))
    return matches[0] if matches else None

def clean_skpd_name(name):
    if pd.isna(name):
        return None
    name = re.sub(r'\d+', '', str(name)).strip().upper()
    name = re.sub(r'(KECAMATAN|KELURAHAN|BADAN|DINAS)\s*', '', name, flags=re.IGNORECASE)
    return name.strip()

def extract_skpd_code(description):
    if pd.isna(description):
        return None
    parts = str(description).split('/')
    if len(parts) >= 6:
        return clean_skpd_name(parts[5])
    return None

@st.cache_data
def perform_vouching(rk_df, sp2d_df):
    # Debugging awal
    st.write("### Debugging Kolom SP2D")
    st.write("Kolom SP2D awal:", sp2d_df.columns.tolist())
    
    # Normalisasi kolom
    sp2d_df.columns = sp2d_df.columns.str.strip().str.lower()
    rk_df.columns = rk_df.columns.str.strip().str.lower()
    
    # Cari kolom SKPD dinamis
    skpd_cols = [col for col in sp2d_df.columns if 'skpd' in col]
    if not skpd_cols:
        raise ValueError(f"Kolom SKPD tidak ditemukan! Kolom tersedia: {list(sp2d_df.columns)}")
    
    # Rename kolom SKPD
    sp2d_df = sp2d_df.rename(columns={skpd_cols[0]: 'skpd'})
    st.write(f"Kolom '{skpd_cols[0]}' diubah menjadi 'skpd'")
    
    # Validasi kolom penting
    required_sp2d = {'nosp2d', 'tglsp2d', 'jumlah', 'skpd'}
    missing = required_sp2d - set(sp2d_df.columns)
    if missing:
        raise ValueError(f"Kolom SP2D kurang: {missing}")
    
    # Preprocessing data
    rk_df['jumlah'] = preprocess_jumlah(rk_df['jumlah'])
    sp2d_df['jumlah'] = preprocess_jumlah(sp2d_df['jumlah'])
    
    # Ekstraksi informasi
    rk_df['nosp2d_6digits'] = rk_df['keterangan'].apply(extract_sp2d_number)
    rk_df['skpd_code'] = rk_df['keterangan'].apply(extract_skpd_code)  # Kolom SKPD dari RK
    
    sp2d_df['nosp2d_6digits'] = sp2d_df['nosp2d'].astype(str).str[:6]
    sp2d_df['skpd_code'] = sp2d_df['skpd'].apply(clean_skpd_name)  # Kolom SKPD dari SP2D
    
    # Konversi tanggal
    rk_df['tanggal'] = pd.to_datetime(rk_df['tanggal'], errors='coerce')
    sp2d_df['tglsp2d'] = pd.to_datetime(sp2d_df['tglsp2d'], errors='coerce')
    
    # Primary matching
    rk_df['key'] = rk_df['nosp2d_6digits'].astype(str) + '_' + rk_df['jumlah'].astype(str)
    sp2d_df['key'] = sp2d_df['nosp2d_6digits'].astype(str) + '_' + sp2d_df['jumlah'].astype(str)
    
    merged = rk_df.merge(
        sp2d_df[['key', 'nosp2d', 'tglsp2d', 'skpd_code']],
        on='key',
        how='left',
        suffixes=('', '_sp2d')  # skpd_code_sp2d dari SP2D, skpd_code dari RK
    )
    
    # Perbaikan utama di sini
    merged['skpd'] = merged['skpd_code_sp2d'].combine_first(merged['skpd_code'])
    merged['status'] = merged['nosp2d'].notna().map({True: 'Matched', False: 'Unmatched'})
    
    # Secondary matching
    unmatched_rk = merged[merged['status'] == 'Unmatched']
    remaining_sp2d = sp2d_df[~sp2d_df['key'].isin(merged['key'])]
    
    if not unmatched_rk.empty and not remaining_sp2d.empty:
        secondary_merge = unmatched_rk.merge(
            remaining_sp2d,
            left_on=['jumlah', 'tanggal'],
            right_on=['jumlah', 'tglsp2d'],
            how='inner',
            suffixes=('', '_y')
        )
        
        if not secondary_merge.empty:
            merged.loc[secondary_merge.index, 'nosp2d'] = secondary_merge['nosp2d_y']
            merged.loc[secondary_merge.index, 'tglsp2d'] = secondary_merge['tglsp2d_y']
            merged.loc[secondary_merge.index, 'skpd'] = secondary_merge['skpd_code_y']
            merged.loc[secondary_merge.index, 'status'] = 'Matched (Secondary)'
    
    return merged, remaining_sp2d

# ... (bagian UI Streamlit tetap sama)
