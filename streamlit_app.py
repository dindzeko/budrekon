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
    # Debugging kolom SP2D
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
    
    # Ekstraksi informasi dari RK
    rk_df['nosp2d_6digits'] = rk_df['keterangan'].apply(extract_sp2d_number)
    rk_df['skpd_code'] = rk_df['keterangan'].apply(extract_skpd_code)  # SKPD dari keterangan RK
    
    # Ekstraksi informasi dari SP2D
    sp2d_df['nosp2d_6digits'] = sp2d_df['nosp2d'].astype(str).str[:6]
    sp2d_df['skpd_code'] = sp2d_df['skpd'].apply(clean_skpd_name)  # SKPD dari data SP2D
    
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
    
    # Perbaikan utama: Gabungkan SKPD dari SP2D dan RK
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

def to_excel(df_list, sheet_names):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, sheet in zip(df_list, sheet_names):
            df.to_excel(writer, index=False, sheet_name=sheet)
    return output.getvalue()

# UI Streamlit
st.title("🔄 Aplikasi Vouching SP2D (Final Fix)")

col1, col2 = st.columns(2)
rk_file = col1.file_uploader("Upload Rekening Koran", type=["xlsx", "xls"])
sp2d_file = col2.file_uploader("Upload Data SP2D", type=["xlsx", "xls"])

if rk_file and sp2d_file:
    try:
        # Load data
        rk_df = pd.read_excel(rk_file)
        sp2d_df = pd.read_excel(sp2d_file)
        
        # Tampilkan preview data
        with st.expander("🔍 Data Mentah"):
            st.write("Rekening Koran (5 baris):")
            st.write(rk_df.head())
            st.write("Data SP2D (5 baris):")
            st.write(sp2d_df.head())
        
        # Validasi kolom wajib
        required_rk = {'tanggal', 'keterangan', 'jumlah'}
        missing_rk = required_rk - set(rk_df.columns.str.lower())
        if missing_rk:
            st.error(f"❌ Kolom RK kurang: {missing_rk}")
            st.stop()
        
        # Proses vouching
        with st.spinner('🔍 Memproses data...'):
            result_df, unmatched_sp2d = perform_vouching(rk_df, sp2d_df)
        
        # Tampilkan hasil
        st.subheader("📊 Ringkasan Hasil")
        cols = st.columns(4)
        cols[0].metric("Total Transaksi RK", len(result_df))
        cols[1].metric("Terekoniliasi (Primer)", 
                      len(result_df[result_df['status'] == 'Matched']))
        cols[2].metric("Terekoniliasi (Sekunder)", 
                      len(result_df[result_df['status'] == 'Matched (Secondary)']))
        cols[3].metric("SP2D Tidak Terpakai", len(unmatched_sp2d))
        
        # Download hasil
        excel_data = to_excel(
            [result_df, unmatched_sp2d],
            ['Hasil Vouching', 'SP2D Unmatched']
        )
        
        st.download_button(
            label="📥 Download Hasil",
            data=excel_data,
            file_name=f"vouching_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.write("Detail Error:")
        st.write(f"- Tipe: {type(e).__name__}")
        st.write(f"- Pesan: {str(e)}")
        st.stop()
