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
    if pd.isna(description):
        return None
    matches = re.findall(r'\b\d{6}\b', str(description))
    return matches[0] if matches else None

def clean_skpd_name(name):
    """Bersihkan nama SKPD dari angka dan prefix"""
    if pd.isna(name):
        return None
    name = re.sub(r'\d+\s*', '', str(name)).strip().upper()
    name = re.sub(r'(?:KECAMATAN|KELURAHAN|BADAN|DINAS)\s*', '', name, flags=re.IGNORECASE)
    return name.strip()

def extract_skpd_code(description):
    """Ekstrak dan bersihkan nama SKPD dari keterangan RK"""
    if pd.isna(description):
        return None
    parts = str(description).split('/')
    if len(parts) >= 6:
        skpd_part = parts[5].strip().upper()
        return clean_skpd_name(skpd_part)
    return None

@st.cache_data
def perform_vouching(rk_df, sp2d_df):
    # Normalisasi data
    rk_df = rk_df.copy()
    sp2d_df = sp2d_df.copy()
    
    # Tampilkan kolom untuk debugging
    st.write("### Debugging Kolom SP2D")
    st.write("Kolom SP2D sebelum normalisasi:", sp2d_df.columns.tolist())
    
    # Normalisasi nama kolom
    rk_df.columns = rk_df.columns.str.strip().str.lower()
    sp2d_df.columns = sp2d_df.columns.str.strip().str.lower()
    
    st.write("Kolom SP2D setelah normalisasi:", sp2d_df.columns.tolist())
    
    # Validasi kolom penting
    if 'skpd' not in sp2d_df.columns:
        raise ValueError(f"""
        🔴 Kolom 'skpd' tidak ditemukan di data SP2D.
        Kolom yang tersedia: {list(sp2d_df.columns)}
        """)
    
    # Preprocessing jumlah
    try:
        rk_df['jumlah'] = preprocess_jumlah(rk_df['jumlah'])
        sp2d_df['jumlah'] = preprocess_jumlah(sp2d_df['jumlah'])
    except Exception as e:
        st.error(f"❌ Error saat memproses kolom jumlah: {str(e)}")
        st.write("Contoh data jumlah RK:", rk_df['jumlah'].head())
        st.write("Contoh data jumlah SP2D:", sp2d_df['jumlah'].head())
        raise
    
    # Ekstraksi informasi
    try:
        rk_df['nosp2d_6digits'] = rk_df['keterangan'].apply(extract_sp2d_number)
        rk_df['skpd_code'] = rk_df['keterangan'].apply(extract_skpd_code)
    except Exception as e:
        st.error(f"❌ Error saat mengekstrak data RK: {str(e)}")
        st.write("Contoh keterangan RK:", rk_df['keterangan'].head().tolist())
        raise
    
    try:
        sp2d_df['nosp2d_6digits'] = sp2d_df['nosp2d'].astype(str).str[:6]
        sp2d_df['skpd_code'] = sp2d_df['skpd'].apply(clean_skpd_name)
    except Exception as e:
        st.error(f"❌ Error saat memproses data SP2D: {str(e)}")
        st.write("Contoh data SP2D:", sp2d_df.head())
        raise
    
    # Konversi tanggal
    rk_df['tanggal'] = pd.to_datetime(rk_df['tanggal'], errors='coerce')
    sp2d_df['tglsp2d'] = pd.to_datetime(sp2d_df['tglsp2d'], errors='coerce')
    
    # Primary Matching
    rk_df['key'] = rk_df['nosp2d_6digits'].astype(str) + '_' + rk_df['jumlah'].astype(str)
    sp2d_df['key'] = sp2d_df['nosp2d_6digits'].astype(str) + '_' + sp2d_df['jumlah'].astype(str)
    
    merged = rk_df.merge(
        sp2d_df[['key', 'nosp2d', 'tglsp2d', 'skpd_code']],
        on='key',
        how='left',
        suffixes=('', '_sp2d')
    )
    
    # Update SKPD
    merged['skpd'] = merged['skpd_code_sp2d'].combine_first(merged['skpd'])
    merged['status'] = merged['nosp2d'].notna().map({True: 'Matched', False: 'Unmatched'})
    
    # Secondary Matching
    unmatched_rk = merged[merged['status'] == 'Unmatched'].copy()
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
        for df, sheet_name in zip(df_list, sheet_names):
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# UI Streamlit
st.title("🔄 Aplikasi Vouching SP2D - Rekening Koran (Debug Mode)")
st.warning("⚠️ Mode debugging aktif - Beberapa data contoh akan ditampilkan")

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
        
        # Tampilkan preview data untuk debugging
        with st.expander("🔍 Data Mentah RK (5 baris pertama)"):
            st.write(rk_df.head())
            
        with st.expander("🔍 Data Mentah SP2D (5 baris pertama)"):
            st.write(sp2d_df.head())
        
        # Validasi kolom wajib
        required_rk = {'tanggal', 'keterangan', 'jumlah'}
        required_sp2d = {'nosp2d', 'tglsp2d', 'jumlah', 'skpd'}
        
        # Cek kolom RK
        missing_rk = required_rk - set(rk_df.columns.str.lower())
        if missing_rk:
            st.error(f"❌ Kolom RK tidak lengkap: {missing_rk}")
            st.write("Kolom yang tersedia:", rk_df.columns.tolist())
            st.stop()
            
        # Cek kolom SP2D
        missing_sp2d = required_sp2d - set(sp2d_df.columns.str.lower())
        if missing_sp2d:
            st.error(f"❌ Kolom SP2D tidak lengkap: {missing_sp2d}")
            st.write("Kolom yang tersedia:", sp2d_df.columns.tolist())
            st.stop()
        
        # Proses vouching
        with st.spinner('🔍 Memproses data...'):
            result_df, unmatched_sp2d = perform_vouching(rk_df, sp2d_df)
        
        # Tampilkan statistik
        st.subheader("📊 Hasil Vouching")
        cols = st.columns(4)
        cols[0].metric("Total Transaksi", len(result_df))
        cols[1].metric("Terekoniliasi (Primer)", 
                      len(result_df[result_df['status'] == 'Matched']))
        cols[2].metric("Terekoniliasi (Sekunder)", 
                      len(result_df[result_df['status'] == 'Matched (Secondary)']))
        cols[3].metric("SP2D Belum Terpakai", 
                      len(unmatched_sp2d))
        
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
        st.error(f"❌ Error Kritikal: {str(e)}")
        st.write("Informasi Debugging:")
        st.write("Tipe error:", type(e).__name__)
        st.write("Pesan error:", str(e))
        st.stop()
