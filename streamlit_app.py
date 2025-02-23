import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

def preprocess_jumlah(series):
    """Fungsi untuk membersihkan format angka dengan separator"""
    series = series.astype(str)
    series = series.str.replace(r'[.]', '', regex=True)  # Hapus separator ribuan
    series = series.str.replace(',', '.', regex=False)    # Ganti desimal koma dengan titik
    return pd.to_numeric(series, errors='coerce')

def extract_sp2d_number(description):
    # Pola regex untuk mengekstrak nomor SP2D dari keterangan
    patterns = [
        r'(?<!\d)\d{6}(?!\d)',  # Pola lama
        r'SP2D NO (\d{6})',     # Pola baru
        r'SP2D NO. (\d{6})',    # Pola dengan titik dua
        r'SP2D NO : (\d{6})',   # Pola dengan titik dua dan spasi
        r'SP2D (\d{6})',        # Pola tanpa "NO"
        r'SP2D: (\d{6})',       # Pola tanpa "NO" dan titik dua
    ]
    for pattern in patterns:
        match = re.search(pattern, str(description))
        if match:
            return match.group(1)
    return None

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
    rk_df['tanggal'] = pd.to_datetime(rk_df['tanggal'], format='%Y-%m-%d', errors='coerce')
    sp2d_df['tglsp2d'] = pd.to_datetime(sp2d_df['tglsp2d'], format='%d/%m/%Y', errors='coerce')
    
    # Membuat kunci
    rk_df['key'] = rk_df['nosp2d_6digits'] + '_' + rk_df['jumlah'].astype(str)
    sp2d_df['key'] = sp2d_df['nosp2d_6digits'] + '_' + sp2d_df['jumlah'].astype(str)
    
    # Debugging preprocessing
    print("RK DataFrame after preprocessing:")
    print(rk_df[['tanggal', 'keterangan', 'jumlah', 'nosp2d_6digits', 'key']].head())
    print("SP2D DataFrame after preprocessing:")
    print(sp2d_df[['tglsp2d', 'nosp2d', 'jumlah', 'nosp2d_6digits', 'key']].head())
    
    # Vouching pertama (kunci SP2D + jumlah)
    merged = rk_df.merge(
        sp2d_df[['key', 'nosp2d', 'tglsp2d', 'skpd']],
        on='key',
        how='left',
        suffixes=('', '_SP2D')
    )
    merged['status'] = merged['nosp2d'].notna().map({True: 'Matched', False: 'Unmatched'})
    
    # Debugging merge primary
    print("Merged DataFrame primary:")
    print(merged[['tanggal', 'jumlah', 'nosp2d_6digits', 'key', 'nosp2d', 'status']].head())
    
    # Identifikasi data belum terhubung
    used_sp2d = set(merged.loc[merged['status'] == 'Matched', 'key'])
    unmatched_sp2d = sp2d_df[~sp2d_df['key'].isin(used_sp2d)]
    unmatched_rk = merged[merged['status'] == 'Unmatched'].copy()
    
    # Debugging data yang tidak cocok
    print("Unmatched RK DataFrame:")
    print(unmatched_rk[['tanggal', 'jumlah', 'nosp2d_6digits', 'key']].head())
    
    # Vouching kedua (jumlah + tanggal + skpd)
    if not unmatched_rk.empty and not unmatched_sp2d.empty:
        second_merge = unmatched_rk.merge(
            unmatched_sp2d,
            left_on=['jumlah', 'tanggal', 'skpd'],
            right_on=['jumlah', 'tglsp2d', 'skpd'],
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
        
        # Debugging merge secondary
        print("Second Merge DataFrame:")
        print(second_merge[['tanggal', 'jumlah', 'nosp2d_6digits', 'key', 'nosp2d_y', 'status']].head())
    else:
        print("Tidak ada data yang tersisa untuk merge kedua.")
    
    # Layer tambahan untuk kasus khusus
    if not unmatched_rk.empty and not unmatched_sp2d.empty:
        third_merge = unmatched_rk.merge(
            unmatched_sp2d,
            left_on=['jumlah', 'tanggal'],
            right_on=['jumlah', 'tglsp2d'],
            how='inner',
            suffixes=('', '_y')
        )
        
        if not third_merge.empty:
            # Update data hasil merge ketiga
            merged.loc[third_merge.index, 'nosp2d'] = third_merge['nosp2d_y']
            merged.loc[third_merge.index, 'tglsp2d'] = third_merge['tglsp2d_y']
            merged.loc[third_merge.index, 'skpd'] = third_merge['skpd_y']
            merged.loc[third_merge.index, 'status'] = 'Matched (Third Layer)'
            
            # Update daftar SP2D yang digunakan
            used_sp2d.update(third_merge['key_y'])
            unmatched_sp2d = sp2d_df[~sp2d_df['key'].isin(used_sp2d)]
        
        # Debugging merge third layer
        print("Third Merge DataFrame:")
        print(third_merge[['tanggal', 'jumlah', 'nosp2d_6digits', 'key', 'nosp2d_y', 'status']].head())
    else:
        print("Tidak ada data yang tersisa untuk merge ketiga.")
    
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
        cols[3].metric("Matched (Third Layer)", len(merged_rk[merged_rk['status'] == 'Matched (Third Layer)']))
        cols[4].metric("Unmatched SP2D", len(unmatched_sp2d))
        
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
