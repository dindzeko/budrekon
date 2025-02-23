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
    
    # Membuat kunci unik
    rk_df['key'] = rk_df['nosp2d_6digits'] + '_' + rk_df['jumlah'].astype(str)
    sp2d_df['key'] = sp2d_df['nosp2d_6digits'] + '_' + sp2d_df['jumlah'].astype(str)
    
    # Proses matching
    merged = rk_df.merge(
        sp2d_df[['key', 'nosp2d', 'tglsp2d', 'skpd', 'jumlah']],
        on='key',
        how='left',
        suffixes=('', '_SP2D')
    )
    
    # Klasifikasi data
    merged['status'] = merged['nosp2d'].notna().map({True: 'Matched', False: 'Unmatched'})
    
    # Identifikasi SP2D yang tidak terpakai
    used_sp2d_keys = set(merged.loc[merged['status'] == 'Matched', 'key'])
    unmatched_sp2d = sp2d_df[~sp2d_df['key'].isin(used_sp2d_keys)]
    
    return merged, unmatched_sp2d

# Fungsi alternatif untuk matching berdasarkan jumlah dan tanggal
def alternative_matching(unmatched_rk, sp2d_df):
    #konversi data
    unmatched_rk = unmatched_rk.copy()
    sp2d_df = sp2d_df.copy()
    unmatched_rk['tanggal'] = pd.to_datetime(unmatched_rk['tanggal'], errors='coerce')
    sp2d_df['tglsp2d'] = pd.to_datetime(sp2d_df['tglsp2d'], errors='coerce')

    matched_indices = []
    for rk_index, rk_row in unmatched_rk.iterrows():
        for sp2d_index, sp2d_row in sp2d_df.iterrows():
            if rk_row['jumlah'] == sp2d_row['jumlah'] and rk_row['tanggal'].date() == sp2d_row['tglsp2d'].date():
                matched_indices.append((rk_index, sp2d_index))
                break  # Stop searching for this RK row once matched

    # Update the matched RK rows with SP2D data
    for rk_index, sp2d_index in matched_indices:
        rk_row = unmatched_rk.loc[rk_index]
        sp2d_row = sp2d_df.loc[sp2d_index]
        
        unmatched_rk.loc[rk_index, 'nosp2d'] = sp2d_row['nosp2d']
        unmatched_rk.loc[rk_index, 'tglsp2d'] = sp2d_row['tglsp2d']
        unmatched_rk.loc[rk_index, 'skpd'] = sp2d_row['skpd']
        unmatched_rk.loc[rk_index, 'status'] = 'Matched (Alt)'

    # Create a DataFrame for truly unmatched SP2D
    matched_sp2d_indices = [sp2d_index for _, sp2d_index in matched_indices]
    unmatched_sp2d = sp2d_df[~sp2d_df.index.isin(matched_sp2d_indices)]

    return unmatched_rk, unmatched_sp2d

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
        
        # Proses vouching awal
        with st.spinner('Memproses data (Tahap 1)...'):
            merged_rk, unmatched_sp2d = perform_vouching(rk_df, sp2d_df)

        # Filter data RK yang unmatched
        unmatched_rk = merged_rk[merged_rk['status'] == 'Unmatched'].copy()

        # Alternative Matching
        with st.spinner('Memproses data (Tahap 2 - Alternatif)...'):
            unmatched_rk_alt, unmatched_sp2d_alt = alternative_matching(unmatched_rk, sp2d_df)

        # Gabungkan hasil alternative matching ke data utama
        merged_rk.loc[unmatched_rk_alt.index, 'nosp2d'] = unmatched_rk_alt['nosp2d']
        merged_rk.loc[unmatched_rk_alt.index, 'tglsp2d'] = unmatched_rk_alt['tglsp2d']
        merged_rk.loc[unmatched_rk_alt.index, 'skpd'] = unmatched_rk_alt['skpd']
        merged_rk.loc[unmatched_rk_alt.index, 'status'] = unmatched_rk_alt['status']
        # Update data SP2D yang masih unmatched
        unmatched_sp2d = unmatched_sp2d_alt
        # Tampilkan statistik
        st.subheader("Statistik")
        cols = st.columns(3)

        matched_count = len(merged_rk[merged_rk['status'].str.contains('Matched')])
        unmatched_count = len(merged_rk[merged_rk['status'] == 'Unmatched'])
        cols[0].metric("RK Matched", matched_count)
        cols[1].metric("RK Unmatched", unmatched_count)
        cols[2].metric("SP2D Unmatched", len(unmatched_sp2d))
        
        # Buat file Excel untuk di-download
        df_list = [merged_rk, unmatched_sp2d]
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
