import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

def preprocess_jumlah(series):
    """Membersihkan format angka dengan separator ribuan dan desimal"""
    series = series.astype(str)
    series = series.str.replace(r'[.]', '', regex=True)  # Hapus titik ribuan
    series = series.str.replace(',', '.', regex=False)    # Ganti koma desimal
    return pd.to_numeric(series, errors='coerce')

def extract_sp2d_number(description):
    """Ekstraksi 6 digit terakhir SP2D dari keterangan RK"""
    numbers = re.findall(r'\d+', str(description))  # Cari semua angka
    
    if numbers:
        last_number = numbers[-1]  # Ambil angka terakhir
        trimmed = last_number[-6:] if len(last_number) >=6 else last_number  # Potong 6 digit terakhir
        padded = trimmed.zfill(6)  # Tambahkan 0 di depan jika kurang dari 6 digit
        return padded
    return None

@st.cache_data
def perform_vouching(rk_df, sp2d_df):
    # Salin dataframe untuk menghindari modifikasi data asli
    rk_df = rk_df.copy()
    sp2d_df = sp2d_df.copy()
    
    # Normalisasi nama kolom
    rk_df.columns = rk_df.columns.str.strip().str.lower()
    sp2d_df.columns = sp2d_df.columns.str.strip().str.lower()
    
    # Preprocessing jumlah
    rk_df['jumlah'] = preprocess_jumlah(rk_df['jumlah'])
    sp2d_df['jumlah'] = preprocess_jumlah(sp2d_df['jumlah'])
    
    # Ekstraksi 6 digit SP2D (RK)
    rk_df['nosp2d_6digits'] = rk_df['keterangan'].apply(extract_sp2d_number)
    
    # Ekstraksi 6 digit terakhir SP2D (Data SP2D)
    sp2d_df['nosp2d_6digits'] = (
        sp2d_df['nosp2d']
        .astype(str)
        .str[-6:]  # Ambil 6 digit terakhir
        .str.zfill(6)  # Tambahkan 0 di depan jika perlu
    )
    
    # Konversi tanggal
    rk_df['tanggal'] = pd.to_datetime(rk_df['tanggal'], errors='coerce')
    sp2d_df['tglsp2d'] = pd.to_datetime(sp2d_df['tglsp2d'], errors='coerce')
    
    # Membuat kunci gabungan
    rk_df['key'] = rk_df['nosp2d_6digits'] + '_' + rk_df['jumlah'].astype(str)
    sp2d_df['key'] = sp2d_df['nosp2d_6digits'] + '_' + sp2d_df['jumlah'].astype(str)
    
    # Vouching pertama: Cocokkan kunci SP2D + jumlah
    merged = rk_df.merge(
        sp2d_df[['key', 'nosp2d', 'tglsp2d', 'skpd']],
        on='key',
        how='left',
        suffixes=('', '_sp2d')
    )
    merged['status'] = merged['nosp2d'].notna().map({True: 'Matched', False: 'Unmatched'})
    
    # Identifikasi SP2D yang belum terpakai
    used_sp2d = set(merged.loc[merged['status'] == 'Matched', 'key'])
    unmatched_sp2d = sp2d_df[~sp2d_df['key'].isin(used_sp2d)]
    
    # Vouching kedua: Cocokkan jumlah + tanggal
    unmatched_rk = merged[merged['status'] == 'Unmatched'].copy()
    unmatched_rk['original_index'] = unmatched_rk.index  # Simpan index asli
    
    if not unmatched_rk.empty and not unmatched_sp2d.empty:
        second_merge = unmatched_rk.merge(
            unmatched_sp2d,
            left_on=['jumlah', 'tanggal'],
            right_on=['jumlah', 'tglsp2d'],
            how='inner',
            suffixes=('', '_y')
        )
        
        if not second_merge.empty:
            # Update data di dataframe utama menggunakan index asli
            merged.loc[second_merge['original_index'], 'nosp2d'] = second_merge['nosp2d_y']
            merged.loc[second_merge['original_index'], 'tglsp2d'] = second_merge['tglsp2d_y']
            merged.loc[second_merge['original_index'], 'skpd'] = second_merge['skpd_y']
            merged.loc[second_merge['original_index'], 'status'] = 'Matched (Secondary)'
            
            # Tandai SP2D yang sudah digunakan
            used_sp2d.update(second_merge['key_y'])
    
    return merged, sp2d_df[~sp2d_df['key'].isin(used_sp2d)]

def to_excel(df_list, sheet_names):
    """Konversi dataframe ke file Excel"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, sheet_name in zip(df_list, sheet_names):
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# Antarmuka Streamlit
st.title("🕵️ Aplikasi Vouching SP2D vs Rekening Koran")
st.write("**Fitur Perbaikan:**")
st.write("- Pencarian 6 digit terakhir SP2D di keterangan RK")
st.write("- Pencocokan ganda (nomor+jumlah dan tanggal+jumlah)")

rk_file = st.file_uploader("Upload Rekening Koran (Format Excel)", type="xlsx")
sp2d_file = st.file_uploader("Upload Data SP2D (Format Excel)", type="xlsx")

if rk_file and sp2d_file:
    try:
        rk_df = pd.read_excel(rk_file)
        sp2d_df = pd.read_excel(sp2d_file)
        
        # Validasi kolom
        required_rk = {'tanggal', 'keterangan', 'jumlah'}
        required_sp2d = {'skpd', 'nosp2d', 'tglsp2d', 'jumlah'}
        
        if not required_rk.issubset(rk_df.columns.str.lower()):
            missing = required_rk - set(rk_df.columns.str.lower())
            st.error(f"Kolom RK tidak lengkap! Yang kurang: {missing}")
            st.stop()
            
        if not required_sp2d.issubset(sp2d_df.columns.str.lower()):
            missing = required_sp2d - set(sp2d_df.columns.str.lower())
            st.error(f"Kolom SP2D tidak lengkap! Yang kurang: {missing}")
            st.stop()
        
        # Proses vouching
        with st.spinner('🔍 Mencocokkan data...'):
            merged_rk, unmatched_sp2d = perform_vouching(rk_df, sp2d_df)
        
        # Tampilkan statistik
        st.subheader("📊 Hasil Vouching")
        cols = st.columns(4)
        cols[0].metric("Total Transaksi RK", len(merged_rk))
        cols[1].metric("Tercocokkan (Primary)", 
                      len(merged_rk[merged_rk['status'] == 'Matched']),
                      help="Cocok berdasarkan 6 digit SP2D + jumlah")
        cols[2].metric("Tercocokkan (Secondary)", 
                      len(merged_rk[merged_rk['status'] == 'Matched (Secondary)']),
                      help="Cocok berdasarkan jumlah + tanggal")
        cols[3].metric("SP2D Tidak Terpakai", 
                      len(unmatched_sp2d),
                      help="SP2D yang tidak memiliki transaksi terkait")
        
        # Download hasil
        excel_data = to_excel(
            [merged_rk, unmatched_sp2d],
            ['Hasil Vouching', 'SP2D Belum Terpakai']
        )
        
        st.download_button(
            label="💾 Download Hasil",
            data=excel_data,
            file_name=f"hasil_vouching_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"🚨 Error: {str(e)}")
        st.stop()
