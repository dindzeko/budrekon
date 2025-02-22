import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Fungsi untuk mengekstrak nomor SP2D 6 digit dari kolom Keterangan
def extract_sp2d_number(keterangan):
    matches = re.findall(r'\d{6}', str(keterangan))
    return matches[0] if matches else None

# Fungsi utama untuk memproses file RK dan SP2D
def process_files(rk_file, sp2d_file):
    try:
        # Proses file RK
        rk_df = pd.read_excel(rk_file)
        # Validasi kolom RK
        required_columns_rk = {'Keterangan', 'Jumlah'}
        if not required_columns_rk.issubset(rk_df.columns):
            raise ValueError(f"File RK harus memiliki kolom: {required_columns_rk}")
        
        rk_df['NoSP2D_6digits'] = rk_df['Keterangan'].apply(extract_sp2d_number)
        rk_df['Jumlah'] = pd.to_numeric(rk_df['Jumlah'], errors='coerce')
        rk_df['key'] = rk_df['NoSP2D_6digits'].astype(str) + '_' + rk_df['Jumlah'].astype(str)

        # Proses file SP2D
        sp2d_df = pd.read_excel(sp2d_file)
        # Validasi kolom SP2D
        required_columns_sp2d = {'NoSP2D', 'Jumlah'}
        if not required_columns_sp2d.issubset(sp2d_df.columns):
            raise ValueError(f"File SP2D harus memiliki kolom: {required_columns_sp2d}")
        
        sp2d_df['NoSP2D_6digits'] = sp2d_df['NoSP2D'].astype(str).str[:6]
        sp2d_df['Jumlah'] = pd.to_numeric(sp2d_df['Jumlah'], errors='coerce')
        sp2d_df['key'] = sp2d_df['NoSP2D_6digits'] + '_' + sp2d_df['Jumlah'].astype(str)

        # Dapatkan keys unik dari kedua DataFrame
        sp2d_keys = set(sp2d_df['key'].unique())
        rk_keys = set(rk_df['key'].unique())

        # Pisahkan data RK menjadi matched dan unmatched
        rk_matched = rk_df[rk_df['key'].isin(sp2d_keys)]
        rk_unmatched = rk_df[~rk_df['key'].isin(sp2d_keys)]

        # Pisahkan data SP2D menjadi matched dan unmatched
        sp2d_matched = sp2d_df.merge(
            rk_df[['key', 'Keterangan']], 
            on='key', 
            how='inner'
        ).rename(columns={'Keterangan': 'Keterangan RK'})

        sp2d_unmatched = sp2d_df[~sp2d_df['key'].isin(rk_keys)]

        # Siapkan output dalam format Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Sheet untuk RK Matched
            rk_matched[['Tanggal', 'Keterangan', 'Jumlah', 'NoSP2D_6digits']].rename(
                columns={'NoSP2D_6digits': 'NO SP2D'}
            ).to_excel(writer, sheet_name='RK Matched', index=False)

            # Sheet untuk RK Unmatched
            rk_unmatched[['Tanggal', 'Keterangan', 'Jumlah', 'NoSP2D_6digits']].rename(
                columns={'NoSP2D_6digits': 'NO SP2D'}
            ).to_excel(writer, sheet_name='RK Unmatched', index=False)

            # Sheet untuk SP2D Matched
            sp2d_matched[['SKPD', 'NoSP2D', 'TglSP2D', 'Jumlah', 'Keterangan RK']].to_excel(
                writer, sheet_name='SP2D Matched', index=False
            )

            # Sheet untuk SP2D Unmatched
            sp2d_unmatched[['SKPD', 'NoSP2D', 'TglSP2D', 'Jumlah']].assign(**{'Keterangan RK': ''}).to_excel(
                writer, sheet_name='SP2D Unmatched', index=False
            )

        output.seek(0)
        return output, rk_matched, rk_unmatched, sp2d_matched, sp2d_unmatched

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
        return None, None, None, None, None

# Antarmuka Pengguna Streamlit
st.title('Aplikasi Vouching SP2D-RK')
st.subheader("Upload File")

# Upload file RK dan SP2D
uploaded_rk = st.file_uploader("Rekening Koran (Excel)", type=['xlsx'])
uploaded_sp2d = st.file_uploader("SP2D (Excel)", type=['xlsx'])

if uploaded_rk and uploaded_sp2d:
    if st.button('Proses Vouching'):
        with st.spinner('Memproses...'):
            # Simpan DataFrame dari file yang diunggah
            try:
                rk_df = pd.read_excel(uploaded_rk)
                sp2d_df = pd.read_excel(uploaded_sp2d)
            except Exception as e:
                st.error(f"Gagal membaca file Excel: {e}")
                st.stop()

            # Proses file
            result, rk_matched, rk_unmatched, sp2d_matched, sp2d_unmatched = process_files(uploaded_rk, uploaded_sp2d)

            if result is not None:
                st.success('Proses selesai!')
                # Tombol unduh hasil
                st.download_button(
                    label="Unduh Hasil",
                    data=result.getvalue(),
                    file_name="hasil_vouching.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Tampilkan statistik hasil
                st.subheader("Statistik Hasil")
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Total Transaksi RK", len(rk_df))
                with col2:
                    st.metric("Total Transaksi SP2D", len(sp2d_df))

                col1, col2 = st.columns(2)
                with col1:
                    st.metric("RK Matched", len(rk_matched))
                    st.metric("RK Unmatched", len(rk_unmatched))
                with col2:
                    st.metric("SP2D Matched", len(sp2d_matched))
                    st.metric("SP2D Unmatched", len(sp2d_unmatched))
