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
    sp2d_df['nosp2d_6digits'] = sp2d_df['nosp2d'].astype(str).str[-6:].str.zfill(6)  # Ambil 6 digit terakhir
    
    # Konversi tanggal
    rk_df['tanggal'] = pd.to_datetime(rk_df['tanggal'], errors='coerce')
    sp2d_df['tglsp2d'] = pd.to_datetime(sp2d_df['tglsp2d'], errors='coerce')
    
    # Membuat kunci
    rk_df['key'] = rk_df['nosp2d_6digits'] + '_' + rk_df['jumlah'].astype(str)
    sp2d_df['key'] = sp2d_df['nosp2d_6digits'] + '_' + sp2d_df['jumlah'].astype(str)
    
    # Vouching pertama (kunci SP2D + jumlah)
    merged = rk_df.merge(
        sp2d_df[['key', 'nosp2d', 'tglsp2d', 'skpd']],
        on='key',
        how='left',
        suffixes=('', '_SP2D')
    )
    merged['status'] = merged['nosp2d'].notna().map({True: 'Matched', False: 'Unmatched'})
    
    # Identifikasi data belum terhubung
    used_sp2d = set(merged.loc[merged['status'] == 'Matched', 'key'])
    unmatched_sp2d = sp2d_df[~sp2d_df['key'].isin(used_sp2d)]
    unmatched_rk = merged[merged['status'] == 'Unmatched'].copy()
    
    # Reset indeks untuk memastikan konsistensi
    unmatched_rk = unmatched_rk.reset_index(drop=True)
    unmatched_sp2d = unmatched_sp2d.reset_index(drop=True)
    
    # Vouching kedua (jumlah + tanggal)
    if not unmatched_rk.empty and not unmatched_sp2d.empty:
        second_merge = unmatched_rk.merge(
            unmatched_sp2d,
            left_on=['jumlah', 'tanggal'],
            right_on=['jumlah', 'tglsp2d'],
            how='inner',
            suffixes=('', '_y')
        )
        
        if not second_merge.empty:
            # Update data hasil merge kedua
            merged.loc[second_merge.index, 'nosp2d'] = second_merge['nosp2d_y'].values
            merged.loc[second_merge.index, 'tglsp2d'] = second_merge['tglsp2d_y'].values
            merged.loc[second_merge.index, 'skpd'] = second_merge['skpd_y'].values
            merged.loc[second_merge.index, 'status'] = 'Matched (Secondary)'
            
            # Update daftar SP2D yang digunakan
            used_sp2d.update(second_merge['key_y'])
            unmatched_sp2d = sp2d_df[~sp2d_df['key'].isin(used_sp2d)]
    
    return merged, unmatched_sp2d
