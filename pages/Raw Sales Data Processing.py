import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

st.set_page_config(page_title='Raw Sales Data Processing', layout='centered')
st.title('🔗 Raw Sales Data Processing')
st.markdown('---')

# Fungsi untuk merge AP Invoice (FIXED)
def merge_ap_invoice(df_contens, df_awal, apply_filter=True):
    # Strip spasi dari nama kolom
    df_contens.columns = df_contens.columns.str.strip()
    df_awal.columns = df_awal.columns.str.strip()
    
    # Simpan jumlah baris awal untuk metric
    original_contens_len = len(df_contens)
    
    # === FILTER SEBELUM MERGE (FIXED!) ===
    item_desc_col = None
    if apply_filter:
        # Cari kolom Item Description
        for col in df_contens.columns:
            if "item" in col.lower() and "desc" in col.lower():
                item_desc_col = col
                break
        
        if item_desc_col:
            # Filter df_contens DULU sebelum merge
            mask = df_contens[item_desc_col].astype(str).str.lower().str.contains(
                r'jagung|zak|wheat bran', 
                case=False, 
                na=False, 
                regex=True
            )
            df_contens = df_contens[mask].reset_index(drop=True)
    
    # Ambil "base name" (tanpa angka di belakang) untuk kolom dari contens
    kolom_base = set(re.sub(r"[\s._]*\d+$", "", col).strip() for col in df_contens.columns)
    
    # Tentukan kolom tambahan dari df_awal yang belum ada di df_contens
    kolom_tambahan = [
        col for col in df_awal.columns
        if re.sub(r"[\s._]*\d+$", "", col).strip() not in kolom_base and col != "Doc Number"
    ]
    
    # Merge: base = df_contens (sudah difilter), tambahkan kolom tambahan dari df_awal
    df_merged = pd.merge(
        df_contens,
        df_awal[["Doc Number"] + kolom_tambahan],
        on="Doc Number",
        how="inner"  # hanya Doc Number yang cocok
    )
    
    # === Bersihkan nama kolom duplikat berbasis nama dasar ===
    def base_name(col):
        return re.sub(r"[\s._]*\d+$", "", col).strip()
    
    base_names = pd.Series([base_name(c) for c in df_merged.columns], index=df_merged.columns)
    df_merged = df_merged.loc[:, ~base_names.duplicated()]
    
    # === Rename khusus kalau ada 'doc date 2' ===
    if "doc date 2" in df_merged.columns:
        df_merged = df_merged.rename(columns={"doc date 2": "Doc Date"})
    
    # === Hapus kolom yang tidak dipakai ===
    kolom_hapus = ["Total", "Applied Amount", "Total Before Diskon", "Customer", "Unnamed: 45", "Netto 1"]
    df_merged = df_merged.drop(columns=[c for c in kolom_hapus if c in df_merged.columns])
    
    return df_merged, original_contens_len

# Fungsi untuk merge GRPO (FIXED - sama dengan AP Invoice)
def merge_grpo(df_contens, df_awal, apply_filter=True):
    # Strip spasi dari nama kolom
    df_contens.columns = df_contens.columns.str.strip()
    df_awal.columns = df_awal.columns.str.strip()
    
    # Simpan jumlah baris awal untuk metric
    original_contens_len = len(df_contens)
    
    # === FILTER SEBELUM MERGE (FIXED!) ===
    item_desc_col = None
    if apply_filter:
        # Cari kolom Item Description
        for col in df_contens.columns:
            if "item" in col.lower() and "desc" in col.lower():
                item_desc_col = col
                break
        
        if item_desc_col:
            # Filter df_contens DULU sebelum merge
            mask = df_contens[item_desc_col].astype(str).str.lower().str.contains(
                r'jagung|zak|wheat bran', 
                case=False, 
                na=False, 
                regex=True
            )
            df_contens = df_contens[mask].reset_index(drop=True)
    
    # Ambil "base name" (tanpa angka di belakang) untuk kolom dari contens
    kolom_base = set(re.sub(r"[\s._]*\d+$", "", col).strip() for col in df_contens.columns)
    
    # Tentukan kolom tambahan dari df_awal yang belum ada di df_contens
    kolom_tambahan = [
        col for col in df_awal.columns
        if re.sub(r"[\s._]*\d+$", "", col).strip() not in kolom_base and col != "Doc Number"
    ]
    
    # Merge: base = df_contens (sudah difilter), tambahkan kolom tambahan dari df_awal
    df_merged = pd.merge(
        df_contens,
        df_awal[["Doc Number"] + kolom_tambahan],
        on="Doc Number",
        how="inner"  # hanya Doc Number yang cocok
    )
    
    # === Bersihkan nama kolom duplikat berbasis nama dasar ===
    def base_name(col):
        return re.sub(r"[\s._]*\d+$", "", col).strip()
    
    base_names = pd.Series([base_name(c) for c in df_merged.columns], index=df_merged.columns)
    df_merged = df_merged.loc[:, ~base_names.duplicated()]
    
    # === Rename khusus kalau ada 'doc date 2' ===
    if "doc date 2" in df_merged.columns:
        df_merged = df_merged.rename(columns={"doc date 2": "Doc Date"})
    
    # === Hapus kolom yang tidak dipakai ===
    kolom_hapus = ["Total", "Applied Amount", "Total Before Diskon", "Customer", "Unnamed: 45", "Netto 1"]
    df_merged = df_merged.drop(columns=[c for c in kolom_hapus if c in df_merged.columns])
    
    return df_merged, original_contens_len

# Fungsi untuk convert dataframe ke Excel
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# Section 1: AR Invoice
st.header('📄 A/R Invoice')

col1, col2 = st.columns(2)

with col1:
    file_contens_invoice = st.file_uploader(
        'Upload A/R Invoice Contents File:', 
        type=['xlsx'], 
        key='contens_invoice',
        help='File yang berisi data contents/detail'
    )

with col2:
    file_awal_invoice = st.file_uploader(
        'Upload A/R Invoice Base File:', 
        type=['xlsx'], 
        key='awal_invoice',
        help='File yang berisi data awal/header'
    )

if file_contens_invoice and file_awal_invoice:
    try:
        with st.spinner('Process A/R Invoice...'):
            # Baca file
            df_contens = pd.read_excel(file_contens_invoice)
            df_awal = pd.read_excel(file_awal_invoice)
            
            # Simpan jumlah baris awal
            original_awal_len = len(df_awal)
            
            # Merge data dengan logika AR Invoice
            df_merged, original_contens_len = merge_ap_invoice(
                df_contens, df_awal, apply_filter=True  # Selalu filter
            )
            
            # Tampilkan info
            st.success('Success!')
            
            col_info1, col_info2, col_info3 = st.columns(3)
            with col_info1:
                st.metric('Baris Contents', original_contens_len)
            with col_info2:
                st.metric('Baris Awal', original_awal_len)
            with col_info3:
                st.metric('Baris Hasil', len(df_merged))
            
            # Preview data
            st.subheader('Preview Result Data:')
            st.dataframe(df_merged.head(10), use_container_width=True)
            
            # Ekstrak tahun dari Posting Date untuk nama file
            years_in_data = []
            if 'Posting Date' in df_merged.columns:
                # Convert ke datetime dan ekstrak tahun
                posting_dates = pd.to_datetime(df_merged['Posting Date'], errors='coerce')
                years_in_data = sorted(posting_dates.dt.year.dropna().unique().astype(int).tolist())
            
            # Buat nama file dengan tahun
            if years_in_data:
                if len(years_in_data) == 1:
                    year_str = str(years_in_data[0])
                else:
                    year_str = f"{years_in_data[0]}-{years_in_data[-1]}"
                file_name = f'AR_INVOICE_{year_str}.xlsx'
            else:
                # Fallback jika tidak ada tahun
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name = f'AR_INVOICE_{timestamp}.xlsx'
            
            # Button download
            excel_data = to_excel(df_merged)
            
            st.download_button(
                label='📥 Download A/R Invoice Data',
                data=excel_data,
                file_name=file_name,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary'
            )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.info('💡 Make sure both files have a "Doc Number" column')

st.markdown('---')

# Section 2: Delivery Order
st.header('🚚 Delivery Order')

col3, col4 = st.columns(2)

with col3:
    file_contens_do = st.file_uploader(
        'Upload Delivery Order Contents File:', 
        type=['xlsx'], 
        key='contens_do',
        help='File yang berisi data contents/detail'
    )

with col4:
    file_awal_do = st.file_uploader(
        'Upload Delivery Order Base File:', 
        type=['xlsx'], 
        key='awal_do',
        help='File yang berisi data awal/header'
    )

if file_contens_do and file_awal_do:
    try:
        with st.spinner('Process Delivery Order...'):
            # Baca file
            df_contens_do = pd.read_excel(file_contens_do)
            df_awal_do = pd.read_excel(file_awal_do)
            
            # Simpan jumlah baris awal
            original_awal_do_len = len(df_awal_do)
            
            # Merge data dengan logika yang sama
            df_merged_do, original_contens_do_len = merge_grpo(
                df_contens_do, df_awal_do, apply_filter=True  # Selalu filter
            )
            
            # Tampilkan info
            st.success('Success!')
            
            col_info4, col_info5, col_info6 = st.columns(3)
            with col_info4:
                st.metric('Baris Contents', original_contens_do_len)
            with col_info5:
                st.metric('Baris Awal', original_awal_do_len)
            with col_info6:
                st.metric('Baris Hasil', len(df_merged_do))
            
            # Preview data
            st.subheader('Preview Result Data:')
            st.dataframe(df_merged_do.head(10), use_container_width=True)
            
            # Ekstrak tahun dari Posting Date untuk nama file
            years_in_data_do = []
            if 'Posting Date' in df_merged_do.columns:
                # Convert ke datetime dan ekstrak tahun
                posting_dates_do = pd.to_datetime(df_merged_do['Posting Date'], errors='coerce')
                years_in_data_do = sorted(posting_dates_do.dt.year.dropna().unique().astype(int).tolist())
            
            # Buat nama file dengan tahun
            if years_in_data_do:
                if len(years_in_data_do) == 1:
                    year_str_do = str(years_in_data_do[0])
                else:
                    year_str_do = f"{years_in_data_do[0]}-{years_in_data_do[-1]}"
                file_name_do = f'DO_{year_str_do}.xlsx'
            else:
                # Fallback jika tidak ada tahun
                timestamp_do = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name_do = f'DO_{timestamp_do}.xlsx'
            
            # Button download
            excel_data_do = to_excel(df_merged_do)
            
            st.download_button(
                label='📥 Download Delivery Order Data',
                data=excel_data_do,
                file_name=file_name_do,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary'
            )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.info('💡 Make sure both files have a "Doc Number" column')

st.markdown('---')

# Section 3: AR Down Payment
st.header('💰 A/R Down Payment')

col5, col6 = st.columns(2)

with col5:
    file_contens_ardp = st.file_uploader(
        'Upload A/R Down Payment Contents File:', 
        type=['xlsx'], 
        key='contens_ardp',
        help='File yang berisi data contents/detail'
    )

with col6:
    file_awal_ardp = st.file_uploader(
        'Upload A/R Down Payment Base File:', 
        type=['xlsx'], 
        key='awal_ardp',
        help='File yang berisi data awal/header'
    )

if file_contens_ardp and file_awal_ardp:
    try:
        with st.spinner('Process A/R Down Payment...'):
            # Baca file
            df_contens_ardp = pd.read_excel(file_contens_ardp)
            df_awal_ardp = pd.read_excel(file_awal_ardp)
            
            # Simpan jumlah baris awal
            original_awal_ardp_len = len(df_awal_ardp)
            
            # Merge data dengan logika yang sama
            df_merged_ardp, original_contens_ardp_len = merge_ap_invoice(
                df_contens_ardp, df_awal_ardp, apply_filter=True  # Selalu filter
            )
            
            # Tampilkan info
            st.success('Success!')
            
            col_info7, col_info8, col_info9 = st.columns(3)
            with col_info7:
                st.metric('Baris Contents', original_contens_ardp_len)
            with col_info8:
                st.metric('Baris Awal', original_awal_ardp_len)
            with col_info9:
                st.metric('Baris Hasil', len(df_merged_ardp))
            
            # Preview data
            st.subheader('Preview Result Data:')
            st.dataframe(df_merged_ardp.head(10), use_container_width=True)
            
            # Ekstrak tahun dari Posting Date untuk nama file
            years_in_data_ardp = []
            if 'Posting Date' in df_merged_ardp.columns:
                # Convert ke datetime dan ekstrak tahun
                posting_dates_ardp = pd.to_datetime(df_merged_ardp['Posting Date'], errors='coerce')
                years_in_data_ardp = sorted(posting_dates_ardp.dt.year.dropna().unique().astype(int).tolist())
            
            # Buat nama file dengan tahun
            if years_in_data_ardp:
                if len(years_in_data_ardp) == 1:
                    year_str_ardp = str(years_in_data_ardp[0])
                else:
                    year_str_ardp = f"{years_in_data_ardp[0]}-{years_in_data_ardp[-1]}"
                file_name_ardp = f'ARDP_{year_str_ardp}.xlsx'
            else:
                # Fallback jika tidak ada tahun
                timestamp_ardp = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name_ardp = f'ARDP_{timestamp_ardp}.xlsx'
            
            # Button download
            excel_data_ardp = to_excel(df_merged_ardp)
            
            st.download_button(
                label='📥 Download A/R Down Payment Data',
                data=excel_data_ardp,
                file_name=file_name_ardp,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary'
            )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.info('💡 Make sure both files have a "Doc Number" column')

st.markdown('---')

# Section 4: Sales Order
st.header('📋 Sales Order')

col7, col8 = st.columns(2)

with col7:
    file_contens_so = st.file_uploader(
        'Upload Sales Order Contents File:', 
        type=['xlsx'], 
        key='contens_so',
        help='File yang berisi data contents/detail'
    )

with col8:
    file_awal_so = st.file_uploader(
        'Upload Sales Order Base File :', 
        type=['xlsx'], 
        key='awal_so',
        help='File yang berisi data awal/header'
    )

if file_contens_so and file_awal_so:
    try:
        with st.spinner('Process Sales Order...'):
            # Baca file
            df_contens_so = pd.read_excel(file_contens_so)
            df_awal_so = pd.read_excel(file_awal_so)
            
            # Simpan jumlah baris awal
            original_awal_so_len = len(df_awal_so)
            
            # Merge data dengan logika yang sama
            df_merged_so, original_contens_so_len = merge_ap_invoice(
                df_contens_so, df_awal_so, apply_filter=True  # Selalu filter
            )
            
            # Tampilkan info
            st.success('Success!')
            
            col_info10, col_info11, col_info12 = st.columns(3)
            with col_info10:
                st.metric('Baris Contents', original_contens_so_len)
            with col_info11:
                st.metric('Baris Awal', original_awal_so_len)
            with col_info12:
                st.metric('Baris Hasil', len(df_merged_so))
            
            # Preview data
            st.subheader('Preview Result Data:')
            st.dataframe(df_merged_so.head(10), use_container_width=True)
            
            # Ekstrak tahun dari Posting Date untuk nama file
            years_in_data_so = []
            if 'Posting Date' in df_merged_so.columns:
                # Convert ke datetime dan ekstrak tahun
                posting_dates_so = pd.to_datetime(df_merged_so['Posting Date'], errors='coerce')
                years_in_data_so = sorted(posting_dates_so.dt.year.dropna().unique().astype(int).tolist())
            
            # Buat nama file dengan tahun
            if years_in_data_so:
                if len(years_in_data_so) == 1:
                    year_str_so = str(years_in_data_so[0])
                else:
                    year_str_so = f"{years_in_data_so[0]}-{years_in_data_so[-1]}"
                file_name_so = f'SO_{year_str_so}.xlsx'
            else:
                # Fallback jika tidak ada tahun
                timestamp_so = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name_so = f'SO_{timestamp_so}.xlsx'
            
            # Button download
            excel_data_so = to_excel(df_merged_so)
            
            st.download_button(
                label='📥 Download Sales Order Data',
                data=excel_data_so,
                file_name=file_name_so,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary'
            )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.info('💡 Make sure both files have a "Doc Number" column')

st.markdown('---')

# Section 5: Timbangan Jual
st.header('⚖️ Delivery Loading')

col9, col10 = st.columns(2)

with col9:
    file_contens_timbangan = st.file_uploader(
        'Upload Weighing Contents File:', 
        type=['xlsx'], 
        key='contens_timbangan',
        help='File yang berisi data contents/detail'
    )

with col10:
    file_awal_timbangan = st.file_uploader(
        'Upload Weighing Base File:', 
        type=['xlsx'], 
        key='awal_timbangan',
        help='File yang berisi data awal/header'
    )

if file_contens_timbangan and file_awal_timbangan:
    try:
        with st.spinner('Process Weighing...'):
            # Baca file
            df_contens_timbangan = pd.read_excel(file_contens_timbangan)
            df_awal_timbangan = pd.read_excel(file_awal_timbangan)
            
            # Simpan jumlah baris awal
            original_awal_timbangan_len = len(df_awal_timbangan)
            
            # Merge data dengan logika yang sama
            df_merged_timbangan, original_contens_timbangan_len = merge_ap_invoice(
                df_contens_timbangan, df_awal_timbangan, apply_filter=True  # Selalu filter
            )
            
            # Tampilkan info
            st.success('Success!')
            
            col_info13, col_info14, col_info15 = st.columns(3)
            with col_info13:
                st.metric('Baris Contents', original_contens_timbangan_len)
            with col_info14:
                st.metric('Baris Awal', original_awal_timbangan_len)
            with col_info15:
                st.metric('Baris Hasil', len(df_merged_timbangan))
            
            # Preview data
            st.subheader('Preview Result Data:')
            st.dataframe(df_merged_timbangan.head(10), use_container_width=True)
            
            # Ekstrak tahun dari Out Date untuk nama file (khusus Timbangan)
            years_in_data_timbangan = []
            if 'Out Date' in df_merged_timbangan.columns:
                # Convert ke datetime dan ekstrak tahun
                out_dates = pd.to_datetime(df_merged_timbangan['Out Date'], errors='coerce')
                years_in_data_timbangan = sorted(out_dates.dt.year.dropna().unique().astype(int).tolist())
            
            # Buat nama file dengan tahun
            if years_in_data_timbangan:
                if len(years_in_data_timbangan) == 1:
                    year_str_timbangan = str(years_in_data_timbangan[0])
                else:
                    year_str_timbangan = f"{years_in_data_timbangan[0]}-{years_in_data_timbangan[-1]}"
                file_name_timbangan = f'TIMBANGAN JUAL_{year_str_timbangan}.xlsx'
            else:
                # Fallback jika tidak ada tahun
                timestamp_timbangan = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name_timbangan = f'TIMBANGAN JUAL_{timestamp_timbangan}.xlsx'
            
            # Button download
            excel_data_timbangan = to_excel(df_merged_timbangan)
            
            st.download_button(
                label='📥 Download Weighing Data',
                data=excel_data_timbangan,
                file_name=file_name_timbangan,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary'
            )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.info('💡 Make sure both files have a "Doc Number" column')

st.markdown('---')

# Section 6: A/R Credit Memo
st.header('💳 A/R Credit Memo')

col11, col12 = st.columns(2)

with col11:
    file_contens_arcm = st.file_uploader(
        'Upload A/R Credit Memo Contents File:', 
        type=['xlsx'], 
        key='contens_arcm',
        help='File yang berisi data contents/detail'
    )

with col12:
    file_awal_arcm = st.file_uploader(
        'Upload A/R Credit Memo Base File:', 
        type=['xlsx'], 
        key='awal_arcm',
        help='File yang berisi data awal/header'
    )

if file_contens_arcm and file_awal_arcm:
    try:
        with st.spinner('Process A/R Credit Memo...'):
            # Baca file
            df_contens_arcm = pd.read_excel(file_contens_arcm)
            df_awal_arcm = pd.read_excel(file_awal_arcm)
            
            # Simpan jumlah baris awal
            original_awal_arcm_len = len(df_awal_arcm)
            
            # Merge data dengan logika yang sama seperti AR Invoice
            df_merged_arcm, original_contens_arcm_len = merge_ap_invoice(
                df_contens_arcm, df_awal_arcm, apply_filter=True  # Selalu filter
            )
            
            # Tampilkan info
            st.success('Success!')
            
            col_info16, col_info17, col_info18 = st.columns(3)
            with col_info16:
                st.metric('Baris Contents', original_contens_arcm_len)
            with col_info17:
                st.metric('Baris Awal', original_awal_arcm_len)
            with col_info18:
                st.metric('Baris Hasil', len(df_merged_arcm))
            
            # Preview data
            st.subheader('Preview Result Data:')
            st.dataframe(df_merged_arcm.head(10), use_container_width=True)
            
            # Ekstrak tahun dari Posting Date untuk nama file
            years_in_data_arcm = []
            if 'Posting Date' in df_merged_arcm.columns:
                # Convert ke datetime dan ekstrak tahun
                posting_dates_arcm = pd.to_datetime(df_merged_arcm['Posting Date'], errors='coerce')
                years_in_data_arcm = sorted(posting_dates_arcm.dt.year.dropna().unique().astype(int).tolist())
            
            # Buat nama file dengan tahun
            if years_in_data_arcm:
                if len(years_in_data_arcm) == 1:
                    year_str_arcm = str(years_in_data_arcm[0])
                else:
                    year_str_arcm = f"{years_in_data_arcm[0]}-{years_in_data_arcm[-1]}"
                file_name_arcm = f'AR CREDIT MEMO_{year_str_arcm}.xlsx'
            else:
                # Fallback jika tidak ada tahun
                timestamp_arcm = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name_arcm = f'AR CREDIT MEMO_{timestamp_arcm}.xlsx'
            
            # Button download
            excel_data_arcm = to_excel(df_merged_arcm)
            
            st.download_button(
                label='📥 Download A/R Credit Memo Data',
                data=excel_data_arcm,
                file_name=file_name_arcm,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary'
            )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.info('💡 Make sure both files have a "Doc Number" column')

st.markdown('---')

# Section 7: A/R Reserve Invoice
st.header('🔖 A/R Reserve Invoice')

col_res1, col_res2 = st.columns(2)

with col_res1:
    file_contens_arres = st.file_uploader(
        'Upload A/R Reserve Invoice Contents File:', 
        type=['xlsx'], 
        key='contens_arres',
        help='File yang berisi data contents/detail'
    )

with col_res2:
    file_awal_arres = st.file_uploader(
        'Upload A/R Reserve Invoice Base File:', 
        type=['xlsx'], 
        key='awal_arres',
        help='File yang berisi data awal/header'
    )

if file_contens_arres and file_awal_arres:
    try:
        with st.spinner('Process A/R Reserve Invoice...'):
            # Baca file
            df_contens_arres = pd.read_excel(file_contens_arres)
            df_awal_arres = pd.read_excel(file_awal_arres)
            
            # Simpan jumlah baris awal
            original_awal_arres_len = len(df_awal_arres)
            
            # Merge data dengan logika yang sama seperti AR Invoice
            df_merged_arres, original_contens_arres_len = merge_ap_invoice(
                df_contens_arres, df_awal_arres, apply_filter=True  # Selalu filter
            )
            
            # Tampilkan info
            st.success('Success!')
            
            col_info_res1, col_info_res2, col_info_res3 = st.columns(3)
            with col_info_res1:
                st.metric('Baris Contents', original_contens_arres_len)
            with col_info_res2:
                st.metric('Baris Awal', original_awal_arres_len)
            with col_info_res3:
                st.metric('Baris Hasil', len(df_merged_arres))
            
            # Preview data
            st.subheader('Preview Result Data:')
            st.dataframe(df_merged_arres.head(10), use_container_width=True)
            
            # Ekstrak tahun dari Posting Date untuk nama file
            years_in_data_arres = []
            if 'Posting Date' in df_merged_arres.columns:
                # Convert ke datetime dan ekstrak tahun
                posting_dates_arres = pd.to_datetime(df_merged_arres['Posting Date'], errors='coerce')
                years_in_data_arres = sorted(posting_dates_arres.dt.year.dropna().unique().astype(int).tolist())
            
            # Buat nama file dengan tahun
            if years_in_data_arres:
                if len(years_in_data_arres) == 1:
                    year_str_arres = str(years_in_data_arres[0])
                else:
                    year_str_arres = f"{years_in_data_arres[0]}-{years_in_data_arres[-1]}"
                file_name_arres = f'AR RESERVE_{year_str_arres}.xlsx'
            else:
                # Fallback jika tidak ada tahun
                timestamp_arres = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name_arres = f'AR RESERVE_{timestamp_arres}.xlsx'
            
            # Button download
            excel_data_arres = to_excel(df_merged_arres)
            
            st.download_button(
                label='📥 Download A/R Reserve Invoice Data',
                data=excel_data_arres,
                file_name=file_name_arres,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary'
            )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.info('💡 Make sure both files have a "Doc Number" column')


st.markdown('---')
st.caption('💡 Make sure both files have a "Doc Number" column for data merging')