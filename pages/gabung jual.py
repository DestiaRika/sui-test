import streamlit as st
import pandas as pd
import re
from io import BytesIO

# --- Fungsi utilitas untuk konversi ke Excel ---
@st.cache_data
def convert_df_to_excel(df):
    """Mengonversi DataFrame ke format Excel (bytes) untuk di-download."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data Penjualan Gabungan')
    return output.getvalue()

# --- Fungsi untuk membersihkan .0 dari string ---
def clean_decimal_str(val):
    """Menghapus .0 dari string angka"""
    val_str = str(val).strip()
    # Hapus .0 di akhir
    val_str = re.sub(r'\.0+$', '', val_str)
    return val_str

# --- Konfigurasi Halaman Streamlit ---
st.set_page_config('Gabung Jual', layout='wide')
st.title('📊 Gabung Jual - Data Processing Pipeline')
st.markdown("---")

# ==================== BAGIAN UPLOAD FILE (DI TENGAH) ====================
st.header("📁 1. Upload Semua File Excel")

ar_invoice_23 = st.file_uploader('AR Invoice Data:', type=['xlsx'], key='arin')
do_23 = st.file_uploader('DO Data:', type=['xlsx'], key='do')
retur = st.file_uploader('Retur Data:', type=['xlsx'], key='retur')
timbangan_jual_2023 = st.file_uploader('Timbangan Jual Data:', type=['xlsx'], key='timbangan')
ardp = st.file_uploader('ARDP Data:', type=['xlsx'], key='ardp')
arcm = st.file_uploader('ARCM Data:', type=['xlsx'], key='arcm')
so = st.file_uploader('SO Data:', type=['xlsx'], key='so')
customer = st.file_uploader('Customer Data:', type=['xlsx'], key='customer')
ar_reserve_23 = st.file_uploader('AR Reserve Data:', type=['xlsx'], key='arreserve')

uploaded_files = all(x is not None for x in [ar_invoice_23, do_23, retur, timbangan_jual_2023, ardp, arcm, so, customer, ar_reserve_23])

if not uploaded_files:
    st.info("Silakan upload semua 9 file yang diperlukan untuk memulai proses.")
else:
    st.success("✅ Semua file berhasil di-upload. Memulai proses...")
    st.markdown("---")

    # ==================== PERSIAPAN DATA ====================
    with st.spinner("Membaca file Excel dan menyimpan urutan kolom asli..."):
        df_ar_invoice_orig = pd.read_excel(ar_invoice_23)
        cols_arin_orig = df_ar_invoice_orig.columns.tolist()

        df_do_orig = pd.read_excel(do_23)
        cols_do_orig = df_do_orig.columns.tolist()

        df_return_orig = pd.read_excel(retur)
        cols_return_orig = df_return_orig.columns.tolist()

        df_timbangan_jual_orig = pd.read_excel(timbangan_jual_2023)
        cols_dl_orig = df_timbangan_jual_orig.columns.tolist()

        df_ardp_orig = pd.read_excel(ardp)
        cols_ardp_orig = df_ardp_orig.columns.tolist()

        df_arcm_orig = pd.read_excel(arcm)
        cols_arcm_orig = df_arcm_orig.columns.tolist()

        df_so_orig = pd.read_excel(so)
        cols_so_orig = df_so_orig.columns.tolist()

        df_customer_orig = pd.read_excel(customer)
        cols_cmd_orig = df_customer_orig.columns.tolist()
        
        df_ar_reserve_orig = pd.read_excel(ar_reserve_23)
        
    # ==================== BAGIAN 1: PROSES df_dar_ada_233 ====================
    st.header("🔷 BAGIAN 1: Proses df_dar_ada_233")
    
    with st.spinner("Loading files..."):
        df_ar_invoice = pd.read_excel(ar_invoice_23)
        df_do = pd.read_excel(do_23)
        df_return = pd.read_excel(retur)
        df_timbangan_jual = pd.read_excel(timbangan_jual_2023)
        df_ardp = pd.read_excel(ardp)
        df_arcm = pd.read_excel(arcm)
        df_so = pd.read_excel(so)
        df_customer = pd.read_excel(customer)
        st.success("✅ Semua file berhasil di-load")
    
    # === ARIN Processing ===
    st.subheader("1️⃣ AR Invoice - Extract DO Numbers")
    df_ar_invoice.columns = df_ar_invoice.columns.str.strip()
    df_do.columns = df_do.columns.str.strip()
    
    def extract_do_numbers(remarks):
        if pd.isna(remarks):
            return []
        matches = re.findall(r"Deliveries[^\d]*(.*)", str(remarks))
        if not matches:
            return []
        numbers = re.findall(r"\d{8}", matches[0])
        return numbers
    
    df_ar_invoice["DO_List"] = df_ar_invoice["Remarks"].apply(extract_do_numbers)
    
    important_cols = ['Doc Number', 'Customer Code', 'Item No', 'Bag Quantity', 'Quantity', 'Unit Price', 'Whse']
    for col in important_cols:
        if col in df_do.columns:
            df_do[col] = df_do[col].apply(clean_decimal_str)
        if col in df_ar_invoice.columns:
            df_ar_invoice[col] = df_ar_invoice[col].apply(clean_decimal_str)
    
    used_dos = set()
    
    def find_matching_do(row):
        do_list = row["DO_List"]
        if not do_list:
            return None
        if len(do_list) == 1:
            do_num = do_list[0]
            used_dos.add(do_num)
            return do_num
        for do_num in do_list:
            if do_num in used_dos:
                continue
            do_row = df_do[df_do["Doc Number"] == do_num]
            if do_row.empty:
                continue
            do_row = do_row.iloc[0]
            same = True
            for col in important_cols:
                if col == "Doc Number":
                    continue
                val_ar = clean_decimal_str(row.get(col, ""))
                val_do = clean_decimal_str(do_row.get(col, ""))
                if val_ar != val_do:
                    same = False
                    break
            if same:
                used_dos.add(do_num)
                return do_num
        return None
    
    df_ar_invoice["Base On DO"] = df_ar_invoice.apply(find_matching_do, axis=1)
    
    if "Remarks" in df_ar_invoice.columns:
        df_ar_invoice["Remarks"] = df_ar_invoice["Remarks"].astype(str).apply(
            lambda x: re.sub(r"Delivery Order\s+(\d{8})", r"Deliveries \1", x)
        )
    
    df_ar_invoice.drop(columns=["DO_List"], inplace=True)
    st.success(f"✅ Base On DO extracted: {df_ar_invoice['Base On DO'].notna().sum()} matches")
    
    # === ARIN - DO Merge ===
    st.subheader("2️⃣ Merge ARIN dengan DO")
    df_do['Doc Number'] = df_do['Doc Number'].apply(clean_decimal_str)
    df_ar_invoice['Base On DO'] = df_ar_invoice['Base On DO'].apply(clean_decimal_str)
    
    df_do_suffixed = df_do.add_suffix('_DO')
    df_ar_suffixed = df_ar_invoice.add_suffix('_ARIN')
    
    df_merged = pd.merge(df_do_suffixed, df_ar_suffixed, how='left', left_on='Doc Number_DO', right_on='Base On DO_ARIN')
    st.success(f"✅ ARIN-DO merged: {len(df_merged)} rows")
    
    # === RETURN Processing ===
    st.subheader("3️⃣ Return - Extract DO")

    df_return.columns = df_return.columns.str.strip()

    def extract_and_clean_do(remarks):
        if pd.isna(remarks):
            return None
        match = re.search(r'Based On Deliveries[^\d]*(\d{8})', str(remarks))
        if match:
            # Ambil 8 digit, hapus semua karakter non-digit, lalu ambil 8 karakter pertama
            do_number = re.sub(r'\D', '', match.group(1))[:8]
            return do_number
        return None

    df_return['Base on DO'] = df_return['Remarks'].apply(extract_and_clean_do)

    # Bersihkan lagi jika masih ada titik atau karakter aneh
    df_return['Base on DO'] = df_return['Base on DO'].apply(
        lambda x: re.sub(r'\D', '', str(x))[:8] if pd.notna(x) and x else None
    )

    st.success(f"✅ Base on DO extracted from Return: {df_return['Base on DO'].notna().sum()} matches")

    st.subheader("4️⃣ Merge dengan Return")

    # Bersihkan Doc Number di df_return juga
    df_return['Base on DO'] = df_return['Base on DO'].apply(clean_decimal_str)

    # Suffix semua kolom kecuali kolom untuk merge
    df_return_suffixed = df_return.add_suffix('_RETURN')
    # Rename kembali kolom merge agar bisa di-merge
    df_return_suffixed = df_return_suffixed.rename(columns={'Base on DO_RETURN': 'Base on DO_RETURN_temp'})

    # Merge berdasarkan DO Number
    df_merged = pd.merge(
        df_merged, 
        df_return_suffixed, 
        how="left", 
        left_on='Doc Number_DO',  # dari df_merged (hasil merge DO-ARIN)
        right_on='Base on DO_RETURN_temp'  # dari df_return
    )

    # Rename kembali untuk konsistensi
    df_merged = df_merged.rename(columns={'Base on DO_RETURN_temp': 'Base on DO_RETURN'})

    # Urutkan kolom
    do_cols = [c for c in df_merged.columns if c.endswith("_DO")]
    return_cols = [c for c in df_merged.columns if c.endswith("_RETURN")]
    arin_cols = [c for c in df_merged.columns if c.endswith("_ARIN")]
    other_cols = [c for c in df_merged.columns if not (c.endswith("_DO") or c.endswith("_RETURN") or c.endswith("_ARIN"))]

    df_merged = df_merged[do_cols + return_cols + arin_cols + other_cols]
    st.success(f"✅ Return merged: {len(df_merged)} rows")
    
    # === Filter Canceled ===
    st.subheader("5️⃣ Filter Canceled")
    df_merged = df_merged[(df_merged['Canceled_DO'] != 'Yes') & (df_merged['Canceled_DO'] != 'Cancellation')]
    st.success(f"✅ After filter: {len(df_merged)} rows")
    
    # === Merge dengan Timbangan (DL) - TAHAP 1: Merge Normal ===
    st.subheader("6️⃣ Merge dengan Timbangan (DL)")
    
    # Baca ulang file timbangan asli untuk matching alternatif
    df_timbangan_raw = pd.read_excel(timbangan_jual_2023)
    df_timbangan_raw.columns = df_timbangan_raw.columns.str.strip()
    
    df_timbangan_suffixed = df_timbangan_jual.rename(columns={col: f"{col}_DL" for col in df_timbangan_jual.columns})
    df_timbangan_dedup = df_timbangan_suffixed.drop_duplicates(subset=['Base Number_DL'], keep='first')
    
    # Bersihkan .0 pada kolom kunci sebelum merge
    df_merged['Instruction Number_DO'] = df_merged['Instruction Number_DO'].apply(clean_decimal_str)
    df_timbangan_dedup['Base Number_DL'] = df_timbangan_dedup['Base Number_DL'].apply(clean_decimal_str)
    
    df_merged = pd.merge(df_merged, df_timbangan_dedup, how="left", left_on="Instruction Number_DO", right_on="Base Number_DL")
    
    # --- Conditional Clear jika Qty_DL, Weight 1_DL, DAN Weight Difference_DL = 0 ---
    required_dl_cols = ['Qty_DL', 'Weight 1_DL', 'Weight Difference_DL']
    
    if all(col in df_merged.columns for col in required_dl_cols):
        mask_all_zero = (
            (df_merged['Qty_DL'].fillna(0) == 0) &
            (df_merged['Weight 1_DL'].fillna(0) == 0) &
            (df_merged['Weight Difference_DL'].fillna(0) == 0)
        )
        
        dl_cols_to_clear = [col for col in df_merged.columns if col.endswith('_DL')]
        
        if mask_all_zero.any():
            df_merged.loc[mask_all_zero, dl_cols_to_clear] = None
            st.info(f"ℹ️ {mask_all_zero.sum()} baris data _DL telah dikosongkan karena 'Qty_DL', 'Weight 1_DL', dan 'Weight Difference_DL' KETIGANYA bernilai 0.")
    
    st.success(f"✅ Merge pertama selesai: {df_merged['Base Number_DL'].notna().sum()} baris berhasil")
    
    # === TAHAP 2: Matching Alternatif untuk KRGDB ===
    st.subheader("6️⃣.1 Matching Alternatif untuk KRGDB (Instruction Number Kosong)")

    import re
    import pandas as pd

    # === Fungsi helper ===

    def is_empty_or_zero(series):
        """Mengecek apakah nilai dalam series adalah kosong, 0, atau NaN"""
        return (
            series.isna() |
            (series.astype(str).str.strip() == '') |
            (series.astype(str).str.strip() == '0') |
            (series.astype(str).str.strip() == '0.0') |
            (series == 0) |
            (series.astype(str).str.lower() == 'nan')
        )

    def format_date_only(date_val):
        """Mengkonversi tanggal ke format dd.mm.yyyy tanpa time"""
        if pd.isna(date_val):
            return ''
        try:
            # pastikan bisa baca berbagai format (/, ., -)
            date_obj = pd.to_datetime(str(date_val).replace('.', '/').replace('-', '/'), 
                                      errors='coerce', dayfirst=True)
            if pd.notna(date_obj):
                return date_obj.strftime('%d.%m.%Y')  # <-- gunakan titik (.)
            else:
                return ''
        except:
            return ''

    def clean_item_code(val):
        """Membersihkan Item Number/Code"""
        if pd.isna(val):
            return ''
        val_str = str(val).strip()
        if val_str.endswith('.0'):
            val_str = val_str[:-2]
        return val_str.upper()

    def aggressive_clean(val):
        """Membersihkan string dengan lebih agresif"""
        if pd.isna(val):
            return ''
        val_str = str(val).strip()
        val_str = re.sub(r'\s+', ' ', val_str)
        val_str = val_str.replace('\t', '').replace('\n', '').replace('\r', '')
        return val_str.upper()

    # === Identifikasi baris yang perlu matching alternatif ===
    mask_need_alternative = (
        (df_merged['Whse_DO'] == 'KRGDB') &
        (df_merged['Base Number_DL'].isna()) &
        is_empty_or_zero(df_merged['Instruction Number_DO'])
    )

    count_alternative = mask_need_alternative.sum()

    if count_alternative > 0:
        st.info(f"ℹ️ {count_alternative} baris 'KRGDB' dengan Instruction Number kosong/0/NaN akan dicoba matching dengan key alternatif.")
        
        # Ambil baris dari df_merged yang perlu di-match
        df_need_match = df_merged[mask_need_alternative].copy().reset_index(drop=False)
        df_need_match.rename(columns={'index': 'original_index'}, inplace=True)
        
        # Definisikan kolom untuk merged_key
        do_keys = ['Customer Name_DO', 'Posting Date_DO', 'Item No_DO', 'Quantity_DO', 'License Number_DO']
        timb_keys = ['Customer Name', 'Out Date', 'Item Code', 'Weight Difference', 'Nopol']
        
        # Cek kolom
        missing_do_keys = [k for k in do_keys if k not in df_need_match.columns]
        missing_timb_keys = [k for k in timb_keys if k not in df_timbangan_raw.columns]
        
        if missing_do_keys:
            st.error(f"❌ Kolom DO tidak ditemukan: {missing_do_keys}")
        if missing_timb_keys:
            st.error(f"❌ Kolom Timbangan tidak ditemukan: {missing_timb_keys}")
            st.write("Kolom yang tersedia di Timbangan:", df_timbangan_raw.columns.tolist())
        
        if not missing_do_keys and not missing_timb_keys:
            # Bersihkan kolom DO
            for col in do_keys:
                if col == 'Posting Date_DO':
                    df_need_match[col] = df_need_match[col].apply(format_date_only)
                elif col == 'Item No_DO':
                    df_need_match[col] = df_need_match[col].apply(clean_item_code)
                else:
                    df_need_match[col] = df_need_match[col].apply(aggressive_clean)
            
            # Bersihkan kolom Timbangan RAW
            for col in timb_keys:
                if col == 'Out Date':
                    # ubah ke datetime lalu format dd.mm.yyyy
                    df_timbangan_raw[col] = pd.to_datetime(df_timbangan_raw[col], errors='coerce', dayfirst=True)
                    df_timbangan_raw[col] = df_timbangan_raw[col].apply(format_date_only)
                elif col == 'Item Code':
                    df_timbangan_raw[col] = df_timbangan_raw[col].apply(clean_item_code)
                else:
                    df_timbangan_raw[col] = df_timbangan_raw[col].apply(aggressive_clean)
            
            # Buat merged_key
            df_need_match['merged_key_alt'] = df_need_match[do_keys].apply(
                lambda row: '_'.join(row.astype(str)), axis=1
            )
            
            df_timbangan_raw['merged_key_alt'] = df_timbangan_raw[timb_keys].apply(
                lambda row: '_'.join(row.astype(str)), axis=1
            )
            
            st.success("✅ Key alternatif berhasil dibuat untuk matching dengan format tanggal dd.mm.yyyy!")

            # lanjut proses matching dll (kode selanjutnya tetap sama)

            
            # === MATCHING & UPDATE DATA ===
            st.write("### 🔗 Proses Matching Data")
            
            # Deduplikasi Timbangan
            df_timbangan_dedup_alt = df_timbangan_raw.drop_duplicates(subset=['merged_key_alt'], keep='first')
            
            # Merge berdasarkan key - LEFT JOIN untuk dapat semua data DO
            df_all_merge = pd.merge(
                df_need_match,
                df_timbangan_dedup_alt,
                on='merged_key_alt',
                how='left',
                suffixes=('', '_timb'),
                indicator=True
            )
            
            # Pisahkan yang match dan tidak match
            df_matched = df_all_merge[df_all_merge['_merge'] == 'both'].copy()
            df_not_matched = df_all_merge[df_all_merge['_merge'] == 'left_only'].copy()
            
            count_matched = len(df_matched)
            count_not_matched = len(df_not_matched)
            
            # Tampilkan Summary
            st.write("#### 📊 Ringkasan Matching")
            col_summary1, col_summary2, col_summary3 = st.columns(3)
            with col_summary1:
                st.metric("✅ Berhasil Match", count_matched)
            with col_summary2:
                st.metric("❌ Tidak Match", count_not_matched)
            with col_summary3:
                total = count_matched + count_not_matched
                pct = (count_matched / total * 100) if total > 0 else 0
                st.metric("📈 Persentase Match", f"{pct:.1f}%")
            
            st.divider()
            
            # === TAMPILKAN YANG BERHASIL MATCH ===
            st.write("#### 🔑 Merged Key yang Berhasil Match")
            
            if count_matched > 0:
                st.success(f"✅ Berhasil match {count_matched} baris!")
                
                # Buat DataFrame untuk merged keys
                df_keys_do_match = df_matched[['merged_key_alt', 'Customer Name_DO', 'Posting Date_DO', 
                                        'Item No_DO', 'Quantity_DO', 'License Number_DO']].copy()
                df_keys_do_match.columns = ['Merged Key', 'Customer', 'Posting Date', 'Item No', 'Quantity', 'License']
                
                df_keys_timb_match = df_matched[['merged_key_alt', 'Customer Name', 'Out Date', 
                                        'Item Code', 'Weight Difference', 'Nopol']].copy()
                df_keys_timb_match.columns = ['Merged Key', 'Customer', 'Out Date', 'Item Code', 'Weight Diff', 'Nopol']
                
                # Tampilkan side-by-side
                col_frame1, col_frame2 = st.columns(2)
                
                with col_frame1:
                    st.write("**📋 Merged Key dari DO (MATCH):**")
                    st.dataframe(df_keys_do_match, height=400, use_container_width=True)
                    st.caption(f"Total: {len(df_keys_do_match)} baris")
                
                with col_frame2:
                    st.write("**🏭 Merged Key dari Timbangan (MATCH):**")
                    st.dataframe(df_keys_timb_match, height=400, use_container_width=True)
                    st.caption(f"Total: {len(df_keys_timb_match)} baris")
                
                # Update df_merged dengan data yang match
                timb_columns = [col for col in df_timbangan_dedup_alt.columns if col != 'merged_key_alt']
                column_mapping = {col: f"{col}_DL" for col in timb_columns}
                
                updated_rows = 0
                for idx, row in df_matched.iterrows():
                    original_idx = row['original_index']
                    for timb_col, dl_col in column_mapping.items():
                        if timb_col in df_matched.columns and dl_col in df_merged.columns:
                            df_merged.at[original_idx, dl_col] = row[timb_col]
                    updated_rows += 1
                
                st.success(f"✅ {updated_rows} baris berhasil diupdate dengan data Timbangan!")
            else:
                st.warning("⚠️ Tidak ada data yang berhasil di-match dengan key alternatif.")
            
            st.divider()
            
            # === TAMPILKAN YANG TIDAK MATCH ===
            st.write("#### ❌ Merged Key yang TIDAK Match")
            
            if count_not_matched > 0:
                st.warning(f"⚠️ {count_not_matched} baris tidak menemukan pasangan di Timbangan")
                
                df_keys_do_not_match = df_not_matched[['merged_key_alt', 'Customer Name_DO', 'Posting Date_DO', 
                                        'Item No_DO', 'Quantity_DO', 'License Number_DO']].copy()
                df_keys_do_not_match.columns = ['Merged Key', 'Customer', 'Posting Date', 'Item No', 'Quantity', 'License']
                
                st.write("**📋 Merged Key dari DO (TIDAK MATCH):**")
                st.dataframe(df_keys_do_not_match, height=400, use_container_width=True)
                st.caption(f"Total: {len(df_keys_do_not_match)} baris tidak menemukan data di Timbangan")
            else:
                st.success("🎉 Semua data berhasil match!")

                        # === TAMPILKAN DATA TIMBANGAN YANG TIDAK MATCH (RIGHT ONLY) ===
            st.write("#### ❌ Merged Key dari Timbangan yang TIDAK Match")
            
            # Dapatkan semua merged_key_alt di timbangan
            timb_keys_all = set(df_timbangan_dedup_alt['merged_key_alt'].dropna().unique())
            # Dapatkan semua merged_key_alt yang sudah match
            matched_keys = set(df_matched['merged_key_alt'].dropna().unique())
            # Ambil sisa yang tidak match
            unmatched_timb_keys = timb_keys_all - matched_keys
            
            if unmatched_timb_keys:
                df_timb_not_match = df_timbangan_dedup_alt[df_timbangan_dedup_alt['merged_key_alt'].isin(unmatched_timb_keys)].copy()
                
                st.warning(f"⚠️ {len(df_timb_not_match)} data dari Timbangan tidak menemukan pasangan di DO (unmatched).")
                
                # Pilih kolom utama biar rapi
                cols_show = ['merged_key_alt', 'Customer Name', 'Out Date', 'Item Code', 'Weight Difference', 'Nopol']
                existing_cols = [c for c in cols_show if c in df_timb_not_match.columns]
                st.dataframe(df_timb_not_match[existing_cols], height=400, use_container_width=True)
                
                st.caption("Data di atas menunjukkan semua 'merged_key_alt' dari timbangan yang tidak punya pasangan DO.")
            else:
                st.success("🎉 Semua data timbangan sudah memiliki pasangan DO!")

            
            st.divider()
    else:
        st.info("ℹ️ Tidak ada baris KRGDB dengan Instruction Number kosong yang memerlukan matching alternatif.")

    st.subheader("7️⃣ Extract SO Number")
    def extract_sales_order(remarks):
        if pd.isna(remarks):
            return None
        match = re.search(r'Sales Orders (\d{8})', str(remarks))
        if match:
            return match.group(1)
        return None
    
    if 'Remarks_DO' in df_merged.columns:
        df_merged['Base On_SO'] = df_merged['Remarks_DO'].apply(extract_sales_order)
    
    df_merged = df_merged.sort_values(by="Base On_SO", ascending=True, na_position="last").reset_index(drop=True)
    st.success(f"✅ SO extracted: {df_merged['Base On_SO'].notna().sum()} matches")
    
    # === Process ARDP ===
    st.subheader("8️⃣ Process ARDP")
    def extract_so_number(details):
        if pd.isna(details):
            return None
        match = re.search(r'Sales Orders (\d{8})', str(details))
        if match:
            return match.group(1)
        match = re.search(r'SOs?\s*(\d{8})', str(details))
        if match:
            return match.group(1)
        return None
    
    if 'Details' in df_ardp.columns:
        df_ardp['Base On SO'] = df_ardp['Details'].apply(extract_so_number)
    
    # === Process ARCM ===
    st.subheader("9️⃣ Process ARCM")
    def extract_so(text):
        if pd.isna(text):
            return None
        match = re.search(r'Sales Orders (\d{8})', str(text))
        return match.group(1) if match else None
    
    def extract_do(text):
        if pd.isna(text):
            return None
        match = re.search(r'Deliveries (\d{8})', str(text))
        return match.group(1) if match else None
    
    def extract_ardp(text):
        if pd.isna(text):
            return None
        match = re.search(r'A/R Down Payment (\d{8})', str(text))
        return match.group(1) if match else None
    
    if 'Remarks' in df_arcm.columns:
        df_arcm['Base On SO'] = df_arcm['Remarks'].apply(extract_so)
        df_arcm['Base On DO'] = df_arcm['Remarks'].apply(extract_do)
        df_arcm['Base On ARDP'] = df_arcm['Remarks'].apply(extract_ardp)
    
    # === Merge ARCM dengan ARDP ===
    st.subheader("🔟 Merge ARCM dengan ARDP")
    df_arcm_suffixed = df_arcm.add_suffix('_CM')
    df_ardp_suffixed = df_ardp.add_suffix('_ARDP')
    
    df_arcm_suffixed['Base On ARDP_CM'] = df_arcm_suffixed['Base On ARDP_CM'].apply(clean_decimal_str)
    df_ardp_suffixed['Doc Number_ARDP'] = df_ardp_suffixed['Doc Number_ARDP'].apply(clean_decimal_str)
    
    df_cm_ardp = pd.merge(df_arcm_suffixed, df_ardp_suffixed, left_on='Base On ARDP_CM', right_on='Doc Number_ARDP', how='outer')
    
    # === Merge dengan SO ===
    st.subheader("1️⃣1️⃣ Merge dengan SO")
    df_so_suffixed = df_so.add_suffix('_SO')
    df_cm_ardp['Base On SO_ARDP'] = df_cm_ardp['Base On SO_ARDP'].apply(clean_decimal_str)
    df_so_suffixed['Doc Number_SO'] = df_so_suffixed['Doc Number_SO'].apply(clean_decimal_str)
    
    df_so_cm_ardp = pd.merge(df_cm_ardp, df_so_suffixed, left_on='Base On SO_ARDP', right_on='Doc Number_SO', how='outer')
    
    # Urutkan kolom
    all_cols = df_so_cm_ardp.columns.tolist()
    suffix_order = ['_SO', '_ARDP', '_CM']
    ordered_cols = []
    remaining_cols = all_cols.copy()
    
    for suffix in suffix_order:
        current_suffix_cols = [col for col in remaining_cols if col.endswith(suffix)]
        ordered_cols.extend(current_suffix_cols)
        remaining_cols = [col for col in remaining_cols if col not in current_suffix_cols]
    
    ordered_cols.extend(remaining_cols)
    df_so_cm_ardp = df_so_cm_ardp[ordered_cols]
    
    # === Merge df_dar_ada dengan df_so_cm_ardp ===
    st.subheader("1️⃣2️⃣ Merge DAR_ADA dengan SO_CM_ARDP")
    df_so_cm_ardp_dedup = df_so_cm_ardp.drop_duplicates(subset=['Doc Number_SO'], keep='first')
    
    # Bersihkan .0 pada kolom kunci
    df_merged['Base On_SO'] = df_merged['Base On_SO'].apply(clean_decimal_str)
    df_so_cm_ardp_dedup['Doc Number_SO'] = df_so_cm_ardp_dedup['Doc Number_SO'].apply(clean_decimal_str)
    
    df_dar_ada_233 = pd.merge(df_merged, df_so_cm_ardp_dedup, how='left', left_on='Base On_SO', right_on='Doc Number_SO')
    
    # Urutkan kolom final
    all_cols = df_dar_ada_233.columns.tolist()
    so_cols = [col for col in all_cols if col.endswith('_SO')]
    ardp_cols = [col for col in all_cols if col.endswith('_ARDP')]
    cm_cols = [col for col in all_cols if col.endswith('_CM')]
    dl_cols = [col for col in all_cols if col.endswith('_DL')]
    do_cols = [col for col in all_cols if col.endswith('_DO')]
    return_cols = [col for col in all_cols if col.endswith('_RETURN')]
    arin_cols = [col for col in all_cols if col.endswith('_ARIN')]
    suffixed_cols = set(so_cols + ardp_cols + cm_cols + dl_cols + do_cols + return_cols + arin_cols)
    other_cols = [col for col in all_cols if col not in suffixed_cols]
    
    df_dar_ada_233 = df_dar_ada_233[so_cols + ardp_cols + cm_cols + dl_cols + do_cols + return_cols + arin_cols + other_cols]
    st.success(f"✅ df_dar_ada_233 selesai: {len(df_dar_ada_233)} rows")
    
    st.markdown("---")
    
    # ==================== BAGIAN 2: PROSES df_final ====================
    st.header("🔶 BAGIAN 2: Proses df_final (DO + AR Reserve + Timbangan + SO)")

    st.subheader("1️⃣ Extract Invoice Number dari DO")
    df_do_final = pd.read_excel(do_23)

    def extract_invoice_number(text):
        if pd.isna(text):
            return None
        match = re.search(r"Based On A/R Invoices\s+(\d+)", str(text))
        return match.group(1) if match else None

    df_do_final["Base On Invoice"] = df_do_final["Remarks"].apply(extract_invoice_number)
    st.success(f"✅ Invoice extracted: {df_do_final['Base On Invoice'].notna().sum()} matches")

    st.subheader("2️⃣ Merge DO dengan AR Reserve")
    df_ar_reserve = pd.read_excel(ar_reserve_23)

    df_do_final['Base On Invoice'] = df_do_final['Base On Invoice'].apply(clean_decimal_str)
    df_ar_reserve['Doc Number'] = df_ar_reserve['Doc Number'].apply(clean_decimal_str)

    df_do_suffixed = df_do_final.add_suffix("_DO")
    df_ar_suffixed = df_ar_reserve.add_suffix("_ARIN")

    df_do_ar_merged = pd.merge(df_do_suffixed, df_ar_suffixed, how="left", left_on="Base On Invoice_DO", right_on="Doc Number_ARIN")
    df_do_ar_matched = df_do_ar_merged.dropna(subset=["Doc Number_ARIN"])
    st.success(f"✅ AR Reserve merged: {len(df_do_ar_matched)} rows")

    st.subheader("3️⃣ Prepare Timbangan dengan Merged Key")
    df_dl = pd.read_excel(timbangan_jual_2023)
    df_dl.columns = df_dl.columns.str.strip()
    df_do_ar_matched.columns = df_do_ar_matched.columns.str.strip()

    cols_dl = ["Customer Name", "Item Code", "Nopol", "Weight Difference"]
    if all(col in df_dl.columns for col in cols_dl):
        df_dl["merged_key"] = df_dl[cols_dl].apply(lambda row: "_".join([clean_decimal_str(x) for x in row]), axis=1)

    cols_do = ["Customer Name_DO", "Item No_DO", "License Number_DO", "Quantity_DO"]
    if all(col in df_do_ar_matched.columns for col in cols_do):
        df_do_ar_matched["merged_key"] = df_do_ar_matched[cols_do].apply(lambda row: "_".join([clean_decimal_str(x) for x in row]), axis=1)

    st.subheader("4️⃣ Merge dengan Timbangan")
    df_dl = df_dl.drop_duplicates(subset=["merged_key"], keep="first")
    df_dl = df_dl.rename(columns={col: f"{col}_DL" for col in df_dl.columns if col != "merged_key"})

    df_do_dl_merged = pd.merge(df_do_ar_matched, df_dl, on="merged_key", how="left")

    cols_dl = [c for c in df_do_dl_merged.columns if c.endswith("_DL")]
    cols_do = [c for c in df_do_dl_merged.columns if c.endswith("_DO")]
    cols_arin = [c for c in df_do_dl_merged.columns if c.endswith("_ARIN")]
    cols_other = [c for c in df_do_dl_merged.columns if not (c.endswith("_DL") or c.endswith("_DO") or c.endswith("_ARIN"))]

    df_do_dl_merged = df_do_dl_merged[cols_dl + cols_do + cols_arin + cols_other]
    st.success(f"✅ DL merged: {len(df_do_dl_merged)} rows")

    st.subheader("5️⃣ Extract SO Number")
    def extract_so_number_arin(text):
        if pd.isna(text):
            return None
        match = re.search(r"Based On Sales Orders\s+(\d{8})", str(text))
        return match.group(1) if match else None

    if "Remarks_ARIN" in df_do_dl_merged.columns:
        df_do_dl_merged["Base On SO"] = df_do_dl_merged["Remarks_ARIN"].apply(extract_so_number_arin)
        st.success(f"✅ SO extracted: {df_do_dl_merged['Base On SO'].notna().sum()} matches")

    st.subheader("6️⃣ Merge dengan SO")
    df_so_final = pd.read_excel(so)
    df_so_final.columns = df_so_final.columns.str.strip()

    # Bersihkan .0 pada kolom kunci
    df_do_dl_merged['Base On SO'] = df_do_dl_merged['Base On SO'].apply(clean_decimal_str)
    df_so_final['Doc Number'] = df_so_final['Doc Number'].apply(clean_decimal_str)

    # --- AWAL PERBAIKAN KODE ---
    # Ganti nama kolom 'Doc Number' menjadi 'Doc Number_SO' sebelum merge
    # Ini memastikan kolom kunci dari df_so_final memiliki suffix '_SO'
    df_so_final = df_so_final.rename(columns={"Doc Number": "Doc Number_SO"})
    
    # Ganti nama kolom lainnya dengan suffix '_SO'
    df_so_final = df_so_final.rename(columns={col: f"{col}_SO" for col in df_so_final.columns if col != "Doc Number_SO"})

    # Lakukan merge menggunakan 'Doc Number_SO' sebagai kunci
    df_final = pd.merge(df_do_dl_merged, df_so_final, left_on="Base On SO", right_on="Doc Number_SO", how="inner")
    # --- AKHIR PERBAIKAN KODE ---

    cols_so = [c for c in df_final.columns if c.endswith("_SO") or c == "Doc Number_SO"] # Diperbarui untuk menyertakan Doc Number_SO
    cols_dl = [c for c in df_final.columns if c.endswith("_DL")]
    cols_do = [c for c in df_final.columns if "_DO" in c and not c.endswith("_DL")]
    cols_arin = [c for c in df_final.columns if "_ARIN" in c]
    other_cols = [c for c in df_final.columns if c not in cols_so + cols_dl + cols_do + cols_arin]

    df_final = df_final[cols_so + cols_dl + cols_do + cols_arin + other_cols]
    st.success(f"✅ df_final selesai: {len(df_final)} rows")

    # ==================== BAGIAN 3: GABUNGKAN, MERGE & URUTKAN ====================
    st.header("🔗 4. Gabungkan, Merge & Urutkan")

    # Gabungkan dengan mempertahankan semua kolom
    all_cols_233 = set(df_dar_ada_233.columns)
    all_cols_final = set(df_final.columns)
    all_possible_cols = list(all_cols_233.union(all_cols_final))

    # Tambahkan kolom yang tidak ada dengan nilai None
    for col in all_possible_cols:
        if col not in df_dar_ada_233.columns:
            df_dar_ada_233[col] = None
        if col not in df_final.columns:
            df_final[col] = None

    # Pastikan urutan kolom sama
    df_dar_ada_233 = df_dar_ada_233[all_possible_cols]

    # --- AWAL PERBAIKAN KODE (Filter df_dar_ada_233) ---
    # Menghapus baris dari df_dar_ada_233 di mana 'Item Description_SO' adalah 'Corn Grits FGIIIA'.
    FILTER_VALUE = 'Corn Grits FGIIIA'
    FILTER_COL = 'Item Description_SO'

    if FILTER_COL in df_dar_ada_233.columns:
        initial_count = len(df_dar_ada_233)
        
        # Filter: Hanya pertahankan baris di mana kolom TIDAK SAMA dengan FILTER_VALUE
        df_dar_ada_233 = df_dar_ada_233[df_dar_ada_233[FILTER_COL] != FILTER_VALUE]
        
        removed_count = initial_count - len(df_dar_ada_233)
        
        # Menambahkan pesan sukses/informasi untuk Streamlit
        if removed_count > 0:
            st.info(f"ℹ️ Dihapus {removed_count} baris dari df_dar_ada_233 karena '{FILTER_COL}' adalah '{FILTER_VALUE}'.")
        st.success(f"✅ df_dar_ada_233 difilter: {len(df_dar_ada_233)} baris tersisa.")
    else:
        st.warning(f"⚠️ Kolom '{FILTER_COL}' tidak ditemukan di df_dar_ada_233. Filter tidak diterapkan.")
    # --- AKHIR PERBAIKAN KODE ---

    df_final = df_final[all_possible_cols]

    df_combined = pd.concat([df_dar_ada_233, df_final], axis=0, ignore_index=True)
    st.success(f"✅ Data digabungkan: {len(df_combined)} baris")

    col_arin_key = 'Customer Code_ARIN'
    col_do_key = 'Customer Code_DO'
    df_combined['key_customer'] = None
    if col_arin_key in df_combined.columns:
        df_combined['key_customer'] = df_combined[col_arin_key]
    if col_do_key in df_combined.columns:
        df_combined['key_customer'].fillna(df_combined[col_do_key], inplace=True)

    df_customer_orig['Customer Code'] = df_customer_orig['Customer Code'].apply(clean_decimal_str)
    df_combined['key_customer'] = df_combined['key_customer'].apply(clean_decimal_str)
    df_cmd_suffixed = df_customer_orig.add_suffix("_CMD")
    df_combined = pd.merge(df_combined, df_cmd_suffixed, how="left", left_on='key_customer', right_on='Customer Code_CMD')
    df_combined = df_combined.drop(columns=['key_customer'], errors='ignore')
    st.success(f"✅ Customer Master di-merge")

    st.subheader("🔄 Mengurutkan Kolom Sesuai Urutan Asli")
    order_map = {
        '_CMD': cols_cmd_orig, '_SO': cols_so_orig, '_ARDP': cols_ardp_orig,
        '_CM': cols_arcm_orig, '_DL': cols_dl_orig, '_DO': cols_do_orig,
        '_RETURN': cols_return_orig, '_ARIN': cols_arin_orig
    }
    final_ordered_cols = []
    for suffix, original_cols in order_map.items():
        for col_name in original_cols:
            suffixed_col = f"{col_name}{suffix}"
            if suffixed_col in df_combined.columns:
                final_ordered_cols.append(suffixed_col)
    remaining_cols = [col for col in df_combined.columns if col not in final_ordered_cols]
    final_ordered_cols.extend(remaining_cols)
    df_combined = df_combined[final_ordered_cols]
    st.success("✅ Kolom berhasil diurutkan.")
    st.markdown("---")

    # ==================== BAGIAN 4: PEMBERSIHAN DUPLIKAT KHUSUS ====================
    st.header("🧹 5. Membersihkan Duplikat Data ARIN (< 8 Digit)")

    if 'Doc Number_ARIN' in df_combined.columns:
        df_combined['Doc Number_ARIN'] = df_combined['Doc Number_ARIN'].astype(str).replace('nan', '')
        arin_cols_to_clear = [col for col in df_combined.columns if col.endswith('_ARIN')]
        
        # Buat 'mask' untuk menandai baris duplikat yang memenuhi kondisi
        mask_duplicates = (
            (df_combined['Doc Number_ARIN'].str.len() < 8) &
            (df_combined['Doc Number_ARIN'] != '') &
            (df_combined.duplicated(subset=['Doc Number_ARIN'], keep='first'))
        )
        
        if arin_cols_to_clear:
            # Gunakan .loc untuk mengubah nilai pada baris dan kolom yang ditandai
            df_combined.loc[mask_duplicates, arin_cols_to_clear] = None
            
            num_cleaned = mask_duplicates.sum()
            if num_cleaned > 0:
                st.success(f"✅ Berhasil membersihkan data duplikat pada {num_cleaned} baris.")
            else:
                st.info("ℹ️ Tidak ada data duplikat (< 8 digit) yang perlu dibersihkan.")
    else:
        st.warning("⚠️ Kolom 'Doc Number_ARIN' tidak ditemukan, langkah pembersihan duplikat dilewati.")
    st.markdown("---")

    # ==================== BAGIAN DOWNLOAD ====================
    st.header("📥 6. Download Hasil Akhir")

    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total Baris Final", len(df_combined))
    with col2:
        st.metric("Total Kolom Final", len(df_combined.columns))

    # --- Daftar kolom yang ingin dihapus secara eksplisit ---
    cols_to_drop = ["Customer Code_DO", "BP Balance_CMD", "Payment Terms Code_CMD", "#_SO", "Series_SO",
                    "Customer Code_SO", "Customer Name_SO", "Delivery Date_SO", "Customer Due Date_SO",
                    "Item No_SO", "Item Description_SO", "Quantity_SO", "UOM Name_SO", "Unit Price_SO",
                    "Tax Code_SO", "Whse_SO", "Dist Rule_SO", "Ship To_SO", "Unnamed: 28_SO", "Date_SO", "Customer_SO",
                    "Remarks_SO", "Due Date_SO", "Customer/Vendor Code_SO", "Shipping Type_SO", "#_ARDP", "Series_ARDP",
                    "Customer Code_ARDP", "Customer_ARDP", "Customer Ref No_ARDP", "Customer Name_ARDP", "Currency_ARDP",
                    "Document Date_ARDP", "Item No_ARDP", "Due Date_ARDP", "Item Description_ARDP", "Tax Code_ARDP", "Whse_ARDP",
                    "Dist Rule_ARDP", "Bag Quantity_ARDP", "Ship To_ARDP", "Total_ARDP", "Balance Due_ARDP",
                    "Date_ARDP", "Details_ARDP", "No Faktur Pajak_ARDP", "#_CM", "Series_CM", "Customer Code_CM", "Customer Name_CM",
                    "Due Date_CM", "Item No_CM", "Item Description_CM", "Bag Quantity_CM", "Quantity_CM", "KG_CM", "Netto Quantity_CM", "Netto 1_CM", "KA Type_CM",
                    "Rafaksi %_CM", "UOM Name_CM", "Unit Price_CM", "Discount %_CM", "Tax Code_CM", "Wtax Liable_CM", "Whse_CM", "Dist Rule_CM",
                    "COGS Dist Rule_CM", "Ship To_CM", "Total_CM", "Applied Amount_CM", "Total Before Diskon_CM", "Date_CM", "Customer_CM", "Remarks_CM",
                    "Document Total_CM", "Series_DL", "Status_DL", "Create Date_DL", "Doc Date_DL", "Customer Name_DL", "Item Code_DL", "Item Name_DL", "Qty_DL",
                    "UoM_DL", "Note_DL", "In Time_DL", "Out Time_DL", "Cont No./Seal_DL", "#_DO", "Series_DO", "Customer Code_DO", "Customer Name_DO", "Customer Ref No_DO", "Due Date_DO",
                    "Item No_DO", "Netto 1_DO", "KA Type_DO", "Rafaksi %_DO", "UOM Name_DO", "Wtax Liable_DO", "Whse_DO", "Dist Rule_DO", "Amount Rafaksi %_DO", "Potongan_DO",
                    "Base Price_DO", "Amount Rafaksi_DO", "Ship To_DO", "Driver Name_DO", "Bill Of Lading_DO", "Total Before Diskon_DO", "Date_DO", "Customer_DO", "Remarks_DO", "Document Total_DO", "No. Cont_DO", "Remarks OA_DO", "Vendor Name_DO", "Total Service Price_DO", "Customer/Vendor Ref. No._DO", "DO Draft DocNum_DO", "Document Date_DO", "whs_DO", "doc date 2_DO", "No Invoice OA_DO",
                    "No List Transport_DO", "Total Service After Inv_DO", "SHIP_TO_OVERWRITTEN_DO", "Canceled_DO", "Customer Code_RETURN", "Customer_RETURN", "Item Description_RETURN", "License Number_RETURN", "Remarks_RETURN",
                    "Series_ARIN", "Due Date_ARIN", "Date_ARIN", "Remarks_ARIN", "Document Total_ARIN", "Doc Date_ARIN", "Document Status_ARIN", "Bank Code_ARIN", "Canceled_ARIN", "Document Total (FC)_ARIN",
                    "Base On DO_ARIN", "Base On_SO", "Base On SO_CM", "Base On DO_CM", "Base On ARDP_CM", "Base On SO", "Base On SO_ARDP", "Base on DO_RETURN", "merged_key"]

    df_download = df_combined.copy()

    # --- Tambahan: hapus semua kolom yang mengandung "Unnamed" ---
    unnamed_cols = [col for col in df_download.columns if "Unnamed" in col]

    # Gabungkan dengan daftar kolom yang ingin dihapus
    all_cols_to_drop = list(set(cols_to_drop + unnamed_cols))

    # Hapus hanya kolom yang ada di dataframe
    existing_cols_to_drop = [col for col in all_cols_to_drop if col in df_download.columns]
    if existing_cols_to_drop:
        df_download = df_download.drop(columns=existing_cols_to_drop)
        st.info(f"✅ {len(existing_cols_to_drop)} kolom (termasuk 'Unnamed') berhasil dihapus.")

    # --- DETEKSI TAHUN OTOMATIS ---
    def detect_year_from_data(df):
        date_columns = [
            'Posting Date_DO', 'Posting Date_ARIN', 'Posting Date_SO',
            'Document Date_DO', 'Document Date_ARIN', 'Document Date_SO',
            'Out Date_DL', 'Delivery Date_DO'
        ]
        
        for col in date_columns:
            if col in df.columns:
                first_date = df[col].dropna().head(1)
                if not first_date.empty:
                    try:
                        date_val = pd.to_datetime(first_date.iloc[0], errors='coerce')
                        if pd.notna(date_val):
                            return date_val.year
                    except:
                        continue
        
        return pd.Timestamp.now().year

    detected_year = detect_year_from_data(df_download)
    file_name = f"DATA PENJUALAN {detected_year}.xlsx"

    st.info(f"ℹ️ Tahun terdeteksi dari data: **{detected_year}**")

    # ✅ Gunakan df_download (yang sudah bersih)
    excel_data = convert_df_to_excel(df_download)
    st.download_button(
        label=f"📥 Download Data Penjualan {detected_year} (Excel)",
        data=excel_data,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
