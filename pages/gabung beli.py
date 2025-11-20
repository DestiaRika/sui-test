import streamlit as st
import pandas as pd
import re
import openpyxl
import io
import datetime

st.set_page_config('Gabung Beli', layout='centered')
st.title('Gabung Beli')

vendor = st.file_uploader('Vendor Master Data:', type=['xlsx'], accept_multiple_files=False)
po = st.file_uploader('Purchase Order Data:', type=['xlsx'], accept_multiple_files=False)
timbangan = st.file_uploader('Timbangan Data:', type=['xlsx'], accept_multiple_files=False)
grpo = st.file_uploader('GRPO Data:', type=['xlsx'], accept_multiple_files=False)
apdp = st.file_uploader('APDP Data:', type=['xlsx'], accept_multiple_files=False)
invoice = st.file_uploader('INVOICE Data:', type=['xlsx'], accept_multiple_files=False)

uploaded_files = all(x is not None for x in [vendor, po, timbangan, grpo, apdp, invoice])
st.write(f"Semua file terupload: {uploaded_files}")

if uploaded_files:
    st.success("Files uploaded successfully!")
    if st.button("Process Data"):
        # ========== FILTER CANCELED GRPO DI PALING AWAL ==========
        st.write("=" * 50)
        st.write("PRE-PROCESSING: Membersihkan Data GRPO yang Canceled")
        st.write("=" * 50)
        
        df_grpo_raw = pd.read_excel(grpo)
        df_grpo_raw.columns = df_grpo_raw.columns.str.strip()
        
        if 'Canceled' in df_grpo_raw.columns:
            before_filter = len(df_grpo_raw)
            # Hapus baris yang memiliki nilai 'Yes' atau 'Cancellation' di kolom Canceled
            df_grpo_raw = df_grpo_raw[~df_grpo_raw['Canceled'].astype(str).str.strip().str.upper().isin(['YES', 'CANCELLATION'])]
            after_filter = len(df_grpo_raw)
            removed = before_filter - after_filter
            st.write(f"🗑️ Baris GRPO dengan Canceled = 'Yes' atau 'Cancellation' dihapus: {removed} baris")
            st.write(f"✓ GRPO tersisa: {after_filter} baris")
        else:
            st.warning("⚠️ Kolom 'Canceled' tidak ditemukan di GRPO")
        
        # Simpan GRPO yang sudah dibersihkan ke dalam buffer untuk digunakan di semua tahap
        grpo_buffer = io.BytesIO()
        df_grpo_raw.to_excel(grpo_buffer, index=False, engine='openpyxl')
        grpo_buffer.seek(0)
        
        st.success("✓ Pre-processing GRPO selesai")
        # =========================================================

        st.write("=" * 50)
        st.write("TAHAP 1: Memproses AP INVOICE dengan GRPO")
        st.write("=" * 50)
        
        df_ap = pd.read_excel(invoice)
        df_grpo = pd.read_excel(grpo_buffer)  # ← Gunakan GRPO yang sudah dibersihkan

        df_ap.columns = df_ap.columns.str.strip()
        df_grpo.columns = df_grpo.columns.str.strip()

        # ========== HAPUS DUPLIKAT DI GRPO BERDASARKAN DOC NUMBER ==========
        st.write("🔍 Memeriksa duplikat di GRPO...")
        if 'Doc Number' in df_grpo.columns:
            before_drop_grpo = len(df_grpo)
            df_grpo = df_grpo.drop_duplicates(subset=['Doc Number'], keep='first')
            after_drop_grpo = len(df_grpo)
            removed_grpo = before_drop_grpo - after_drop_grpo
            
            if removed_grpo > 0:
                st.warning(f"⚠️ Ditemukan {removed_grpo} duplikat di GRPO (Doc Number). Duplikat telah dihapus.")
            else:
                st.success("✓ Tidak ada duplikat Doc Number di GRPO")
        
        st.write(f"📊 Total GRPO setelah cleaning: {len(df_grpo)} baris")
        st.write("-" * 50)

        def extract_grpo_numbers(remarks):
            if pd.isna(remarks):
                return []
            matches = re.findall(r"Based On Goods Receipt(?: PO)?\s*([\d\. ]+)", str(remarks))
            if not matches:
                return []
            numbers = re.findall(r"\d{8}", matches[0])
            return numbers

        df_ap["GRPO_List"] = df_ap["Remarks"].apply(extract_grpo_numbers)
        important_cols = ['Doc Number', 'Vendor Code', 'Item No', 'Quantity', 'Unit Price', 'Whse']

        for col in important_cols:
            if col in df_grpo.columns:
                df_grpo[col] = df_grpo[col].astype(str).str.strip()
                df_grpo[col] = df_grpo[col].str.replace(r"\.0+$", "", regex=True).str.upper()
            else:
                st.warning(f"Kolom '{col}' tidak ditemukan di GRPO.")
                df_grpo[col] = None

        used_grpos = set()

        def clean_val(v):
            return re.sub(r"\.0+$", "", str(v).strip().upper())
        
        def find_matching_grpo(row):
            grpo_list = row["GRPO_List"]
            if not grpo_list:
                return None

            best_match = None
            best_score = 0
            ap_values = {col: clean_val(row.get(col, "")) for col in important_cols if col in df_ap.columns}

            for grpo_num in grpo_list:
                if grpo_num in used_grpos:
                    continue

                matches = df_grpo[df_grpo["Doc Number"] == grpo_num]
                if matches.empty:
                    continue

                grpo_row = matches.iloc[0]
                grpo_values = {col: clean_val(grpo_row.get(col, "")) for col in important_cols}
                same_count = sum(ap_values.get(col, "") == grpo_values.get(col, "")
                                for col in important_cols if col != "Doc Number")
                if same_count > best_score:
                    best_score = same_count
                    best_match = grpo_num

            if best_match:
                used_grpos.add(best_match)
                return best_match
            return None

        st.write("Mencocokkan Base On GRPO dari Remarks...")
        df_ap["Base On GRPO"] = df_ap.apply(find_matching_grpo, axis=1)

        if "Remarks" in df_ap.columns:
            df_ap["Remarks"] = df_ap["Remarks"].astype(str).apply(
                lambda x: re.sub(r"Goods Receipt PO\s+(\d{8})", r"Goods Receipt \1", x))

        st.write("Mencari kecocokan tambahan berdasarkan 5 kolom utama...")
        for idx, row in df_ap[df_ap["Base On GRPO"].isna()].iterrows():
            ap_values = {col: clean_val(row.get(col, "")) for col in important_cols if col in df_ap.columns}

            possible_matches = df_grpo.copy()
            for col in ['Vendor Code', 'Item No', 'Quantity', 'Unit Price', 'Whse']:
                if col in possible_matches.columns:
                    possible_matches = possible_matches[
                        possible_matches[col].apply(clean_val) == ap_values.get(col, "")]
                if possible_matches.empty:
                    break

            if not possible_matches.empty:
                match_row = possible_matches[~possible_matches["Doc Number"].isin(used_grpos)]
                if not match_row.empty:
                    docnum = match_row.iloc[0]["Doc Number"]
                    df_ap.at[idx, "Base On GRPO"] = docnum
                    used_grpos.add(docnum)

        df_ap.drop(columns=["GRPO_List"], inplace=True)

        # ========== PENGECEKAN DAN PENANGANAN DUPLIKAT BASE ON GRPO ==========
        st.write("-" * 50)
        st.write("🔍 Memeriksa duplikat di Base On GRPO...")
        
        if "Base On GRPO" in df_ap.columns:
            df_cleaned = df_ap.dropna(subset=['Base On GRPO'])
            duplicate_counts = df_cleaned["Base On GRPO"].value_counts()
            duplicate_values = duplicate_counts[duplicate_counts > 1]

            if not duplicate_values.empty:
                st.warning(f"⚠️ Ditemukan {len(duplicate_values)} nilai duplikat di Base On GRPO")
                
                # Tampilkan duplikat sebelum dihapus
                with st.expander("📋 Lihat detail duplikat yang ditemukan"):
                    for grpo_num in duplicate_values.index:
                        duplicate_rows = df_ap[df_ap["Base On GRPO"] == grpo_num]
                        st.write(f"**GRPO {grpo_num}** muncul **{len(duplicate_rows)} kali**:")
                        display_cols = [col for col in ['Doc Number', 'Vendor Code', 'Item No', 'Quantity', 'Unit Price', 'Base On GRPO'] if col in duplicate_rows.columns]
                        st.dataframe(duplicate_rows[display_cols])
                        st.write("")
                
                # Hapus duplikat - pertahankan yang pertama muncul
                before_drop_apin = len(df_ap)
                df_ap = df_ap.drop_duplicates(subset=['Base On GRPO'], keep='first')
                after_drop_apin = len(df_ap)
                removed_apin = before_drop_apin - after_drop_apin
                
                st.success(f"✓ {removed_apin} baris duplikat berhasil dihapus. Total baris sekarang: {len(df_ap)}")
            else:
                st.success("✓ Tidak ada duplikat Base On GRPO ditemukan")
        
        # ========== MERGE GRPO DENGAN AP INVOICE (RIGHT JOIN) ==========
        st.write("-" * 50)
        st.write("🔗 Menggabungkan GRPO dengan AP Invoice (prioritas APIN)...")
        
        # Standardisasi kolom untuk merge
        df_grpo['Doc Number'] = df_grpo['Doc Number'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        df_ap['Base On GRPO'] = df_ap['Base On GRPO'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        
        # Tambahkan suffix untuk membedakan kolom
        df_grpo_suffixed = df_grpo.add_suffix('_GRPO')
        df_ap_suffixed = df_ap.add_suffix('_APIN')
        
        # Rename key columns kembali tanpa suffix
        df_grpo_suffixed = df_grpo_suffixed.rename(columns={'Doc Number_GRPO': 'merge_key'})
        df_ap_suffixed = df_ap_suffixed.rename(columns={'Base On GRPO_APIN': 'merge_key'})
        
        # Merge dengan RIGHT JOIN - semua APIN dipertahankan
        df_grpo_apin = pd.merge(
            df_grpo_suffixed,
            df_ap_suffixed,
            on='merge_key',
            how='right',  # ← RIGHT JOIN: pertahankan semua APIN
            indicator=True
        )
        
        # Rename merge_key menjadi kolom yang lebih deskriptif
        df_grpo_apin = df_grpo_apin.rename(columns={'merge_key': 'Doc Number_GRPO'})
        
        # Reorder columns: GRPO columns first, then APIN columns
        grpo_cols = [col for col in df_grpo_apin.columns if col.endswith('_GRPO')]
        apin_cols = [col for col in df_grpo_apin.columns if col.endswith('_APIN')]
        other_cols = [col for col in df_grpo_apin.columns if col not in grpo_cols + apin_cols]
        
        df_grpo_apin = df_grpo_apin[grpo_cols + apin_cols + other_cols]
        
        # Statistik merge
#region timbangan
#region timbangan
        st.write("=" * 50)
        st.write("TAHAP 2: Matching pertama (berdasarkan Weight Difference dan tanggal)")
        st.write("=" * 50)
        
        df_timbangan = pd.read_excel(timbangan)
        df_grpo = df_grpo_apin.copy()  # ← Gunakan hasil Tahap 1, bukan load ulang

        # Bersihkan nama kolom dari spasi
        df_timbangan.columns = df_timbangan.columns.str.strip()
        df_grpo.columns = df_grpo.columns.str.strip()

        st.write("📂 File berhasil dibaca!")
        st.write(f"- TIMBANGAN: {len(df_timbangan)} baris")
        st.write(f"- GRPO (dari Tahap 1): {len(df_grpo)} baris")
        st.write("=" * 70)

        # ===================================================================
        # BAGIAN 2: DEFINISI FUNGSI HELPER (REVISI)
        # ===================================================================

        # FUNGSI HELPER UNTUK FORMAT ANGKA (dipakai di semua tahap)
        def format_number(x):
            """Format angka: hapus .0 jika integer, kosongkan jika NaN"""
            if pd.isna(x) or str(x).strip() == "":
                return ""
            try:
                fx = float(x)
                # Tambahkan pengecekan untuk menghindari representasi float yang tidak perlu
                if abs(fx - round(fx)) < 1e-9:
                    return str(int(round(fx)))
                return str(fx)
            except:
                return str(x).strip()

        # FUNGSI HELPER UNTUK FORMAT TANGGAL (REVISI)
        def format_date(series, format="%d-%m-%Y"):
            """Standardisasi format tanggal, kembalikan string kosong jika NaT"""
            date_series = pd.to_datetime(series, errors='coerce')
            
            def date_formatter(x):
                if pd.isna(x) or pd.isnull(x):
                    return ""
                try:
                    return x.strftime(format)
                except:
                    return ""
                    
            return date_series.apply(date_formatter)

        # ===================================================================
        # BAGIAN 3: PROSES MATCHING
        # ===================================================================

        # Kolom yang digunakan (SESUAIKAN dengan suffix _GRPO)
        cols_timbangan = ["Vendor Code", "Product Code", "Weight Difference", "Out Date"]
        cols_grpo = ["Vendor Code_GRPO", "Item No_GRPO", "Quantity_GRPO", "Posting Date_GRPO"]

        # ===============================================
        # 🔹 1. Format tanggal jadi datetime dan samakan format
        # ===============================================
        if "Out Date" in df_timbangan.columns:
            df_timbangan["Out Date"] = pd.to_datetime(df_timbangan["Out Date"], errors='coerce', dayfirst=True)
        
        if "Posting Date_GRPO" in df_grpo.columns:
            df_grpo["Posting Date_GRPO"] = pd.to_datetime(df_grpo["Posting Date_GRPO"], errors='coerce', dayfirst=True)

        # =======================================================
        # 🔹 2. Buat salinan untuk merged_key (string format seragam dd.MM.yyyy)
        # =======================================================
        df_timbangan_temp = df_timbangan.copy()
        df_grpo_temp = df_grpo.copy()

        if "Out Date" in df_timbangan_temp.columns:
            df_timbangan_temp["Out Date"] = df_timbangan_temp["Out Date"].dt.strftime("%d.%m.%Y")

        if "Posting Date_GRPO" in df_grpo_temp.columns:
            df_grpo_temp["Posting Date_GRPO"] = df_grpo_temp["Posting Date_GRPO"].dt.strftime("%d.%m.%Y")

        # =======================================================
        # 🔹 3. Bersihkan ".0" dari kolom numerik/string agar match sempurna
        # =======================================================
        # Timbangan
        for col in cols_timbangan:
            if col in df_timbangan_temp.columns:
                df_timbangan_temp[col] = df_timbangan_temp[col].astype(str).str.replace(r"\.0$", "", regex=True)

        # GRPO
        for col in cols_grpo:
            if col in df_grpo_temp.columns:
                df_grpo_temp[col] = df_grpo_temp[col].astype(str).str.replace(r"\.0$", "", regex=True)

        # =======================================================
        # 🔹 4. Buat merged_key dengan format tanggal seragam dd.MM.yyyy
        # =======================================================
        df_timbangan["merged_key"] = df_timbangan_temp[cols_timbangan].agg("_".join, axis=1)
        df_grpo["merged_key"] = df_grpo_temp[cols_grpo].agg("_".join, axis=1)

 # ========================================
        # 🔹 5. Hapus duplikat di timbangan BERDASARKAN MERGED_KEY
        # ========================================
        st.write("🔍 Memeriksa duplikat di Timbangan (merged_key)...")
        before_drop_timb = len(df_timbangan)
        df_timbangan = df_timbangan.drop_duplicates(subset=["merged_key"], keep="first")
        after_drop_timb = len(df_timbangan)
        removed_timb = before_drop_timb - after_drop_timb
        
        if removed_timb > 0:
            st.warning(f"⚠️ Ditemukan {removed_timb} duplikat di Timbangan (merged_key). Duplikat telah dihapus.")
        else:
            st.success("✓ Tidak ada duplikat merged_key di Timbangan")

        # ========================================
        # 🔹 6. Rename kolom agar aman saat merge
        # ========================================
        df_timbangan = df_timbangan.add_suffix("_TIMB")
        # GRPO sudah punya suffix _GRPO dan _APIN dari Tahap 1, jadi tidak perlu add_suffix lagi
        
        df_timbangan = df_timbangan.rename(columns={"merged_key_TIMB": "merged_key"})

        # ========================================
        # 🔹 7. Lakukan merge (matching tahap pertama)
        # ========================================
        df_grpo_timbangan_pertama = pd.merge(df_grpo, df_timbangan, on="merged_key", how="left")

        # ========== HAPUS DUPLIKAT _GRPO SETELAH MERGE ==========
        st.write("🔍 Memeriksa duplikat Doc Number_GRPO setelah merge...")
        if 'Doc Number_GRPO' in df_grpo_timbangan_pertama.columns:
            before_drop = len(df_grpo_timbangan_pertama)
            # Drop duplikat berdasarkan Doc Number_GRPO, keep first
            df_grpo_timbangan_pertama = df_grpo_timbangan_pertama.drop_duplicates(subset=['Doc Number_GRPO'], keep='first')
            after_drop = len(df_grpo_timbangan_pertama)
            removed = before_drop - after_drop
            
            if removed > 0:
                st.warning(f"⚠️ Ditemukan {removed} duplikat Doc Number_GRPO setelah merge. Duplikat telah dihapus (keep first).")
            else:
                st.success("✓ Tidak ada duplikat Doc Number_GRPO setelah merge")

        # ========================================
        # 🔹 8. Pisahkan hasil MATCH dan TIDAK MATCH
        # ========================================
        if "Doc Number_TIMB" in df_grpo_timbangan_pertama.columns:
            df_grpo_timb_ada = df_grpo_timbangan_pertama[
                df_grpo_timbangan_pertama["Doc Number_TIMB"].notna() &
                (df_grpo_timbangan_pertama["Doc Number_TIMB"].astype(str).str.strip() != "")
            ].copy()

            df_grpo_timb_tidak_ada = df_grpo_timbangan_pertama[
                df_grpo_timbangan_pertama["Doc Number_TIMB"].isna() |
                (df_grpo_timbangan_pertama["Doc Number_TIMB"].astype(str).str.strip() == "")
            ].copy()

            # Hapus kolom _TIMB dari yang belum match
            kolom_timb = [col for col in df_grpo_timb_tidak_ada.columns if col.endswith("_TIMB")]
            df_grpo_timb_tidak_ada = df_grpo_timb_tidak_ada.drop(columns=kolom_timb)

            # ========================================
            # 🔹 9. Tampilkan hasil
            # ========================================
            st.write(f"✅ Jumlah baris MATCH pertama (Weight Difference): **{len(df_grpo_timb_ada)}**")
            st.write(f"❌ Jumlah baris belum match: **{len(df_grpo_timb_tidak_ada)}**")

            if not df_grpo_timb_ada.empty:
                st.write("🔑 Daftar merged_key yang berhasil match (10 teratas):")
                st.dataframe(df_grpo_timb_ada[["merged_key"]].head(10))
            else:
                st.info("Tidak ada data yang match pada tahap ini.")

        else:
            st.warning("Kolom 'Doc Number_TIMB' belum ditemukan di hasil merge pertama.")

        # ========================================
        # 🔹 10. Preview hasil akhir
        # ========================================
        st.success("✓ Matching pertama selesai")
        st.write("Preview hasil TAHAP 2 (10 data teratas):")
        st.dataframe(df_grpo_timbangan_pertama.head(10))
        st.write("=" * 50)


        # ========================================================================
        st.write("\nTAHAP 3: Matching Kedua (License Number + Aggregated Quantity)")
        st.write("=" * 70)

        # Reload timbangan untuk matching kedua
        df_timbangan_kedua = pd.read_excel(timbangan)
        df_timbangan_kedua.columns = df_timbangan_kedua.columns.str.strip()
        df_grpo_timb_tidak_ada.columns = df_grpo_timb_tidak_ada.columns.str.strip()

        # ========== STANDARDISASI TIMBANGAN KEDUA ==========
        # Uppercase untuk text columns
        for col in ["Vendor Code", "Product Code", "License No."]:
            if col in df_timbangan_kedua.columns:
                df_timbangan_kedua[col] = df_timbangan_kedua[col].astype(str).str.strip().str.upper()

        # Format tanggal KONSISTEN (sama dengan tahap 2)
        if "Out Date" in df_timbangan_kedua.columns:
            df_timbangan_kedua["Out Date"] = format_date(df_timbangan_kedua["Out Date"])

        # Format Weight Difference
        if "Weight Difference" in df_timbangan_kedua.columns:
            df_timbangan_kedua["Weight Difference"] = df_timbangan_kedua["Weight Difference"].apply(format_number)

        # ========== STANDARDISASI GRPO (yang belum match) ==========
        # Uppercase untuk text columns
        for col in ["Vendor Code_GRPO", "Item No_GRPO", "License Number_GRPO"]:
            if col in df_grpo_timb_tidak_ada.columns:
                df_grpo_timb_tidak_ada[col] = df_grpo_timb_tidak_ada[col].astype(str).str.strip().str.upper()

        # Format Posting Date KONSISTEN (sama dengan tahap 2)
        if "Posting Date_GRPO" in df_grpo_timb_tidak_ada.columns:
            df_grpo_timb_tidak_ada["Posting Date_GRPO"] = format_date(df_grpo_timb_tidak_ada["Posting Date_GRPO"])

        # Konversi Quantity ke numeric
        if "Quantity_GRPO" in df_grpo_timb_tidak_ada.columns:
            df_grpo_timb_tidak_ada["Quantity_GRPO"] = pd.to_numeric(
                df_grpo_timb_tidak_ada["Quantity_GRPO"], errors="coerce"
            ).fillna(0)

        # ========== AGGREGASI QUANTITY ==========
        key_cols = ["Vendor Code_GRPO", "Item No_GRPO", "License Number_GRPO", "Posting Date_GRPO"]

        # Sum Quantity per group
        df_grpo_timb_tidak_ada["_qty_sum"] = (
            df_grpo_timb_tidak_ada
            .groupby(key_cols, dropna=False)["Quantity_GRPO"]
            .transform("sum")
        )

        # Format quantity sum
        df_grpo_timb_tidak_ada["_qty_sum_fmt"] = df_grpo_timb_tidak_ada["_qty_sum"].apply(format_number)

        # ========== BUAT MERGED KEY KEDUA ==========
        # Fungsi untuk membersihkan .0 dari string
        def clean_decimal_zero(text):
            """Hapus .0 dari akhir string jika ada"""
            text = str(text).strip()
            if text.endswith('.0'):
                return text[:-2]
            return text

        # GRPO: Vendor_Item_License_SumQty_Date (bersihkan setiap komponen)
        df_grpo_timb_tidak_ada["merged_key"] = (
            df_grpo_timb_tidak_ada["Vendor Code_GRPO"].apply(clean_decimal_zero) + "_" +
            df_grpo_timb_tidak_ada["Item No_GRPO"].apply(clean_decimal_zero) + "_" +
            df_grpo_timb_tidak_ada["License Number_GRPO"].apply(clean_decimal_zero) + "_" +
            df_grpo_timb_tidak_ada["_qty_sum_fmt"].apply(clean_decimal_zero) + "_" +
            df_grpo_timb_tidak_ada["Posting Date_GRPO"].apply(clean_decimal_zero)
        )

        # TIMBANGAN: Vendor_Product_License_WeightDiff_Date (bersihkan setiap komponen)
        df_timbangan_kedua["merged_key"] = (
            df_timbangan_kedua["Vendor Code"].apply(clean_decimal_zero) + "_" +
            df_timbangan_kedua["Product Code"].apply(clean_decimal_zero) + "_" +
            df_timbangan_kedua["License No."].apply(clean_decimal_zero) + "_" +
            df_timbangan_kedua["Weight Difference"].apply(clean_decimal_zero) + "_" +
            df_timbangan_kedua["Out Date"].apply(clean_decimal_zero)
        )

        st.write("🔍 Preview merged_key TIMBANGAN kedua (5 baris pertama):")
        st.dataframe(df_timbangan_kedua[["Vendor Code", "Product Code", "License No.", "Weight Difference", "Out Date", "merged_key"]].head(5))

        st.write("\n🔍 Preview merged_key GRPO kedua (5 baris pertama):")
        st.dataframe(df_grpo_timb_tidak_ada[["Vendor Code_GRPO", "Item No_GRPO", "License Number_GRPO", "_qty_sum_fmt", "Posting Date_GRPO", "merged_key"]].head(5))

        # ========== MERGE KEDUA ==========
        # Clean merged_key
        df_timbangan_kedua['merged_key'] = df_timbangan_kedua['merged_key'].astype(str).str.strip()
        df_grpo_timb_tidak_ada['merged_key'] = df_grpo_timb_tidak_ada['merged_key'].astype(str).str.strip()

        # Rename kolom timbangan
        cols_to_rename = {col: f"{col}_TIMB" for col in df_timbangan_kedua.columns if col != 'merged_key'}
        df_timbangan_kedua = df_timbangan_kedua.rename(columns=cols_to_rename)

        # Merge
# Merge
        df_grpo_timbangan_kedua = pd.merge(df_grpo_timb_tidak_ada, df_timbangan_kedua, on='merged_key', how='left')

        # ========== HAPUS DUPLIKAT _GRPO SETELAH MERGE TAHAP 3 ==========
        st.write("🔍 Memeriksa duplikat Doc Number_GRPO setelah merge Tahap 3...")
        if 'Doc Number_GRPO' in df_grpo_timbangan_kedua.columns:
            before_drop = len(df_grpo_timbangan_kedua)
            df_grpo_timbangan_kedua = df_grpo_timbangan_kedua.drop_duplicates(subset=['Doc Number_GRPO'], keep='first')
            after_drop = len(df_grpo_timbangan_kedua)
            removed = before_drop - after_drop
            
            if removed > 0:
                st.warning(f"⚠️ Ditemukan {removed} duplikat Doc Number_GRPO di Tahap 3. Duplikat telah dihapus.")
            else:
                st.success("✓ Tidak ada duplikat Doc Number_GRPO di Tahap 3")

        # Hitung yang match di tahap kedua
        matched_second = df_grpo_timbangan_kedua['Doc Number_TIMB'].notna().sum()

        # Hitung yang match di tahap kedua
        matched_second = df_grpo_timbangan_kedua['Doc Number_TIMB'].notna().sum()
        st.write(f"\n✅ **Matching Kedua Berhasil**: {matched_second} baris")

        st.success("✓ TAHAP 3 selesai")
        st.write("=" * 70)


        # ========================================================================
        # TAHAP 4: MATCHING KETIGA (TANPA DATE)
        # ========================================================================
        st.write("\nTAHAP 4: Matching Ketiga (Tanpa Date)")
        st.write("=" * 70)

        # Filter yang belum match di tahap 2
        df_grpo_timb_tidak_ada_kedua = df_grpo_timbangan_kedua[df_grpo_timbangan_kedua['Doc Number_TIMB'].isna()].copy()
        
        if len(df_grpo_timb_tidak_ada_kedua) > 0:
            st.write(f"🔍 Data GRPO yang belum match di tahap 2: {len(df_grpo_timb_tidak_ada_kedua)} baris")
            
            # Reload timbangan untuk matching ketiga
            df_timbangan_ketiga = pd.read_excel(timbangan)
            df_timbangan_ketiga.columns = df_timbangan_ketiga.columns.str.strip()
            
            # ========== STANDARDISASI TIMBANGAN KETIGA ==========
            for col in ["Vendor Code", "Product Code", "License No."]:
                if col in df_timbangan_ketiga.columns:
                    df_timbangan_ketiga[col] = df_timbangan_ketiga[col].astype(str).str.strip().str.upper()
            
            if "Weight Difference" in df_timbangan_ketiga.columns:
                df_timbangan_ketiga["Weight Difference"] = df_timbangan_ketiga["Weight Difference"].apply(format_number)
            
            # ========== BUAT MERGED KEY KETIGA (TANPA DATE) ==========
            # GRPO: Vendor_Item_License_SumQty (TANPA Date)
            df_grpo_timb_tidak_ada_kedua["merged_key_nodate"] = (
                df_grpo_timb_tidak_ada_kedua["Vendor Code_GRPO"].apply(clean_decimal_zero) + "_" +
                df_grpo_timb_tidak_ada_kedua["Item No_GRPO"].apply(clean_decimal_zero) + "_" +
                df_grpo_timb_tidak_ada_kedua["License Number_GRPO"].apply(clean_decimal_zero) + "_" +
                df_grpo_timb_tidak_ada_kedua["_qty_sum_fmt"].apply(clean_decimal_zero)
            )
            
            # TIMBANGAN: Vendor_Product_License_WeightDiff (TANPA Date)
            df_timbangan_ketiga["merged_key_nodate"] = (
                df_timbangan_ketiga["Vendor Code"].apply(clean_decimal_zero) + "_" +
                df_timbangan_ketiga["Product Code"].apply(clean_decimal_zero) + "_" +
                df_timbangan_ketiga["License No."].apply(clean_decimal_zero) + "_" +
                df_timbangan_ketiga["Weight Difference"].apply(clean_decimal_zero)
            )
            
            st.write("🔍 Preview merged_key_nodate TIMBANGAN ketiga (5 baris pertama):")
            st.dataframe(df_timbangan_ketiga[["Vendor Code", "Product Code", "License No.", "Weight Difference", "merged_key_nodate"]].head(5))
            
            st.write("\n🔍 Preview merged_key_nodate GRPO ketiga (5 baris pertama):")
            st.dataframe(df_grpo_timb_tidak_ada_kedua[["Vendor Code_GRPO", "Item No_GRPO", "License Number_GRPO", "_qty_sum_fmt", "merged_key_nodate"]].head(5))
            
            # ========== MERGE KETIGA ==========
            # Clean merged_key_nodate
            df_timbangan_ketiga['merged_key_nodate'] = df_timbangan_ketiga['merged_key_nodate'].astype(str).str.strip()
            df_grpo_timb_tidak_ada_kedua['merged_key_nodate'] = df_grpo_timb_tidak_ada_kedua['merged_key_nodate'].astype(str).str.strip()
            
            # Rename kolom timbangan
            cols_to_rename_ketiga = {col: f"{col}_TIMB" for col in df_timbangan_ketiga.columns if col != 'merged_key_nodate'}
            df_timbangan_ketiga = df_timbangan_ketiga.rename(columns=cols_to_rename_ketiga)
            
            # Drop kolom _TIMB dari tahap sebelumnya untuk hindari duplikasi
            cols_to_drop = [col for col in df_grpo_timb_tidak_ada_kedua.columns if col.endswith('_TIMB')]
            df_grpo_timb_tidak_ada_kedua = df_grpo_timb_tidak_ada_kedua.drop(columns=cols_to_drop)
            
# Merge
            df_grpo_timbangan_ketiga = pd.merge(
                df_grpo_timb_tidak_ada_kedua,
                df_timbangan_ketiga, 
                on='merged_key_nodate', 
                how='left'
            )
            
            # ========== HAPUS DUPLIKAT _GRPO SETELAH MERGE TAHAP 4 ==========
            st.write("🔍 Memeriksa duplikat Doc Number_GRPO setelah merge Tahap 4...")
            if 'Doc Number_GRPO' in df_grpo_timbangan_ketiga.columns:
                before_drop = len(df_grpo_timbangan_ketiga)
                df_grpo_timbangan_ketiga = df_grpo_timbangan_ketiga.drop_duplicates(subset=['Doc Number_GRPO'], keep='first')
                after_drop = len(df_grpo_timbangan_ketiga)
                removed = before_drop - after_drop
                
                if removed > 0:
                    st.warning(f"⚠️ Ditemukan {removed} duplikat Doc Number_GRPO di Tahap 4. Duplikat telah dihapus.")
                else:
                    st.success("✓ Tidak ada duplikat Doc Number_GRPO di Tahap 4")
            
            # Hitung yang match di tahap ketiga
            matched_third = df_grpo_timbangan_ketiga['Doc Number_TIMB'].notna().sum()
            
            # Hitung yang match di tahap ketiga
            matched_third = df_grpo_timbangan_ketiga['Doc Number_TIMB'].notna().sum()
            st.write(f"\n✅ **Matching Ketiga Berhasil**: {matched_third} baris")
            
            # ========== GABUNGKAN HASIL AKHIR ==========
            # Ambil yang match di tahap 2
            df_matched_tahap2 = df_grpo_timbangan_kedua[df_grpo_timbangan_kedua['Doc Number_TIMB'].notna()].copy()
            
            # Gabungkan: Tahap 1 + Tahap 2 (matched) + Tahap 3 (all - karena sudah include yang matched dan tidak)
            df_grpo_timbangan_final = pd.concat([
                df_grpo_timb_ada,           # Tahap 2 (match pertama)
                df_matched_tahap2,          # Tahap 3 yang matched
                df_grpo_timbangan_ketiga    # Tahap 4 (semua yang dicoba match)
            ], ignore_index=True)
            
            total_matched = len(df_grpo_timb_ada) + matched_second + matched_third
            
            st.write(f"\n📊 **Total Baris Final**: {len(df_grpo_timbangan_final)}")
            st.write(f"📊 **Total Match Keseluruhan**:")
            st.write(f"   - Tahap 2 (Vendor+Item+Qty+Date): {len(df_grpo_timb_ada)} baris")
            st.write(f"   - Tahap 3 (License+Qty+Date): {matched_second} baris")
            st.write(f"   - Tahap 4 (License+Qty, tanpa Date): {matched_third} baris")
            st.write(f"   - **TOTAL MATCHED**: {total_matched} baris")
            
            st.success("✓ TAHAP 4 selesai")
            st.write("=" * 70)
        else:
            st.info("✓ Semua data sudah match di Tahap 2 dan 3. Tidak ada yang perlu diproses di Tahap 4.")
            df_grpo_timbangan_final = pd.concat([df_grpo_timb_ada, df_grpo_timbangan_kedua], ignore_index=True)
            
            st.write(f"\n📊 **Total Baris Final**: {len(df_grpo_timbangan_final)}")
            st.success("✓ Data final disimpan di df_grpo_timbangan_final")
            st.write("=" * 70)
            
        st.write("TAHAP 5-6: Memproses APDP dan Menggabungkan dengan PO")
        st.write("=" * 50)

        # ===== TAHAP 5: Proses APDP =====
        df_combined = pd.read_excel(apdp)

        def extract_do_number(Details):
            if Details:
                match = re.search(r'Purchase Orders (\d{8})', str(Details))
                if match:
                    return match.group(1)
            return None

        df_combined['Base on PO'] = df_combined['Details'].apply(extract_do_number)

        # Cek duplikat
        duplicates = df_combined[df_combined.duplicated(subset=['Base on PO'], keep=False)]
        if len(duplicates) > 0:
            st.warning(f"⚠️ Ditemukan {duplicates['Base on PO'].nunique()} PO yang memiliki baris duplikat")

        # Ambil baris TERAKHIR untuk setiap PO (bukan dijumlahkan)
        df_apdp_232 = df_combined.drop_duplicates(subset=['Base on PO'], keep='last')

        st.success("✓ APDP diproses - duplikat dihapus (mengambil baris terakhir)")
        st.write(f"   Baris awal: {len(df_combined)} → Baris akhir: {len(df_apdp_232)}")

        # ===== TAHAP 6: Gabung dengan PO =====
        df_po = pd.read_excel(po)
        df_apdp = df_apdp_232

        po_key = 'Doc Number'
        apdp_key = 'Base on PO'

        # Tambah suffix
        df_po = df_po.add_suffix('_PO')
        df_apdp = df_apdp.add_suffix('_APDP')

        # Standarisasi format (strip whitespace)
        df_po[f"{po_key}_PO"] = df_po[f"{po_key}_PO"].astype(str).str.strip()
        df_apdp[f"{apdp_key}_APDP"] = df_apdp[f"{apdp_key}_APDP"].astype(str).str.strip()

        # LEFT MERGE: Prioritas PO, hapus APDP yang tidak ada di PO
        df_merged = pd.merge(df_po, df_apdp, how='left', left_on=f"{po_key}_PO", right_on=f"{apdp_key}_APDP")

        # Info hasil merge
        po_dengan_apdp = df_merged[df_merged[f"{apdp_key}_APDP"].notna()].shape[0]
        po_tanpa_apdp = df_merged[df_merged[f"{apdp_key}_APDP"].isna()].shape[0]

        st.info(f"📊 PO dengan APDP: {po_dengan_apdp} | PO tanpa APDP: {po_tanpa_apdp}")

        # Urutkan kolom: PO dulu, baru APDP
        po_cols = [col for col in df_merged.columns if col.endswith('_PO')]
        apdp_cols = [col for col in df_merged.columns if col.endswith('_APDP')]
        df_merged = df_merged[po_cols + apdp_cols]

        df_po_apdp_2023 = df_merged

        st.success("✓ PO dan APDP digabungkan (LEFT JOIN - prioritas PO)")
        st.write(f"   Total baris hasil merge: {len(df_po_apdp_2023)}")

        st.write("TAHAP 7: Menggabungkan PO-APDP dengan TIMBANGAN-GRPO-APIN")
        st.write("=" * 50)

        df_combined = df_grpo_timbangan_final.copy()

        # Cek dulu nama kolom yang ada
        st.write("🔍 Kolom yang tersedia:")
        st.write([col for col in df_combined.columns if 'Remarks' in col])

        def extract_po_number(Details):
            if Details:
                match = re.search(r'Purchase Orders (\d{8})', str(Details))
                if match:
                    return match.group(1)
            return None

        # ✅ PERBAIKAN: Gunakan nama kolom yang BENAR sesuai hasil Tahap 4
        # Cek apakah kolom Remarks_GRPO atau Remarks_APIN yang ada
        if 'Remarks_GRPO' in df_combined.columns:
            df_combined['Base on PO GRPO'] = df_combined['Remarks_GRPO'].apply(extract_po_number)
        elif 'Remarks_APIN' in df_combined.columns:
            df_combined['Base on PO GRPO'] = df_combined['Remarks_APIN'].apply(extract_po_number)
        else:
            st.error("❌ Kolom Remarks tidak ditemukan!")
            st.stop()

        df_timbangan_grpo_apin_key = df_combined
        df_timbangan_mapin = df_timbangan_grpo_apin_key
        df_po_apdp = df_po_apdp_2023.copy()

        # Standardisasi format
        df_timbangan_mapin['Base on PO GRPO'] = (
            df_timbangan_mapin['Base on PO GRPO']
            .astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        )

        df_po_apdp['Doc Number_PO'] = (
            df_po_apdp['Doc Number_PO']
            .astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        )

        # LEFT MERGE
        df_merged = pd.merge(
            df_timbangan_mapin, 
            df_po_apdp, 
            how='left', 
            left_on='Base on PO GRPO', 
            right_on='Doc Number_PO'
        )

        # Info hasil merge
        po_match = df_merged['Doc Number_PO'].notna().sum()
        po_no_match = df_merged['Doc Number_PO'].isna().sum()
        st.info(f"📊 Match dengan PO: {po_match} | Tidak match: {po_no_match}")

        # Urutkan kolom
        order_suffix = ['_PO', '_APDP', '_TIMB', '_GRPO', '_APIN']
        ordered_columns = []
        for suffix in order_suffix:
            cols = [col for col in df_merged.columns if col.endswith(suffix)]
            ordered_columns.extend(cols)

        other_columns = [col for col in df_merged.columns if col not in ordered_columns]
        final_columns = ordered_columns + other_columns
        df_merged = df_merged[final_columns]

        df_po_apin = df_merged
        st.success("✓ PO-APDP digabung dengan TIMBANGAN-GRPO-APIN")
        st.write(f"   Total baris: {len(df_po_apin)}")
        st.write("=" * 50)

        # ========== TAHAP 8: SUPPLIER ==========
        st.write("TAHAP 8: Menggabungkan dengan SUPPLIER")
        st.write("=" * 50)

        df_supplier = pd.read_excel(vendor)
        df_poin = df_po_apin.copy()

        # Standardisasi Vendor Code
        df_supplier['Vendor Code'] = (
            df_supplier['Vendor Code']
            .astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        )

        # ✅ PERBAIKAN: Cek kolom Vendor Code yang tersedia
        vendor_col = None
        if 'Vendor Code_GRPO' in df_poin.columns:
            vendor_col = 'Vendor Code_GRPO'
        elif 'Vendor Code_TIMB' in df_poin.columns:
            vendor_col = 'Vendor Code_TIMB'
        elif 'Vendor Code_APIN' in df_poin.columns:
            vendor_col = 'Vendor Code_APIN'
        else:
            st.error("❌ Kolom Vendor Code tidak ditemukan!")
            st.stop()

        df_poin[vendor_col] = (
            df_poin[vendor_col]
            .astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        )

        # Tambah suffix
        df_supplier = df_supplier.add_suffix('_SUPPLIER')

        # LEFT MERGE
        df_merged = pd.merge(
            df_poin,
            df_supplier,
            how='left',
            left_on=vendor_col,
            right_on='Vendor Code_SUPPLIER'
        )

        # Info hasil merge
        supplier_match = df_merged['Vendor Code_SUPPLIER'].notna().sum()
        supplier_no_match = df_merged['Vendor Code_SUPPLIER'].isna().sum()
        st.info(f"📊 Match dengan Supplier: {supplier_match} | Tidak match: {supplier_no_match}")

        # Urutkan kolom
        order_suffix = ['_SUPPLIER', '_PO', '_APDP', '_TIMB', '_GRPO', '_APIN']
        ordered_columns = []
        for suffix in order_suffix:
            cols = [col for col in df_merged.columns if col.endswith(suffix)]
            ordered_columns.extend(cols)

        other_columns = [col for col in df_merged.columns if col not in ordered_columns]
        final_columns = ordered_columns + other_columns
        df_merged = df_merged[final_columns]

        df_supplier_apin = df_merged
        st.success("✓ Supplier berhasil digabungkan")
        st.write(f"   Total baris: {len(df_supplier_apin)}")
        st.write("=" * 50)

        # ========== TAHAP 9: CLEANING & EXPORT ==========
        st.write("TAHAP 9: Membersihkan Kolom dan Export")
        st.write("=" * 50)

        df = df_supplier_apin.copy()

        # Daftar kolom yang akan dihapus
        cols_to_drop = [
            'Vendor Code_PO', 'Vendor Name_PO', 'UOM Name_PO', 'Unit Price_PO', '#_GRPO',
            'Tax Code_PO', 'Wtax Liable_PO', 'Whse_PO', 'Ship To_PO', 'Unnamed: 27_PO',
            'Vendor_PO', 'Remarks_PO', 'Customer/Vendor Code_PO', 'Series_APDP', 'Freight Document_GRPO', '#_APIN',
            'Vendor Code_APDP', 'Vendor_APDP', 'Vendor Ref No_APDP', 'Vendor Name_APDP',
            'Currency_APDP', 'Item No_APDP', 'Item Description_APDP', 'UOM Name_APDP',
            'Unit Price_APDP', 'Tax Code_APDP', 'Wtax Liable_APDP', 'Whse_APDP',
            'Ship To_APDP', 'Unnamed: 30_APDP', 'Details_APDP', 'Customer/Vendor Code_APDP',
            'Base on PO_APDP', 'Series_timbangan', 'Product Code_timbangan',
            'Unnamed: 6_timbangan', 'Customer Name_timbangan', 'Item Name_timbangan',
            'Vendor Code_timbangan', 'merged_key_timbangan', 'Series_GRPO',
            'Vendor Code_GRPO', 'Vendor Name_GRPO', 'Vendor Ref No_GRPO',
            'BP Currency_GRPO', 'Item No_GRPO', 'KA Type_GRPO',
            'Rafaksi %_GRPO', 'UOM Name_GRPO', 'Unit Price_GRPO', 'Discount %_GRPO',
            'Tax Code_GRPO', 'Wtax Liable_GRPO', 'Whse_GRPO', 'Ship To_GRPO',
            'Shipping Type_GRPO', 'Vendor Transport Code_GRPO', 'Vendor Transport Name_GRPO',
            'Driver Name_GRPO', 'Bill Of Lading_GRPO', 'License Number_GRPO',
            'Unnamed: 58_GRPO', 'Vendor_GRPO', 'Remarks_GRPO', 'Series_APIN',
            'Vendor Code_APIN', 'Vendor_APIN', 'Vendor Ref No_APIN', 'Vendor Name_APIN',
            'Currency_APIN', 'Series_PO', "Vendor Foreign Name_SUPPLIER", 
            "VendorBalance_SUPPLIER", "Payment Terms Code_SUPPLIER",
            "Active_SUPPLIER", "Telephone 1_SUPPLIER", "Telephone 2_SUPPLIER", 
            "Mobile Phone_SUPPLIER", "Creation Date_SUPPLIER",
            "Unnamed: 28_SUPPLIER", "Delivery Date_PO", "Document Date_PO", 
            "Vendor Reff No_PO", "BP Currency_PO", "Item No_PO",
            "Item Description_PO", "Quantity_PO", "Document Date_APDP", "Due Date_APDP", 
            "Quantity_APDP", "Price After Discount_APDP",
            "Discount %_APDP", "Canceled_APDP", "Document Total_APDP", 
            "Document Total (FC)_APDP", "Series_TIMB", "Product Code_TIMB",
            "Warehouse_TIMB", "Canceled_TIMB", "Unnamed: 6_TIMB", "In Time_TIMB", 
            "Out Time_TIMB", "Customer Name_TIMB", "Item Name_TIMB",
            "Status_TIMB", "Vendor Code_TIMB", "Due Date_GRPO", "Bag Quantity_GRPO", 
            "Quantity_GRPO", "KG_GRPO", "Netto Quantity_GRPO",
            "Dist Rule_GRPO", "Container Number_GRPO", "Date_GRPO", "Document Date_GRPO", 
            "Total Service Price_GRPO", "Document Total_GRPO", "Canceled_GRPO", 
            "Due Date_APIN", "Dist Rule_APIN", "Unnamed: 51_APIN", "Date_APIN", 
            "Remarks_APIN", "Paid to Date_APIN",
            "Canceled_APIN", "merged_key", "_qty_sum", "_qty_sum_fmt", 
            "merged_key_nodate", "Base on PO GRPO"
        ]

        # ========== NORMALISASI SEMUA KOLOM POSTING DATE SEBELUM DOWNLOAD ==========

        # Cari semua kolom yang mengandung "Posting Date"
        posting_date_cols = [col for col in df.columns if 'Posting Date' in col]

        if posting_date_cols:
            st.write(f"🔍 Ditemukan {len(posting_date_cols)} kolom Posting Date:")
            for col in posting_date_cols:
                st.write(f"   - {col}")
            
            # Fungsi untuk parse mixed date format
            def parse_mixed_date(date_val):
                """Parse tanggal dengan berbagai format menjadi format standar YYYY-MM-DD HH:MM:SS"""
                if pd.isna(date_val):
                    return pd.NaT
                
                if isinstance(date_val, pd.Timestamp):
                    return date_val
                
                date_str = str(date_val).strip()
                
                # Coba parsing dengan dayfirst=False dulu (MM-DD-YYYY atau YYYY-MM-DD)
                try:
                    parsed = pd.to_datetime(date_str, errors='coerce', dayfirst=False)
                    if pd.notna(parsed):
                        return parsed
                except:
                    pass
                
                # Jika gagal, coba dengan dayfirst=True (DD-MM-YYYY)
                try:
                    parsed = pd.to_datetime(date_str, errors='coerce', dayfirst=True)
                    if pd.notna(parsed):
                        return parsed
                except:
                    pass
                
                return pd.NaT
            
            # Normalisasi setiap kolom Posting Date
            for col in posting_date_cols:
                st.write(f"⏳ Memproses kolom: **{col}**")
                
                # Simpan kolom original untuk perbandingan
                df[f'{col}_Original'] = df[col].astype(str)
                
                # Normalisasi tanggal
                df[col] = df[col].apply(parse_mixed_date)
                
                # Hitung berapa yang berhasil dinormalisasi
                valid_count = df[col].notna().sum()
                total_count = len(df)
                
                st.success(f"   ✅ {valid_count}/{total_count} tanggal berhasil dinormalisasi")
                
                # Tampilkan sample perbandingan
                if valid_count > 0:
                    sample_df = df[[f'{col}_Original', col]].head(5).copy()
                    sample_df.columns = ['Format Asli', 'Format Normalized']
                    st.write("   📋 Sample hasil normalisasi:")
                    st.dataframe(sample_df, use_container_width=True)
            
            st.success(f"🎉 Semua kolom Posting Date berhasil dinormalisasi!")
            
        else:
            st.info("ℹ️ Tidak ada kolom Posting Date yang ditemukan")

        st.markdown("---")

        # ========== TAMBAHAN: Cari dan tambahkan SEMUA kolom yang mengandung "Unnamed" ==========
        unnamed_cols = [col for col in df.columns if 'Unnamed' in col]

        if unnamed_cols:
            st.write(f"🔍 Ditemukan {len(unnamed_cols)} kolom 'Unnamed' yang akan dihapus:")
            for col in unnamed_cols:
                st.write(f"   - {col}")
            # Tambahkan ke daftar cols_to_drop
            cols_to_drop.extend(unnamed_cols)

        # Hapus duplikat dari daftar (karena beberapa Unnamed sudah ada di list manual)
        cols_to_drop = list(set(cols_to_drop))

        # ========== HAPUS KOLOM ORIGINAL SEBELUM DOWNLOAD ==========
        # Hapus kolom "_Original" yang dibuat untuk perbandingan
        original_cols = [col for col in df.columns if col.endswith('_Original')]
        if original_cols:
            cols_to_drop.extend(original_cols)
            st.write(f"🗑️ Menghapus {len(original_cols)} kolom sementara '_Original'")

        # Hanya drop kolom yang ada
        existing_cols_to_drop = [col for col in cols_to_drop if col in df.columns]

        if existing_cols_to_drop:
            df = df.drop(columns=existing_cols_to_drop)
            st.success(f"✅ {len(existing_cols_to_drop)} kolom berhasil dihapus")
            
            # Hitung kolom Unnamed yang dihapus
            unnamed_dropped = [col for col in existing_cols_to_drop if 'Unnamed' in col]
            if unnamed_dropped:
                st.write(f"   ↳ Termasuk {len(unnamed_dropped)} kolom 'Unnamed'")
        else:
            st.write("ℹ️ Tidak ada kolom yang perlu dihapus")

        st.write(f"📊 Total kolom tersisa: {len(df.columns)}")

        # Dapatkan tahun dari data
        def get_year_from_data(df, date_column="Posting Date_APIN"):
            if date_column in df.columns:
                # Cek apakah sudah datetime atau masih string
                if not pd.api.types.is_datetime64_any_dtype(df[date_column]):
                    df[date_column] = pd.to_datetime(df[date_column], errors='coerce')
                year = df[date_column].dt.year.min()
                return int(year) if not pd.isna(year) else datetime.datetime.now().year
            else:
                return datetime.datetime.now().year

        year = get_year_from_data(df)
        output_file = f"DATA PEMBELIAN {year}.xlsx"

        # Export ke Excel
        try:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            buffer.seek(0)

            st.download_button(
                label="📥 Download Excel Final",
                data=buffer,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.success(f"✓ File berhasil dibuat: '{output_file}'")
            st.write(f"   Total baris: {len(df)}")
            st.write(f"   Total kolom: {len(df.columns)}")
            
        except Exception as e:
            st.error(f"❌ Gagal menyimpan file: {e}")

        st.write("=" * 50)
        st.success("🎉 SEMUA TAHAP SELESAI!")