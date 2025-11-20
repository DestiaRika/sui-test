import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import calendar
from datetime import datetime, timedelta
from plotly.subplots import make_subplots

st.set_page_config(page_title="📊 Dashboard Analisis", layout="wide")
st.title("📦 Purchasing Analysis Dashboard")

# Upload hanya 2 file utama
uploaded_file1 = st.file_uploader("Upload File Tahun 1:", type=['xlsx'], key="file1")
uploaded_file2 = st.file_uploader("Upload File Tahun 2 (Opsional):", type=['xlsx'], key="file2")

if 'page' not in st.session_state:
    st.session_state.page = "upload"
if 'analysis_started' not in st.session_state:
    st.session_state.analysis_started = False
if 'last_item_option' not in st.session_state:
    st.session_state.last_item_option = None

def categorize_item(description):
    description = str(description).lower()
    if 'jagung' in description:
        if any(k in description for k in ['argentina','brazil','india','pakistan']):
            return 'Jagung Import'
        return 'Jagung Lokal'
    if 'zak' in description:
        return 'Zak'
    if 'wheat bran' in description:
        return 'Wheat Bran'
    return 'Lainnya'

def extract_year(df):
    """Ekstrak tahun dari kolom Posting Date"""
    # Cek apakah ada kolom Posting Date (sudah diproses) atau Posting Date_APIN (belum diproses)
    date_col = 'Posting Date' if 'Posting Date' in df.columns else 'Posting Date_APIN'
    
    if date_col not in df.columns:
        return []
    
    df = df.copy()
    # Parse berbagai format tanggal: "1/13/2023 12:00:00 AM", "2023-01-13", dll
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    years = df[date_col].dt.year.dropna().astype(int).unique()
    return sorted(years.tolist())

def process_file(file):
    """Proses file Excel dan pisahkan kolom APIN, GRPO, dan SUPPLIER"""
    df_raw = pd.read_excel(file)
    
    # Pisahkan kolom berdasarkan suffix
    apin_cols = [col for col in df_raw.columns if col.endswith('_APIN')]
    grpo_cols = [col for col in df_raw.columns if col.endswith('_GRPO')]
    supplier_cols = [col for col in df_raw.columns if col.endswith('_SUPPLIER')]
    
    # Buat DataFrame untuk Invoice (APIN)
    invoice_df = df_raw[apin_cols].copy()
    # Hapus suffix _APIN dari nama kolom
    invoice_df.columns = [col.replace('_APIN', '') for col in invoice_df.columns]
    
    # Tambahkan kolom Vendor dari SUPPLIER
    if supplier_cols:
        supplier_df = df_raw[supplier_cols].copy()
        supplier_df.columns = [col.replace('_SUPPLIER', '') for col in supplier_df.columns]
        
        # Merge kolom supplier ke invoice_df (karena jumlah baris sama)
        for col in supplier_df.columns:
            if col == 'Vendor Code':
                invoice_df['Vendor Code'] = supplier_df['Vendor Code']
            elif col == 'Vendor Name':
                invoice_df['Vendor'] = supplier_df['Vendor Name']  # Rename ke 'Vendor'
    
    # Parse tanggal di invoice (format: "1/13/2023 12:00:00 AM")
    if 'Posting Date' in invoice_df.columns:
        invoice_df['Posting Date'] = pd.to_datetime(invoice_df['Posting Date'], errors='coerce')
    
    # Buat DataFrame untuk GRPO
    grpo_df = df_raw[grpo_cols].copy()
    # Hapus suffix _GRPO dari nama kolom
    grpo_df.columns = [col.replace('_GRPO', '') for col in grpo_df.columns]
    
    # Parse tanggal di GRPO (jika ada)
    if 'Posting Date' in grpo_df.columns:
        grpo_df['Posting Date'] = pd.to_datetime(grpo_df['Posting Date'], errors='coerce')
    
    # Tambahkan kategori
    if 'Item Description' in invoice_df.columns:
        invoice_df['Grub'] = invoice_df['Item Description'].apply(categorize_item)
    if 'Item Description' in grpo_df.columns:
        grpo_df['Grub'] = grpo_df['Item Description'].apply(categorize_item)
    
    return invoice_df, grpo_df

def refresh_combined_filtered(item_option):
    df1 = st.session_state.invoice_data1
    df2 = st.session_state.invoice_data2 if not st.session_state.invoice_data2.empty else pd.DataFrame()
    f1 = df1[df1['Grub'] == item_option]
    f2 = df2[df2['Grub'] == item_option] if not df2.empty else pd.DataFrame()
    combined = pd.concat([f1, f2], ignore_index=True)
    if not combined.empty and not pd.api.types.is_datetime64_any_dtype(combined['Posting Date']):
        combined['Posting Date'] = pd.to_datetime(combined['Posting Date'], errors='coerce')
    st.session_state.combined_filtered = combined

if st.session_state.page == "upload":
    if st.button('Proses Data') and uploaded_file1 is not None:
        # ========== PROSES FILE 1 ==========
        result = process_file(uploaded_file1)
        invoice_df1, grpo_df1 = result
        st.session_state.invoice_data1 = invoice_df1
        st.session_state.grpo_data1 = grpo_df1
        
        # ========== PROSES FILE 2 (OPSIONAL) ==========
        if uploaded_file2 is not None:
            result2 = process_file(uploaded_file2)
            invoice_df2, grpo_df2 = result2
            st.session_state.invoice_data2 = invoice_df2
            st.session_state.grpo_data2 = grpo_df2
        else:
            st.session_state.invoice_data2 = pd.DataFrame()
            st.session_state.grpo_data2 = pd.DataFrame()

        # ========== EXTRACT YEARS ==========
        years1 = extract_year(st.session_state.invoice_data1)
        if not st.session_state.invoice_data2.empty:
            years2 = extract_year(st.session_state.invoice_data2)
            all_years = sorted(set(years1).union(years2))
        else:
            all_years = years1
        comparison_years = [f"{all_years[i]} vs {all_years[i+1]}" for i in range(len(all_years)-1)]
        st.session_state.available_years = all_years + comparison_years
        st.session_state.analysis_started = False
        st.session_state.combined_filtered = pd.DataFrame()
        st.session_state.last_item_option = None
        st.session_state.page = "filter"
        
        st.success("✅ Data berhasil diproses!")
        st.rerun()

#region Filter page
if st.session_state.page == "filter":
    st.title("Analysis Filter")
    year_option = st.selectbox("Select Year:", st.session_state.available_years)
    item_option = st.selectbox("Select Category:", ["Jagung Import", "Zak", "Jagung Lokal", "Wheat Bran"])
    metric_option = st.selectbox("Filter By:", ["Quantity", "Total"])
    metric_col = "Netto Quantity" if metric_option == "Quantity" else "Total DOC IDR"
    y_label = "Quantity (Netto)" if metric_option == "Quantity" else "Total Pembelian (IDR)"

    if st.button("Start Analysis"):
        refresh_combined_filtered(item_option)
        st.session_state.analysis_started = True
    if st.session_state.analysis_started and st.session_state.last_item_option != item_option:
        refresh_combined_filtered(item_option)
    st.session_state.last_item_option = item_option
    
    if st.session_state.analysis_started:
        combined_filtered = st.session_state.get("combined_filtered", pd.DataFrame()).copy()
        if combined_filtered.empty:
            st.info("Data kosong untuk kombinasi filter ini.")
            st.stop()

       #region Top Produk
        def build_chart(df_year: pd.DataFrame, year_label: int):
            if df_year.empty:
                return None
            if metric_col == "Netto Quantity":
                item_ranking = (
                    df_year.groupby('Item Description', dropna=False)['Netto Quantity']
                    .sum()
                    .sort_values(ascending=True))
                df_ranking = item_ranking.reset_index()
                top_idx = df_ranking['Netto Quantity'].idxmax()
                df_ranking['color'] = ['skyblue'] * len(df_ranking)
                if pd.notna(top_idx):
                    df_ranking.loc[top_idx, 'color'] = 'red'

                df_ranking['DisplayQty'] = df_ranking['Netto Quantity'].apply(lambda x: f"{x:,.0f}")

                fig_local = px.bar(
                    df_ranking,
                    x='Netto Quantity',
                    y='Item Description',
                    orientation='h',
                    text='DisplayQty',
                    color='color',
                    color_discrete_map='identity',
                    title=f'Top Purchased Products in {year_label}',
                    labels={
                        'Netto Quantity': 'Total (Netto)',
                        'Item Description': 'Item Name'})

                fig_local.update_layout(
                    yaxis=dict(title="Item Name"),
                    xaxis=dict(
                        title="Total (Netto)",
                        tickformat=",.0f"),
                    showlegend=False)

            else:
                item_ranking = (
                    df_year.groupby('Item Description', dropna=False)['Total DOC IDR']
                    .sum()
                    .sort_values(ascending=True))
                df_ranking = item_ranking.reset_index()
                top_idx = df_ranking['Total DOC IDR'].idxmax()
                df_ranking['color'] = ['skyblue'] * len(df_ranking)
                if pd.notna(top_idx):
                    df_ranking.loc[top_idx, 'color'] = 'red'

                df_ranking['Formatted DOC'] = df_ranking['Total DOC IDR'].apply(lambda x: f"{x:,.2f}")

                fig_local = px.bar(
                    df_ranking,
                    x='Total DOC IDR',
                    y='Item Description',
                    orientation='h',
                    text='Formatted DOC',
                    color='color',
                    color_discrete_map='identity',
                    title=f'Top Purchased Products (IDR) in {year_label}',
                    labels={
                        'Total DOC IDR': 'Total Purchase (IDR)',
                        'Item Description': 'Item Name'})

                fig_local.update_layout(
                    yaxis=dict(title="Item Name"),
                    xaxis=dict(
                        title="Total Purchase (IDR)",
                        tickformat=",.0f"  ),
                    showlegend=False)

            fig_local.update_yaxes(categoryorder='total ascending')
            return fig_local
        
        if isinstance(year_option, str) and " vs " in year_option:
            y1, y2 = map(int, year_option.split(" vs "))
            df_left  = combined_filtered[combined_filtered['Posting Date'].dt.year == y1]
            df_right = combined_filtered[combined_filtered['Posting Date'].dt.year == y2]
            c1, c2 = st.columns(2)
            with c1:
                st.subheader(f"🟦 years {y1}")
                fig_left = build_chart(df_left, y1)
                if fig_left is not None:
                    st.plotly_chart(fig_left, use_container_width=True)
            with c2:
                st.subheader(f"🟥 years {y2}")
                fig_right = build_chart(df_right, y2)
                if fig_right is not None:
                    st.plotly_chart(fig_right, use_container_width=True)
        else:
            y = int(year_option)
            ydf = combined_filtered[combined_filtered['Posting Date'].dt.year == y]
            st.subheader("🏆 Top Products")
            fig = build_chart(ydf, y)
            if fig is not None:
                st.plotly_chart(fig, use_container_width=True)

#region Top Supplier
        st.subheader("📊 Top Suppliers")
        def top_supplier_dual_full_nominal(df: pd.DataFrame, title_suffix: str = ""):
            sup = (
                df.groupby(['Vendor Code', 'Vendor'], dropna=False)
                .agg({
                    'Netto Quantity': 'sum',
                    'Total DOC IDR': 'sum'})
                .reset_index())

            # URUTKAN BERDASARKAN TOTAL DOC IDR (bukan Netto Quantity)
            sup = sup.sort_values(by='Total DOC IDR', ascending=False).head(10)
            if sup.empty:
                st.info("Tidak ada data supplier untuk periode ini.")
                return

            fig = go.Figure()
            # BAR OREN = TOTAL DOC IDR (di yaxis)
            fig.add_trace(go.Bar(
                x=sup['Vendor Code'], y=sup['Total DOC IDR'], name='Total (IDR)', marker_color='orange',
                hovertemplate=(
                    '<b>Vendor: %{customdata[0]}</b>'
                    '<br>Vendor Code: %{x}'
                    '<br>Total DOC IDR: %{y:,}'
                    '<extra></extra>'),
                customdata=sup[['Vendor']], yaxis='y', offsetgroup='money', alignmentgroup='Vendor'))

            # BAR BIRU = NETTO QUANTITY (di yaxis2 / kanan)
            fig.add_trace(go.Bar(
                x=sup['Vendor Code'], y=sup['Netto Quantity'], name='Quantity (Netto)', marker_color='royalblue',
                hovertemplate=(
                    '<b>Vendor: %{customdata[0]}</b>'
                    '<br>Vendor Code: %{x}'
                    '<br>Quantity: %{y:,}'
                    '<extra></extra>'),
                customdata=sup[['Vendor']], yaxis='y2', offsetgroup='qty', alignmentgroup='Vendor'))

            fig.update_layout(
                title=f"Top Suppliers {title_suffix}".strip(),
                xaxis_title="Vendor",
                yaxis=dict(title="Total DOC IDR", tickformat=",.0f", separatethousands=True, showgrid=True),
                yaxis2=dict(title="Quantity (Netto)", overlaying='y', side='right', showgrid=False, tickformat=",.0f", separatethousands=True),
                barmode='group', bargap=0.15, bargroupgap=0.05, xaxis_tickangle=-45,
                legend=dict(title="Metric", orientation="h", x=0, xanchor="center", y=-0.3, yanchor="top"),
                margin=dict(l=60, r=60, t=60, b=180),)
            st.plotly_chart(fig, use_container_width=True)

        if isinstance(year_option, str) and " vs " in year_option:
            y1, y2 = map(int, year_option.split(" vs "))
            d1 = combined_filtered[combined_filtered['Posting Date'].dt.year == y1]
            d2 = combined_filtered[combined_filtered['Posting Date'].dt.year == y2]
            c1, c2 = st.columns(2)
            with c1:
                top_supplier_dual_full_nominal(d1, f"years {y1}")
            with c2:
                top_supplier_dual_full_nominal(d2, f"years {y2}")

        else:
            y = int(year_option)
            d = combined_filtered[combined_filtered['Posting Date'].dt.year == y]
            top_supplier_dual_full_nominal(d, f"years {y}")

        #region Distribusi Gudang
        st.subheader("🏭 Received Quantity Distribution by Warehouse")

        def pie_wh(df: pd.DataFrame, title_suffix: str = ""):
            wh = (
                df.groupby("Whse", dropna=False)["Quantity"]
                .sum()
                .reset_index()
                .sort_values("Quantity", ascending=False)
            )
            if wh.empty:
                st.info("Tidak ada data warehouse untuk periode ini.")
                return

            total_qty = float(wh["Quantity"].sum())

            fig = px.pie(
                wh,
                values="Quantity",
                names="Whse",
                title=f"Received Quantity Distribution by Warehouse {title_suffix}".strip(),
                hole=0.4
            )

            fig.update_traces(
                textinfo="percent+label",
                textposition="outside",
                hovertemplate="<b>%{label}</b><br>Quantity: %{value:,}<br>Proporsi: %{percent}<extra></extra>"
            )

            # === Layout: legend di kanan dan beri ruang bawah ekstra ===
            fig.update_layout(
                margin=dict(t=100, b=200, l=40, r=200),  # tambahkan b=200 biar lebih lega bawahnya
                legend=dict(
                    orientation="v",
                    y=0.5,
                    x=1.05,
                    xanchor="left",
                    yanchor="middle",
                    title_text="",
                    font=dict(size=12)
                )
            )

            # === Keterangan total (digeser lebih ke bawah agar tidak menimpa chart) ===
            fig.add_annotation(
                x=0.5, y=-0.55, # dari -0.25 jadi -0.35 → lebih ke bawah
                xref="paper", yref="paper",
                text=f"<b>Total Received Quantity:</b> {total_qty:,.0f}",
                showarrow=False,
                align="center",
                font=dict(size=13, color="black")
            )

            st.plotly_chart(fig, use_container_width=True)

        # ==== LOGIKA years ====
        if isinstance(year_option, str) and " vs " in year_option:
            y1, y2 = map(int, year_option.split(" vs "))
            d1 = combined_filtered[combined_filtered['Posting Date'].dt.year == y1]
            d2 = combined_filtered[combined_filtered['Posting Date'].dt.year == y2]
            c1, c2 = st.columns(2)
            with c1:
                pie_wh(d1, f"({y1})")
            with c2:
                pie_wh(d2, f"({y2})")
        else:
            y = int(year_option)
            d = combined_filtered[combined_filtered['Posting Date'].dt.year == y]
            pie_wh(d, f"({y})")



#region Rata-rata Harga Satuan
        st.subheader("⚖️ Average Unit Price per Vendor for Each Item")

        def _calc_unit_price(row):
            qty = row.get('Netto Quantity', 0)
            if not qty:
                return 0
            total_curr = row.get('Total DOC Currency', 0)
            return (total_curr/qty) if (total_curr and total_curr > 0) else (row.get('Total DOC IDR', 0)/qty)

        combined_filtered['Unit Price'] = combined_filtered.apply(_calc_unit_price, axis=1)
        df_sb = st.session_state.get("combined_filtered", pd.DataFrame()).copy()

        if df_sb.empty:
            st.info("Belum ada data yang diproses. Klik **Start Analysis** dulu.")
        else:
            if not pd.api.types.is_datetime64_any_dtype(df_sb['Posting Date']):
                df_sb['Posting Date'] = pd.to_datetime(df_sb['Posting Date'], errors='coerce')
            df_sb = df_sb.dropna(subset=['Posting Date']).copy()
            df_sb = df_sb[df_sb['Grub'] == item_option].copy()
            df_sb['Year'] = df_sb['Posting Date'].dt.year

            def build_sunburst(df_in: pd.DataFrame, chart_title: str):
                if df_in.empty:
                    st.info(f"Tidak ada data untuk {chart_title}.")
                    return
                
                df_in = df_in.copy()
                df_in['Total_Spend_Pick'] = df_in.apply(
                    lambda r: r['Total DOC Currency'] if (
                        r.get('Total DOC Currency', 0) and r['Total DOC Currency'] > 0
                    ) else r.get('Total DOC IDR', 0),
                    axis=1)
                agg_all = (
                    df_in
                    .groupby(['Vendor Code', 'Vendor', 'Item Description'], dropna=False)
                    .agg(
                        Total_Quantity=('Netto Quantity', 'sum'),
                        Total_Spend=('Total_Spend_Pick', 'sum'))
                    .reset_index())
                agg_all['Avg_Unit_Price'] = agg_all.apply(
                    lambda row: (row['Total_Spend'] / row['Total_Quantity']) if row['Total_Quantity'] else 0,
                    axis=1)

                if agg_all.empty:
                    st.info(f"Data agregat kosong untuk {chart_title}.")
                    return

                agg_all['Plot_Quantity'] = (agg_all['Total_Quantity'].clip(lower=1)) ** 0.5
                fig_sb = px.sunburst(
                    agg_all,
                    path=['Vendor', 'Item Description'],  
                    values='Plot_Quantity',
                    color='Avg_Unit_Price',
                    color_continuous_scale='Viridis',
                    custom_data=[
                        'Vendor',
                        'Vendor Code',
                        'Item Description',
                        'Total_Quantity',
                        'Avg_Unit_Price'])

                fig_sb.update_traces(
                    marker=dict(
                        line=dict(color='rgba(0,0,0,0.4)', width=1.2)),
                    textinfo='label',
                    hovertemplate=(
                        "<b>Vendor:</b> %{customdata[0]}<br>"
                        "<b>Vendor Code:</b> %{customdata[1]}<br>"
                        "<b>Item:</b> %{customdata[2]}<br>"
                        "Total Qty: %{customdata[3]:,.0f}<br>"
                        "Avg Unit Price: %{customdata[4]:,.2f}"
                        "<extra></extra>"))

                fig_sb.update_layout(
                    title=chart_title,
                    margin=dict(t=60, l=0, r=0, b=0),
                    coloraxis_colorbar=dict(
                        title="Avg Unit Price",
                        tickformat=",.2f" ))

                st.plotly_chart(fig_sb, use_container_width=True)

            if isinstance(year_option, str) and " vs " in year_option:
                y1, y2 = map(int, year_option.split(" vs "))
                df_pair = df_sb[df_sb['Year'].isin([y1, y2])].copy()
                if df_pair.empty:
                    st.info("Tidak ada data untuk kedua years tersebut pada kategori ini.")
                else:
                    vendor_rank = (
                        df_pair
                        .groupby('Vendor Code', as_index=False)['Netto Quantity']
                        .sum()
                        .sort_values('Netto Quantity', ascending=False))
                    all_vendors_sorted = vendor_rank['Vendor Code'].tolist()

                    picked_vendors = st.multiselect(
                        "Select the vendor to analyze:", options=all_vendors_sorted, default=[],
                        help=f"Chart kiri = {y1}, chart kanan = {y2}")

                    if not picked_vendors:
                        st.warning("Please select at least one vendor first 👆")
                    else:
                        df_pair = df_pair[df_pair['Vendor Code'].isin(picked_vendors)].copy()
                        d1 = df_pair[df_pair['Year'] == y1].copy()
                        d2 = df_pair[df_pair['Year'] == y2].copy()
                        col_left, col_right = st.columns(2)
                        with col_left:
                            build_sunburst(d1, f"Struktur Pembelian {y1}")
                        with col_right:
                            build_sunburst(d2, f"Struktur Pembelian {y2}")
            else:
                yy = int(year_option)
                df_single = df_sb[df_sb['Year'] == yy].copy()
                if df_single.empty:
                    st.info("Tidak ada data setelah filter years & kategori.")
                else:
                    vendor_rank = (
                        df_single
                        .groupby('Vendor Code', as_index=False)['Netto Quantity']
                        .sum()
                        .sort_values('Netto Quantity', ascending=False))
                    all_vendors_sorted = vendor_rank['Vendor Code'].tolist()

                    picked_vendors = st.multiselect(
                        f"Select the vendor to analyze ({yy}):",
                        options=all_vendors_sorted,
                        default=[],
                        help="You can select more than one vendor")

                    if not picked_vendors:
                        st.warning("Please select at least one vendor first 👆")
                    else:
                        df_single = df_single[df_single['Vendor Code'].isin(picked_vendors)].copy()
                        build_sunburst(df_single, f"Struktur Pembelian {yy}")

#region Tren Month
        st.markdown("### 📈 Monthly Purchase Trend")
        source_df = st.session_state.get("combined_filtered", pd.DataFrame()).copy()

        MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        
        def apply_month_order(fig):
            fig.update_xaxes(categoryorder="array", categoryarray=MONTH_ORDER)
            fig.update_yaxes(tickformat=",.0f")
            fig.update_layout(
                yaxis=dict(tickformat=",.0f"),
                legend_title_text="")
            return fig

        def prep_monthly(df_all: pd.DataFrame) -> pd.DataFrame:
            df = df_all.copy()
            if df.empty:
                return pd.DataFrame(columns=["Year","Month","MonthName","Item Description", metric_col])

            # Filter berdasarkan kategori yang dipilih
            df = df[df['Grub'] == item_option].copy()
            
            # Pastikan Posting Date adalah datetime
            if not pd.api.types.is_datetime64_any_dtype(df['Posting Date']):
                df['Posting Date'] = pd.to_datetime(df['Posting Date'], errors='coerce')
            
            # Buang baris dengan Posting Date null
            df = df.dropna(subset=['Posting Date'])

            # Filter tahun sesuai pilihan
            if isinstance(year_option, str) and " vs " in year_option:
                y1, y2 = map(int, year_option.split(" vs "))
                df = df[df['Posting Date'].dt.year.isin([y1, y2])]
                years_to_show = [y1, y2]
            else:
                y = int(year_option)
                df = df[df['Posting Date'].dt.year == y]
                years_to_show = [y]

            if df.empty:
                return pd.DataFrame(columns=["Year","Month","MonthName","Item Description", metric_col])

            # Ekstrak Year dan Month
            df['Year'] = df['Posting Date'].dt.year
            df['Month'] = df['Posting Date'].dt.month
            
            # Map bulan ke nama bulan (English)
            month_map = {
                1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 
                5: "May", 6: "Jun", 7: "Jul", 8: "Aug", 
                9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
            }
            df['MonthName'] = df['Month'].map(month_map)

            # ✅ PERBAIKAN UTAMA: Pastikan metric_col ada dan valid
            if metric_col not in df.columns:
                st.error(f"Kolom '{metric_col}' tidak ditemukan di data!")
                return pd.DataFrame(columns=["Year","Month","MonthName","Item Description", metric_col])
            
            # Konversi metric_col ke numerik, ganti '-' atau string kosong dengan 0
            df[metric_col] = pd.to_numeric(df[metric_col].replace(['-', ''], pd.NA), errors='coerce').fillna(0)

            # ✅ AGREGASI: Group by Year, Month, MonthName, dan Item Description
            agg = (
                df.groupby(['Year', 'Month', 'MonthName', 'Item Description'], dropna=False)
                [metric_col]
                .sum()
                .reset_index()
            )

            # ✅ LENGKAPI DATA: Tambahkan bulan yang kosong dengan nilai 0
            all_months_df = pd.DataFrame({
                'Month': list(range(1, 13)), 
                'MonthName': MONTH_ORDER
            })
            
            # Dapatkan semua item yang ada
            items = agg['Item Description'].dropna().unique().tolist()
            
            # Buat grid lengkap untuk semua kombinasi tahun, bulan, dan item
            complete = []
            for yr in years_to_show:
                for it in items:
                    t = all_months_df.copy()
                    t['Year'] = yr
                    t['Item Description'] = it
                    complete.append(t)
            
            if not complete:
                return agg
            
            grid = pd.concat(complete, ignore_index=True)
            
            # Merge dengan data aktual
            merged = grid.merge(
                agg, 
                on=['Year', 'Month', 'MonthName', 'Item Description'], 
                how='left'
            )
            
            # Fill missing values dengan 0
            merged[metric_col] = merged[metric_col].fillna(0)
            
            # Set MonthName sebagai categorical untuk urutan yang benar
            merged['MonthName'] = pd.Categorical(
                merged['MonthName'], 
                categories=MONTH_ORDER, 
                ordered=True
            )
            
            # Sort berdasarkan Year, MonthName, dan Item
            merged = merged.sort_values(['Year', 'MonthName', 'Item Description']).reset_index(drop=True)
            
            return merged

        # Proses data monthly
        monthly_df = prep_monthly(source_df)
        
        # Debug: Tampilkan raw data untuk verifikasi
        with st.expander("🔍 Debug: Lihat Data Monthly (Raw)"):
            st.dataframe(monthly_df)
        
        # Dapatkan daftar item yang tersedia
        available_items = (
            monthly_df['Item Description']
            .dropna()
            .unique()
            .tolist() if not monthly_df.empty else []
        )
        
        # Sort items alphabetically
        available_items = sorted(available_items)

        # Mode selection
        trend_mode = st.radio(
            "Select Trend Mode:",
            ["Item Trend", "Item Comparison"],
            horizontal=True,
            key="trend_mode_after_start"
        )

        # ✅ MODE 1: Item Trend (Satu chart per item)
        if trend_mode == "Item Trend":
            if not available_items:
                st.info("Tidak ada item untuk ditampilkan.")
            else:
                for it in available_items:
                    df_item = monthly_df[monthly_df["Item Description"] == it].copy()
                    
                    if df_item.empty: 
                        continue
                    
                    # Hitung total untuk item ini
                    total_value = df_item[metric_col].sum()
                    
                    st.subheader(f"📈 {it}")
                    st.caption(f"Total {metric_col}: {total_value:,.0f}")
                    
                    # Buat line chart
                    if isinstance(year_option, str) and " vs " in year_option:
                        fig = px.line(
                            df_item, 
                            x="MonthName", 
                            y=metric_col, 
                            color="Year", 
                            markers=True, 
                            title=f"Monthly Trend – {it}"
                        )
                    else:
                        fig = px.line(
                            df_item, 
                            x="MonthName", 
                            y=metric_col, 
                            markers=True, 
                            title=f"Monthly Trend – {it}"
                        )
                    
                    fig.update_xaxes(title_text="Month")
                    fig.update_yaxes(title_text=y_label, tickformat=",.0f")
                    apply_month_order(fig)
                    st.plotly_chart(fig, use_container_width=True)

        # ✅ MODE 2: Item Comparison (Multiple items dalam satu chart)
        else:  
            st.caption("Check multiple items to compare them in a single chart.")
            
            selected = []
            cols = st.columns(3)
            
            for i, it in enumerate(available_items):
                if cols[i % 3].checkbox(it, key=f"cmp_{item_option}_{year_option}_{i}"):
                    selected.append(it)
            
            if not selected:
                st.warning("Please select at least one item to compare")
            else:
                df_sel = monthly_df[monthly_df["Item Description"].isin(selected)].copy()
                
                if df_sel.empty:
                    st.info("Tidak ada data untuk item yang dipilih.")
                else:
                    # Buat comparison chart
                    if isinstance(year_option, str) and " vs " in year_option:
                        fig = px.line(
                            df_sel, 
                            x="MonthName", 
                            y=metric_col, 
                            color="Item Description", 
                            line_dash="Year", 
                            markers=True, 
                            title=f"Monthly Trend Comparison – {item_option}"
                        )
                    else:
                        fig = px.line(
                            df_sel, 
                            x="MonthName", 
                            y=metric_col, 
                            color="Item Description", 
                            markers=True, 
                            title=f"Monthly Trend Comparison – {item_option}"
                        )
                    
                    fig.update_xaxes(title_text="Month")
                    fig.update_yaxes(title_text=y_label, tickformat=",.0f")
                    apply_month_order(fig)
                    st.plotly_chart(fig, use_container_width=True)
#region Tren Mingguan
        st.markdown("---")
        st.header("📅 Weekly Purchase Trend")

        MONTH_NAMES_ID = ["January", "February", "March", "April", "May", "June",
                            "July", "August", "September", "October", "November", "December"]
        MONTH_NAME_TO_NUM = {name: i+1 for i, name in enumerate(MONTH_NAMES_ID)}

        weekly_source_df = st.session_state.get("combined_filtered", pd.DataFrame()).copy()
        available_months = []
        if not weekly_source_df.empty:
            if not pd.api.types.is_datetime64_any_dtype(weekly_source_df['Posting Date']):
                weekly_source_df['Posting Date'] = pd.to_datetime(weekly_source_df['Posting Date'], errors='coerce')
            tmp = weekly_source_df.dropna(subset=['Posting Date'])
            tmp = tmp[tmp['Grub'] == item_option]
            if isinstance(year_option, str) and " vs " in year_option:
                y1, y2 = map(int, year_option.split(" vs "))
                tmp = tmp[tmp['Posting Date'].dt.year.isin([y1, y2])]
            else:
                y = int(year_option)
                tmp = tmp[tmp['Posting Date'].dt.year == y]
            available_months = sorted(tmp['Posting Date'].dt.month.unique().tolist())

        cols_month = st.columns(3)
        selected_months = []
        for i, mname in enumerate(MONTH_NAMES_ID):
            mnum = MONTH_NAME_TO_NUM[mname]
            col = cols_month[i % 3]
            if mnum in available_months:
                if col.checkbox(mname, key=f"wk_month_{mname}"):
                    selected_months.append(mname)
            else:
                col.markdown(
                    f"""
                    <label style="color:#999; opacity:0.6; user-select:none;">
                    <input type="checkbox" disabled style="margin-right:8px;"> {mname}
                    </label>
                    """,
                    unsafe_allow_html=True)

        weekly_trend_mode = st.radio("Select Weekly Trend Mode:", ["Item Trend", "Item Comparison"], horizontal=True, key="weekly_trend_mode_radio" )

        def _month_week_starts(year: int, month: int, week_start: int = 0):
            first_day = datetime(year, month, 1)
            last_day  = datetime(year, month, calendar.monthrange(year, month)[1])
            delta_back = (first_day.weekday() - week_start) % 7
            start = (first_day - timedelta(days=delta_back)).replace(hour=0, minute=0, second=0, microsecond=0)
            starts = []
            cur = start
            while cur <= last_day:
                starts.append(pd.to_datetime(cur).normalize())
                cur += timedelta(days=7)
            return starts

        def add_week_in_month(df: pd.DataFrame, date_col: str = "Posting Date", week_start: int = 0) -> pd.DataFrame:
            df = df.copy()
            if not pd.api.types.is_datetime64_any_dtype(df[date_col]):
                df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
            df = df.dropna(subset=[date_col])

            delta = (df[date_col].dt.weekday - week_start) % 7
            df["WeekStart"] = (df[date_col] - pd.to_timedelta(delta, unit="D")).dt.normalize()

            df["_Y"] = df[date_col].dt.year
            df["_M"] = df[date_col].dt.month

            parts = []
            for (y, m), sub in df.groupby(["_Y","_M"], sort=False):
                official = _month_week_starts(y, m, week_start=week_start)
                idx_map = {ws: i+1 for i, ws in enumerate(official)}
                parts.append(sub.assign(WeekLabel=sub["WeekStart"].map(idx_map).astype("Int64")))
            df = pd.concat(parts).drop(columns=["_Y","_M"])
            df["WeekLabel"] = df["WeekLabel"].astype(int)
            return df

        def prep_weekly_calendar(df_all: pd.DataFrame, month_num: int, week_start: int = 0) -> pd.DataFrame:
            df = df_all.copy()
            if df.empty:
                return pd.DataFrame(columns=["Year","WeekLabel","WeekName","Item Description", metric_col])

            if not pd.api.types.is_datetime64_any_dtype(df['Posting Date']):
                df['Posting Date'] = pd.to_datetime(df['Posting Date'], errors='coerce')
            df = df.dropna(subset=['Posting Date'])
            df = df[df['Grub'] == item_option].copy()

            if isinstance(year_option, str) and " vs " in year_option:
                y1, y2 = map(int, year_option.split(" vs "))
                df = df[df['Posting Date'].dt.year.isin([y1, y2])]
                years_to_show = [y1, y2]
            else:
                y = int(year_option)
                df = df[df['Posting Date'].dt.year == y]
                years_to_show = [y]

            df = df[df['Posting Date'].dt.month == month_num]
            if df.empty:
                return pd.DataFrame(columns=["Year","WeekLabel","WeekName","Item Description", metric_col])

            df = add_week_in_month(df, date_col="Posting Date", week_start=week_start)
            df["Year"] = df["Posting Date"].dt.year

            agg = (df.groupby(["Year","WeekLabel","Item Description"], dropna=False)[metric_col]
                .sum().reset_index())

            items = df["Item Description"].dropna().unique().tolist()
            grids = []
            for yr in years_to_show:
                n_weeks = len(_month_week_starts(yr, month_num, week_start=week_start))
                for it in items:
                    grids.append(pd.DataFrame({
                        "Year": yr,
                        "Item Description": it,
                        "WeekLabel": list(range(1, n_weeks+1))}))
            full_grid = pd.concat(grids, ignore_index=True)
            merged = full_grid.merge(agg, on=["Year","WeekLabel","Item Description"], how="left")
            merged[metric_col] = merged[metric_col].fillna(0)

            merged["WeekName"] = pd.Categorical(
                merged["WeekLabel"].map(lambda n: f"Minggu {int(n)}"),
                categories=[f"Minggu {i}" for i in sorted(merged["WeekLabel"].unique())],
                ordered=True)
            return merged

        if not selected_months:
            st.info("Please select at least one month to display the weekly trend.")
        else:
            weekly_df_list = []
            for mname in selected_months:
                mnum = MONTH_NAME_TO_NUM[mname]
                df_m = prep_weekly_calendar(weekly_source_df, mnum, week_start=0)  # minggu mulai Senin
                if not df_m.empty:
                    df_m["MonthName"] = mname
                    weekly_df_list.append(df_m)

            weekly_df = (
                pd.concat(weekly_df_list, ignore_index=True)
                if weekly_df_list else
                pd.DataFrame(columns=["Year","WeekName","Item Description", metric_col, "MonthName"]))

            if weekly_trend_mode == "Item Trend":
                if weekly_df.empty:
                    st.info("Tidak ada data untuk ditampilkan.")
                else:
                    st.markdown("### 📊 Weekly Item Trend")
                    for it in sorted(weekly_df['Item Description'].dropna().unique().tolist()):
                        dfi = weekly_df[weekly_df["Item Description"] == it]
                        if dfi.empty:
                            continue
                        sub_title = f"📈 {it}" + ("" if len(selected_months) == 1 else " — per Month")
                        st.subheader(sub_title)

                        if len(selected_months) > 1:
                            if isinstance(year_option, str) and " vs " in year_option:
                                fig_i = px.line(dfi, x="WeekName", y=metric_col, color="Year", line_dash="Year",
                                    facet_col="MonthName", facet_col_wrap=2, markers=True, title=None)
                            else:
                                fig_i = px.line(dfi, x="WeekName", y=metric_col, color="MonthName", markers=True, title=None)
                        else:
                            if isinstance(year_option, str) and " vs " in year_option:
                                fig_i = px.line(dfi, x="WeekName", y=metric_col, color="Year", line_dash="Year", markers=True, title=None)
                            else:
                                fig_i = px.line(dfi, x="WeekName", y=metric_col, markers=True, title=None)

                        fig_i.update_xaxes(title_text="Minggu ke-")
                        fig_i.update_yaxes(title_text=y_label, tickformat=",.0f")
                        fig_i.update_layout(
                            yaxis=dict(tickformat=",.0f"))
                        st.plotly_chart(fig_i, use_container_width=True)

            else:  
                if weekly_df.empty:
                    st.info("Tidak ada data untuk dipilih.")
                else:
                    st.caption("Check multiple items to compare them in a single chart")
                    items_week = sorted(weekly_df['Item Description'].dropna().unique().tolist())
                    selected_items_week = []
                    cols_w = st.columns(3)
                    for i, it in enumerate(items_week):
                        if cols_w[i % 3].checkbox(it, key=f"cmp_week_{it}_{'_'.join(selected_months)}_{year_option}"):
                            selected_items_week.append(it)

                    if not selected_items_week:
                        st.warning("Please select at least one item to compare")
                    else:
                        df_sel_w = weekly_df[weekly_df["Item Description"].isin(selected_items_week)]
                        if df_sel_w.empty:
                            st.info("Tidak ada data untuk item yang dipilih.")
                        else:
                            if isinstance(year_option, str) and " vs " in year_option:
                                fig_c = px.line(df_sel_w, x="WeekName", y=metric_col, color="Item Description", line_dash="Year",
                                    facet_col="MonthName" if len(selected_months) > 1 else None, facet_col_wrap=2 if len(selected_months) > 1 else None,
                                    markers=True, title="Weekly Trend Comparison")
                            else:
                                if len(selected_months) > 1:
                                    fig_c = px.line(df_sel_w, x="WeekName", y=metric_col, color="Item Description", line_dash="MonthName",
                                        markers=True, title="Weekly Trend Comparison")
                                else:
                                    fig_c = px.line(df_sel_w, x="WeekName", y=metric_col, color="Item Description",
                                        markers=True, title=f"Weekly Trend Comparison ({selected_months[0]})")

                            fig_c.update_xaxes(title_text="Minggu ke-")
                            fig_c.update_yaxes(title_text=y_label, tickformat=",.0f")
                            fig_c.update_layout(
                                yaxis=dict(tickformat=",.0f"))
                            st.plotly_chart(fig_c, use_container_width=True)     

#region rata rata
#region Rata-rata Uang per Quantity per Bulan (Fokus hanya pada IDR)
        st.markdown("---")
        st.header("📅 Monthly Average Purchase Value by Quantity")

        # Ambil data dari session state
        invoice_df = st.session_state.get("combined_filtered", pd.DataFrame()).copy()

        # Cek apakah GRPO sudah diproses
        if 'grpo_data1' not in st.session_state:
            st.info("GRPO data belum diproses. Silakan upload dan proses file terlebih dahulu.")
            st.stop()

        # Gabungkan GRPO data
        grpo1 = st.session_state.get("grpo_data1", pd.DataFrame()).copy()
        grpo2 = st.session_state.get("grpo_data2", pd.DataFrame()).copy() if 'grpo_data2' in st.session_state else pd.DataFrame()

        if not grpo2.empty:
            grpo_df = pd.concat([grpo1, grpo2], ignore_index=True)
        else:
            grpo_df = grpo1.copy()

        # Tambahkan kategori ke GRPO jika belum ada
        if 'Grub' not in grpo_df.columns and 'Item Description' in grpo_df.columns:
            grpo_df['Grub'] = grpo_df['Item Description'].apply(categorize_item)

        if invoice_df.empty:
            st.info("Tidak ada data Invoice untuk ditampilkan.")
            st.stop()
        elif grpo_df.empty:
            st.info("Tidak ada data GRPO untuk ditampilkan.")
            st.stop()

        # Pastikan kolom tanggal sudah datetime
        if not pd.api.types.is_datetime64_any_dtype(invoice_df["Posting Date"]):
            invoice_df["Posting Date"] = pd.to_datetime(invoice_df["Posting Date"], errors="coerce")

        if "Posting Date" in grpo_df.columns:
            if not pd.api.types.is_datetime64_any_dtype(grpo_df["Posting Date"]):
                grpo_df["Posting Date"] = pd.to_datetime(grpo_df["Posting Date"], errors="coerce")

        # Drop NaT
        invoice_df = invoice_df.dropna(subset=["Posting Date"]).copy()
        if "Posting Date" in grpo_df.columns:
            grpo_df = grpo_df.dropna(subset=["Posting Date"]).copy()

        # Filter berdasarkan item yang dipilih
        invoice_df = invoice_df[invoice_df["Grub"] == item_option].copy()
        if "Grub" in grpo_df.columns:
            grpo_df = grpo_df[grpo_df["Grub"] == item_option].copy()

        # Nama bulan
        MONTH_NAMES_ID = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]

        invoice_df["Month"] = invoice_df["Posting Date"].dt.month.apply(lambda x: MONTH_NAMES_ID[x - 1])
        if "Posting Date" in grpo_df.columns:
            grpo_df["Month"] = grpo_df["Posting Date"].dt.month.apply(lambda x: MONTH_NAMES_ID[x - 1])

        # Fungsi untuk menghitung rata-rata bulanan
        def calc_monthly_avg(inv_df, grpo_df, year):
            """Menghitung rata-rata pembelian per bulan untuk tahun tertentu"""
            # Filter per tahun
            inv_year = inv_df[inv_df["Posting Date"].dt.year == year].copy()
            grpo_year = grpo_df[grpo_df["Posting Date"].dt.year == year].copy() if "Posting Date" in grpo_df.columns else grpo_df.copy()
            
            # Ambil Discount & Cash Discount UNIK per Doc Number
            discount_unique = inv_year.drop_duplicates(subset=["Doc Number"], keep="first")[
                ["Doc Number", "Month", "Discount", "Cash Discount (FC)"]
            ].copy()
            
            # Isi NaN dengan 0
            discount_unique["Discount"] = discount_unique["Discount"].fillna(0) if "Discount" in discount_unique.columns else 0
            discount_unique["Cash Discount (FC)"] = discount_unique["Cash Discount (FC)"].fillna(0) if "Cash Discount (FC)" in discount_unique.columns else 0
            
            # Aggregate Discount per bulan
            discount_monthly = discount_unique.groupby("Month").agg({
                "Discount": "sum",
                "Cash Discount (FC)": "sum"
            }).reset_index()
            
            # Aggregate kolom lainnya
            inv_monthly = inv_year.groupby("Month").agg({
                "Netto Quantity": "sum",
                "Total DOC IDR": "sum",
                "Total DOC Currency": "sum"
            }).reset_index()
            
            # Merge dengan discount
            inv_monthly = pd.merge(inv_monthly, discount_monthly, on="Month", how="left").fillna(0)
            
            # Aggregate GRPO per bulan
            if not grpo_year.empty and "Month" in grpo_year.columns:
                grpo_monthly = grpo_year.groupby("Month").agg({
                    "Freight": "sum",
                    "Total Freight Charges (FC)": "sum"
                }).reset_index()
                monthly = pd.merge(inv_monthly, grpo_monthly, on="Month", how="left").fillna(0)
            else:
                monthly = inv_monthly.copy()
                monthly["Freight"] = 0
                monthly["Total Freight Charges (FC)"] = 0
            
            # === HITUNG RATA-RATA IDR ===
            monthly["IDR"] = monthly.apply(
                lambda row: (
                    (row["Total DOC IDR"] + row["Freight"] - row["Discount"]) / row["Netto Quantity"]
                    if (row["Total DOC IDR"] > 0 or row["Freight"] > 0) and row["Netto Quantity"] > 0
                    else 0
                ),
                axis=1
            )
            
            # === HITUNG RATA-RATA USD ===
            monthly["USD"] = monthly.apply(
                lambda row: (
                    (row["Total DOC Currency"] + row["Total Freight Charges (FC)"] - row["Cash Discount (FC)"]) / row["Netto Quantity"]
                    if (row["Total DOC Currency"] > 0 or row["Total Freight Charges (FC)"] > 0) and row["Netto Quantity"] > 0
                    else 0
                ),
                axis=1
            )
            
            # Replace inf/nan dengan 0
            monthly["IDR"] = monthly["IDR"].replace([float('inf'), -float('inf')], 0).fillna(0)
            monthly["USD"] = monthly["USD"].replace([float('inf'), -float('inf')], 0).fillna(0)
            
            return monthly[["Month", "Netto Quantity", "Total DOC IDR", "Total DOC Currency", "Freight", "Total Freight Charges (FC)", "Discount", "Cash Discount (FC)", "USD", "IDR"]]

        # Fungsi untuk menghitung rata-rata tahunan
        def calc_yearly_avg(inv_df, grpo_df, year):
            """Menghitung rata-rata pembelian tahunan"""
            inv_year = inv_df[inv_df["Posting Date"].dt.year == year].copy()
            grpo_year = grpo_df[grpo_df["Posting Date"].dt.year == year].copy() if "Posting Date" in grpo_df.columns else grpo_df.copy()
            
            # Total Quantity dan DOC
            total_qty = inv_year["Netto Quantity"].sum()
            total_doc_idr = inv_year["Total DOC IDR"].sum()
            total_doc_usd = inv_year["Total DOC Currency"].sum()
            
            # Discount & Cash Discount UNIK
            discount_unique = inv_year.drop_duplicates(subset=["Doc Number"], keep="first")
            total_discount_idr = discount_unique["Discount"].fillna(0).sum() if "Discount" in discount_unique.columns else 0
            total_discount_usd = discount_unique["Cash Discount (FC)"].fillna(0).sum() if "Cash Discount (FC)" in discount_unique.columns else 0
            
            # Total GRPO
            total_freight_idr = grpo_year["Freight"].sum() if "Freight" in grpo_year.columns and not grpo_year.empty else 0
            total_freight_usd = grpo_year["Total Freight Charges (FC)"].sum() if "Total Freight Charges (FC)" in grpo_year.columns and not grpo_year.empty else 0
            
            # === HITUNG AVERAGE IDR ===
            if (total_doc_idr > 0 or total_freight_idr > 0) and total_qty > 0:
                avg_idr = (total_doc_idr + total_freight_idr - total_discount_idr) / total_qty
            else:
                avg_idr = 0
            
            # === HITUNG AVERAGE USD ===
            if (total_doc_usd > 0 or total_freight_usd > 0) and total_qty > 0:
                avg_usd = (total_doc_usd + total_freight_usd - total_discount_usd) / total_qty
            else:
                avg_usd = 0
            
            return {
                "Netto Quantity": total_qty,
                "Total DOC IDR": total_doc_idr,
                "Total DOC Currency": total_doc_usd,
                "Freight": total_freight_idr,
                "Total Freight Charges (FC)": total_freight_usd,
                "Discount": total_discount_idr,
                "Cash Discount (FC)": total_discount_usd,
                "USD": avg_usd,
                "IDR": avg_idr
            }

        # === MODE PERBANDINGAN TAHUN (VS) ===
        if isinstance(year_option, str) and " vs " in year_option:
            y1, y2 = map(int, year_option.split(" vs "))
            
            # Filter data per tahun
            invoice_df = invoice_df[invoice_df["Posting Date"].dt.year.isin([y1, y2])]
            grpo_df = grpo_df[grpo_df["Posting Date"].dt.year.isin([y1, y2])] if "Posting Date" in grpo_df.columns else grpo_df
            
            # Hitung untuk kedua tahun
            detail_y1 = calc_monthly_avg(invoice_df, grpo_df, y1)
            detail_y2 = calc_monthly_avg(invoice_df, grpo_df, y2)

            # === TABEL DETAIL (SUM) - Gabungkan kedua tahun ===
            detail_y1_display = detail_y1[["Month", "Netto Quantity", "Total DOC IDR", "Total DOC Currency", "Freight", "Total Freight Charges (FC)", "Discount", "Cash Discount (FC)"]].rename(columns={
                "Netto Quantity": f"Sum Netto Quantity ({y1})",
                "Total DOC IDR": f"Sum Total DOC IDR ({y1})",
                "Total DOC Currency": f"Sum Total DOC Currency ({y1})",
                "Freight": f"Sum Freight ({y1})",
                "Total Freight Charges (FC)": f"Sum Total Freight Charges FC ({y1})",
                "Discount": f"Discount IDR ({y1})",
                "Cash Discount (FC)": f"Discount USD ({y1})"
            })
            
            detail_y2_display = detail_y2[["Month", "Netto Quantity", "Total DOC IDR", "Total DOC Currency", "Freight", "Total Freight Charges (FC)", "Discount", "Cash Discount (FC)"]].rename(columns={
                "Netto Quantity": f"Sum Netto Quantity ({y2})",
                "Total DOC IDR": f"Sum Total DOC IDR ({y2})",
                "Total DOC Currency": f"Sum Total DOC Currency ({y2})",
                "Freight": f"Sum Freight ({y2})",
                "Total Freight Charges (FC)": f"Sum Total Freight Charges FC ({y2})",
                "Discount": f"Discount IDR ({y2})",
                "Cash Discount (FC)": f"Discount USD ({y2})"
            })

            # Merge kedua tahun
            detail_compare = pd.merge(detail_y1_display, detail_y2_display, on="Month", how="outer")

            # Urutkan berdasarkan bulan
            detail_compare["Month"] = pd.Categorical(detail_compare["Month"], categories=MONTH_NAMES_ID, ordered=True)
            detail_compare = detail_compare.sort_values("Month")

            # === YEARLY TOTAL untuk detail ===
            yearly_y1 = calc_yearly_avg(invoice_df, grpo_df, y1)
            yearly_y2 = calc_yearly_avg(invoice_df, grpo_df, y2)

            total_row_detail = pd.DataFrame([{
                "Month": "Yearly Total",
                f"Sum Netto Quantity ({y1})": yearly_y1["Netto Quantity"],
                f"Sum Total DOC IDR ({y1})": yearly_y1["Total DOC IDR"],
                f"Sum Total DOC Currency ({y1})": yearly_y1["Total DOC Currency"],
                f"Sum Freight ({y1})": yearly_y1["Freight"],
                f"Sum Total Freight Charges FC ({y1})": yearly_y1["Total Freight Charges (FC)"],
                f"Discount IDR ({y1})": yearly_y1["Discount"],
                f"Discount USD ({y1})": yearly_y1["Cash Discount (FC)"],
                f"Sum Netto Quantity ({y2})": yearly_y2["Netto Quantity"],
                f"Sum Total DOC IDR ({y2})": yearly_y2["Total DOC IDR"],
                f"Sum Total DOC Currency ({y2})": yearly_y2["Total DOC Currency"],
                f"Sum Freight ({y2})": yearly_y2["Freight"],
                f"Sum Total Freight Charges FC ({y2})": yearly_y2["Total Freight Charges (FC)"],
                f"Discount IDR ({y2})": yearly_y2["Discount"],
                f"Discount USD ({y2})": yearly_y2["Cash Discount (FC)"]
            }])

            detail_compare = pd.concat([detail_compare, total_row_detail], ignore_index=True).fillna(0)

            # Format angka untuk detail
            for col in detail_compare.columns:
                if col != "Month":
                    detail_compare[col] = detail_compare[col].apply(
                        lambda x: f"{x:,.2f}" if pd.notna(x) and x != 0 else "0.00"
                    )

            st.subheader("📊 Detail Sum by Month")
            st.dataframe(detail_compare, use_container_width=True, hide_index=True)

            # === TABEL AVERAGE (USD & IDR) ===
            avg_y1 = detail_y1[["Month", "USD", "IDR"]].rename(columns={"USD": f"USD {y1}", "IDR": f"IDR {y1}"})
            avg_y2 = detail_y2[["Month", "USD", "IDR"]].rename(columns={"USD": f"USD {y2}", "IDR": f"IDR {y2}"})
            
            # Merge
            monthly_compare = pd.merge(avg_y1, avg_y2, on="Month", how="outer")
            
            # Urutkan berdasarkan bulan
            monthly_compare["Month"] = pd.Categorical(monthly_compare["Month"], categories=MONTH_NAMES_ID, ordered=True)
            monthly_compare = monthly_compare.sort_values("Month")
            
            total_row_avg = pd.DataFrame([{
                "Month": "Yearly Average",
                f"USD {y1}": yearly_y1["USD"],
                f"IDR {y1}": yearly_y1["IDR"],
                f"USD {y2}": yearly_y2["USD"],
                f"IDR {y2}": yearly_y2["IDR"]
            }])
            
            # Gabungkan
            monthly_compare = pd.concat([monthly_compare, total_row_avg], ignore_index=True).fillna("-")
            
            # Format angka
            for col in monthly_compare.columns:
                if col == "Month":
                    continue
                if "USD" in col:
                    monthly_compare[col] = monthly_compare[col].apply(
                        lambda x: f"${x:,.2f}" if x != "-" and pd.notna(x) and x > 0 else "-"
                    )
                elif "IDR" in col:
                    monthly_compare[col] = monthly_compare[col].apply(
                        lambda x: f"Rp{x:,.2f}" if x != "-" and pd.notna(x) and x > 0 else "-"
                    )
            
            st.subheader("💰 Average USD & IDR by Month")
            st.dataframe(monthly_compare, use_container_width=True, hide_index=True)

        # === MODE SATU TAHUN ===
        else:
            y = int(year_option)
            
            # Filter per tahun
            invoice_year = invoice_df[invoice_df["Posting Date"].dt.year == y].copy()
            grpo_year = grpo_df[grpo_df["Posting Date"].dt.year == y].copy() if "Posting Date" in grpo_df.columns else grpo_df.copy()
            
            if invoice_year.empty:
                st.info("Tidak ada data Invoice untuk item dan tahun yang dipilih.")
            else:
                # Hitung monthly average (dengan detail)
                monthly_detail = calc_monthly_avg(invoice_df, grpo_df, y)
                
                # Urutkan
                monthly_detail["Month"] = pd.Categorical(monthly_detail["Month"], categories=MONTH_NAMES_ID, ordered=True)
                monthly_detail = monthly_detail.sort_values("Month")
                
                # === YEARLY AVERAGE ===
                yearly_avg = calc_yearly_avg(invoice_df, grpo_df, y)

                # === TABEL DETAIL (SUM) ===
                detail_table = monthly_detail[["Month", "Netto Quantity", "Total DOC IDR", "Total DOC Currency", "Freight", "Total Freight Charges (FC)", "Discount", "Cash Discount (FC)"]].copy()
                detail_table = detail_table.rename(columns={
                    "Netto Quantity": "Sum Netto Quantity",
                    "Total DOC IDR": "Sum Total DOC IDR",
                    "Total DOC Currency": "Sum Total DOC Currency",
                    "Freight": "Sum Freight",
                    "Total Freight Charges (FC)": "Sum Total Freight Charges FC",
                    "Discount": "Discount IDR",
                    "Cash Discount (FC)": "Discount USD"
                })

                # Tambah row total
                total_row_detail = pd.DataFrame([{
                    "Month": "Yearly Total",
                    "Sum Netto Quantity": yearly_avg["Netto Quantity"],
                    "Sum Total DOC IDR": yearly_avg["Total DOC IDR"],
                    "Sum Total DOC Currency": yearly_avg["Total DOC Currency"],
                    "Sum Freight": yearly_avg["Freight"],
                    "Sum Total Freight Charges FC": yearly_avg["Total Freight Charges (FC)"],
                    "Discount IDR": yearly_avg["Discount"],
                    "Discount USD": yearly_avg["Cash Discount (FC)"]
                }])

                detail_table = pd.concat([detail_table, total_row_detail], ignore_index=True).fillna(0)

                # Format angka
                for col in ["Sum Netto Quantity", "Sum Total DOC IDR", "Sum Total DOC Currency", "Sum Freight", "Sum Total Freight Charges FC", "Discount IDR", "Discount USD"]:
                    detail_table[col] = detail_table[col].apply(
                        lambda x: f"{x:,.2f}" if pd.notna(x) and x != 0 else "0.00"
                    )

                st.subheader("📊 Detail Sum per Month")
                st.dataframe(detail_table, use_container_width=True, hide_index=True)

                # === TABEL AVERAGE (USD & IDR) ===
                tampil_tabel = monthly_detail[["Month", "USD", "IDR"]].copy()
                
                total_row = pd.DataFrame([{
                    "Month": "Yearly Average",
                    "USD": yearly_avg["USD"],
                    "IDR": yearly_avg["IDR"]
                }])
                
                # Gabungkan
                tampil_tabel = pd.concat([tampil_tabel, total_row], ignore_index=True).fillna("-")
                
                # Format angka
                tampil_tabel["USD"] = tampil_tabel["USD"].apply(
                    lambda x: f"${x:,.2f}" if x != "-" and pd.notna(x) and x > 0 else "-"
                )
                tampil_tabel["IDR"] = tampil_tabel["IDR"].apply(
                    lambda x: f"Rp{x:,.2f}" if x != "-" and pd.notna(x) and x > 0 else "-"
                )
                
                st.subheader("💰 Average USD & IDR per Month")
                st.dataframe(tampil_tabel, use_container_width=True, hide_index=True)