import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import calendar
from datetime import datetime, timedelta
from plotly.subplots import make_subplots

# ===============================
# KONFIGURASI DASHBOARD
# ===============================
st.set_page_config(page_title="📊 Dashboard Analisis", layout="wide")
st.title("📦 Sales Analysis Dashboard")

uploaded_file1 = st.file_uploader("Upload File Tahun 1:", type=['xlsx'], key="file1")
uploaded_file2 = st.file_uploader("Upload File Tahun 2 (Opsional):", type=['xlsx'], key="file2")
customerdata_file = st.file_uploader('Customer Data File (opsional):', type=['xlsx'], accept_multiple_files=False)

# ===============================
# INISIALISASI SESSION STATE
# ===============================
if 'page' not in st.session_state:
    st.session_state.page = "upload"
if 'analysis_started' not in st.session_state:
    st.session_state.analysis_started = False
if 'last_item_option' not in st.session_state:
    st.session_state.last_item_option = None

# ===============================
# FUNGSI KATEGORISASI ITEM
# ===============================
def categorize_item(description: str) -> str:
    description = str(description).lower().strip()
    
    # Cek exact match untuk Jagung Pecah M60, M50, dan M16 @25 (masuk kategori Jagung)
    if description == "jagung pecah m60" or description == "jagung pecah m50" or description == "jagung pecah m16 @25":
        return "Jagung"
    
    # Baru cek keywords untuk Jagung Pecah
    jagung_pecah_keywords = [
        "jagung pecah", "ampok @50", "corn grits fgiiia", "corn grits fgiiaa",
        "ingredient", "js"
    ]
    if any(k in description for k in jagung_pecah_keywords):
        return "Jagung Pecah"
    
    # Cek kategori Jagung
    if "jagung" in description or "ampok" in description or "ratu merak" in description:
        return "Jagung"
    
    # Cek Wheat Bran
    if "wheat bran" in description:
        return "Wheat Bran"
    
    return "Lainnya"

# ===============================
# UTIL
# ===============================
def _to_num(series_like):
    return pd.to_numeric(pd.Series(series_like).replace('-', pd.NA), errors='coerce')

# ===============================
# PROSES FILE - FIXED VERSION
# ===============================
def process_file(file):
    """Proses file Excel dan pisahkan kolom DO, CMD, dan ARIN"""
    df_raw = pd.read_excel(file)
    
    # Pisahkan kolom berdasarkan suffix untuk DO, CMD, dan ARIN
    do_cols = [col for col in df_raw.columns if col.endswith('_DO')]
    cmd_cols = [col for col in df_raw.columns if col.endswith('_CMD')]
    arin_cols = [col for col in df_raw.columns if col.endswith('_ARIN')]
    
    # Buat DataFrame untuk DO (misalnya untuk Freight IDR)
    do_df = df_raw[do_cols].copy()
    # Hapus suffix _DO dari nama kolom
    do_df.columns = [col.replace('_DO', '') for col in do_df.columns]
    
    # Buat DataFrame untuk CMD (misalnya untuk Customer Code dan Customer Name)
    cmd_df = df_raw[cmd_cols].copy()
    cmd_df.columns = [col.replace('_CMD', '') for col in cmd_df.columns]
    
    # Buat DataFrame untuk ARIN (misalnya untuk data transaksi utama)
    arin_df = df_raw[arin_cols].copy()
    arin_df.columns = [col.replace('_ARIN', '') for col in arin_df.columns]

    # Pastikan kolom tanggal sudah diubah menjadi datetime
    if 'Posting Date' in arin_df.columns:
        arin_df['Posting Date'] = pd.to_datetime(arin_df['Posting Date'], errors='coerce')
    
    if 'Posting Date' in do_df.columns:
        do_df['Posting Date'] = pd.to_datetime(do_df['Posting Date'], errors='coerce')
    
    if 'Posting Date' in cmd_df.columns:
        cmd_df['Posting Date'] = pd.to_datetime(cmd_df['Posting Date'], errors='coerce')

    # ✅ Bersihkan kolom numerik untuk ARIN
    numeric_cols_arin = ['Total LC', 'Netto Quantity', 'Discount']
    for col in numeric_cols_arin:
        if col in arin_df.columns:
            arin_df[col] = pd.to_numeric(arin_df[col].replace('-', pd.NA), errors='coerce').fillna(0)
    
    # ✅ Bersihkan kolom numerik untuk DO (misalnya Freight dan lainnya)
    numeric_cols_do = ['Freight']
    for col in numeric_cols_do:
        if col in do_df.columns:
            do_df[col] = pd.to_numeric(do_df[col].replace('-', pd.NA), errors='coerce').fillna(0)
    
    # ✅ Bersihkan kolom numerik untuk CMD (misalnya jika ada data numerik lainnya)
    numeric_cols_cmd = []
    for col in cmd_df.columns:
        if col not in ['Customer Code', 'Customer Name']:
            numeric_cols_cmd.append(col)
    
    for col in numeric_cols_cmd:
        if col in cmd_df.columns:
            cmd_df[col] = pd.to_numeric(cmd_df[col].replace('-', pd.NA), errors='coerce').fillna(0)

    # Tambahkan kategori (Grub) jika ada kolom Item Description
    if 'Item Description' in arin_df.columns:
        arin_df['Grub'] = arin_df['Item Description'].apply(categorize_item)
    
    if 'Item Description' in do_df.columns:
        do_df['Grub'] = do_df['Item Description'].apply(categorize_item)
    
    if 'Item Description' in cmd_df.columns:
        cmd_df['Grub'] = cmd_df['Item Description'].apply(categorize_item)

    # Kembalikan hasil dalam bentuk tuple (DO DataFrame, CMD DataFrame, ARIN DataFrame)
    return do_df, cmd_df, arin_df

# ===============================
# FUNGSI EKSTRAK TAHUN
# ===============================
def extract_year(df: pd.DataFrame):
    if 'Posting Date' not in df.columns:
        return []
    s = pd.to_datetime(df['Posting Date'], errors='coerce')
    years = s.dt.year.dropna().astype(int).unique()
    return sorted(years.tolist())

# ===============================
# FUNGSI FILTER KOMBINASI DATA
# ===============================
def refresh_combined_filtered(item_option: str):
    df1 = st.session_state.invoice_data1
    df2 = st.session_state.invoice_data2 if not st.session_state.invoice_data2.empty else pd.DataFrame()

    f1 = df1[df1['Grub'] == item_option]
    f2 = df2[df2['Grub'] == item_option] if not df2.empty else pd.DataFrame()
    combined = pd.concat([f1, f2], ignore_index=True)

    if 'Posting Date' in combined.columns:
        combined['Posting Date'] = pd.to_datetime(combined['Posting Date'], errors='coerce')

    st.session_state.combined_filtered = combined

# ===============================
# HALAMAN UPLOAD FILE - FIXED
# ===============================
if st.session_state.page == "upload":
    if st.button('Proses Data') and uploaded_file1 is not None:
        # Proses file pertama
        do1, cmd1, arin1 = process_file(uploaded_file1)
        st.session_state.invoice_data1 = arin1  # ARIN sebagai invoice data
        st.session_state.cmd_data1 = cmd1
        st.session_state.arin_data1 = arin1
        st.session_state.do_data1 = do1  # DO data langsung disimpan

        # Extract available years dari ARIN data pertama
        years1 = extract_year(arin1)

        # Jika ada file kedua
        if uploaded_file2 is not None:
            do2, cmd2, arin2 = process_file(uploaded_file2)
            st.session_state.invoice_data2 = arin2
            st.session_state.cmd_data2 = cmd2
            st.session_state.arin_data2 = arin2
            st.session_state.do_data2 = do2
            
            # Extract available years dari ARIN data kedua
            years2 = extract_year(arin2)
            all_years = sorted(set(years1 + years2))
            
            # Buat opsi perbandingan
            if len(all_years) >= 2:
                vs_options = [f"{all_years[i]} vs {all_years[j]}" 
                             for i in range(len(all_years)) 
                             for j in range(i+1, len(all_years))]
                st.session_state.available_years = all_years + vs_options
            else:
                st.session_state.available_years = all_years
        else:
            st.session_state.invoice_data2 = pd.DataFrame()
            st.session_state.do_data2 = pd.DataFrame()
            st.session_state.available_years = years1

        st.success("✅ Data berhasil diproses! Lanjutkan ke Filter.")
        st.session_state.page = "filter"
        st.rerun()

# ===============================
# HALAMAN FILTER & ANALISIS
# ===============================
if st.session_state.page == "filter":
    st.header("🔍 Filter Analisis")

    if 'available_years' not in st.session_state or not st.session_state.available_years:
        st.warning("Tidak ada data tahun yang ditemukan. Silakan upload ulang file.")
        st.stop()

    colf1, colf2, colf3, colf4 = st.columns(4)
    with colf1:
        year_option = st.selectbox("Pilih Tahun:", st.session_state.available_years, key="year_opt")
    with colf2:
        item_option = st.selectbox("Pilih Kategori:", ["Jagung", "Jagung Pecah", "Wheat Bran"], key="item_opt")
    with colf3:
        metric_option = st.selectbox(
            "Filter Berdasarkan:",
            ["Quantity", "Total"],
            key="metric_opt"
        )

    # tentukan kolom metrik
    if metric_option == "Quantity":
        metric_col = "Netto Quantity"
        y_label = "Quantity (Netto)"
    else:
        metric_col = "Total LC"
        y_label = "Total Penjualan"

    if st.button("Mulai Analisis", key="start_btn"):
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

        #region top product
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

                fig_local = px.bar(df_ranking, x='Netto Quantity', y='Item Description', orientation='h',
                    text='DisplayQty', color='color', color_discrete_map='identity', title=f'Top Sold Products in {year_label}',
                    labels={'Netto Quantity': 'Total (Netto)', 'Item Description': 'Item Name'})

                fig_local.update_layout(
                    yaxis=dict(title="Item Name"),
                    xaxis=dict(title="Total (Netto)", tickformat=",.0f"), showlegend=False)

            else:
                item_ranking = (
                    df_year.groupby('Item Description', dropna=False)['Total LC']
                    .sum()
                    .sort_values(ascending=True))
                df_ranking = item_ranking.reset_index()
                top_idx = df_ranking['Total LC'].idxmax()
                df_ranking['color'] = ['skyblue'] * len(df_ranking)
                if pd.notna(top_idx):
                    df_ranking.loc[top_idx, 'color'] = 'red'

                df_ranking['Formatted DOC'] = df_ranking['Total LC'].apply(lambda x: f"{x:,.2f}")

                fig_local = px.bar(df_ranking, x='Total LC', y='Item Description', orientation='h',
                    text='Formatted DOC', color='color', color_discrete_map='identity', title=f'Top Sales Products (LC) in {year_label}',
                    labels={'Total LC': 'Total (LC)', 'Item Description': 'Item Name'})

                fig_local.update_layout(
                    yaxis=dict(title="Item Name"),
                    xaxis=dict(title="Total Sales (LC)", tickformat=",.0f"), showlegend=False)

            fig_local.update_yaxes(categoryorder='total ascending')
            return fig_local
        
        if isinstance(year_option, str) and " vs " in year_option:
            y1, y2 = map(int, year_option.split(" vs "))
            df_left = combined_filtered[combined_filtered['Posting Date'].dt.year == y1]
            df_right = combined_filtered[combined_filtered['Posting Date'].dt.year == y2]
            c1, c2 = st.columns(2)
            with c1:
                st.subheader(f"🟦 Tahun {y1}")
                fig_left = build_chart(df_left, y1)
                if fig_left is not None:
                    st.plotly_chart(fig_left, use_container_width=True)
            with c2:
                st.subheader(f"🟥 Tahun {y2}")
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

        #region warehouse
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

            fig.update_layout(
                margin=dict(t=100, b=100, l=40, r=200),
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

            fig.add_annotation(
                x=0.5, y=-0.30,
                xref="paper", yref="paper",
                text=f"<b>Total Received Quantity:</b> {total_qty:,.0f}",
                showarrow=False,
                align="center",
                font=dict(size=13, color="black")
            )

            st.plotly_chart(fig, use_container_width=True)

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

        #region Tren Month
        st.markdown("### 📈 Monthly Sales Trend")
        source_df = st.session_state.get("combined_filtered", pd.DataFrame()).copy()

        MONTH_ORDER = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"]
        def apply_month_order(fig):
            fig.update_xaxes(categoryorder="array", categoryarray=MONTH_ORDER)
            fig.update_yaxes(tickformat=",.0f")
            fig.update_layout(yaxis=dict(tickformat=",.0f"), legend_title_text="")
            return fig

        def prep_monthly(df_all: pd.DataFrame) -> pd.DataFrame:
            df = df_all.copy()
            if df.empty:
                return pd.DataFrame(columns=["Year","Month","MonthName","Item Description", metric_col])

            df = df[df['Grub'] == item_option].copy()
            if not pd.api.types.is_datetime64_any_dtype(df['Posting Date']):
                df['Posting Date'] = pd.to_datetime(df['Posting Date'], errors='coerce')
            df = df.dropna(subset=['Posting Date'])

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

            df['Year'] = df['Posting Date'].dt.year
            df['Month'] = df['Posting Date'].dt.month
            month_map = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"Mei",6:"Jun",7:"Jul",8:"Agu",9:"Sep",10:"Okt",11:"Nov",12:"Des"}
            df['MonthName'] = df['Month'].map(month_map)
            agg = (
                df.groupby(['Year','Month','MonthName','Item Description'], dropna=False)[metric_col]
                .sum().reset_index())
            all_months_df = pd.DataFrame({'Month': list(range(1, 13)), 'MonthName': MONTH_ORDER})
            items = df['Item Description'].dropna().unique().tolist()
            complete = []
            for yr in years_to_show:
                for it in items:
                    t = all_months_df.copy()
                    t['Year'] = yr
                    t['Item Description'] = it
                    complete.append(t)
            grid = pd.concat(complete, ignore_index=True)
            merged = grid.merge(agg, on=['Year','Month','MonthName','Item Description'], how='left')
            merged[metric_col] = merged[metric_col].fillna(0)
            merged['MonthName'] = pd.Categorical(merged['MonthName'], categories=MONTH_ORDER, ordered=True)
            return merged

        monthly_df = prep_monthly(source_df)
        available_items = monthly_df['Item Description'].dropna().sort_values().unique().tolist() if not monthly_df.empty else []
        trend_mode = st.radio(
            "Select Trend Mode:",
            ["Item Trend", "Item Comparison"], horizontal=True, key="trend_mode_after_start")

        if trend_mode == "Item Trend":
            if not available_items:
                st.info("Tidak ada item untuk ditampilkan.")
            else:
                for it in available_items:
                    df_item = monthly_df[monthly_df["Item Description"] == it]
                    if df_item.empty: 
                        continue
                    st.subheader(f"📈 {it}")
                    if isinstance(year_option, str) and " vs " in year_option:
                        fig = px.line(df_item, x="MonthName",
                            y=metric_col, color="Year", markers=True, title=f"Monthly Trend – {it}")
                    else:
                        fig = px.line(df_item, x="MonthName", y=metric_col, markers=True, title=f"Monthly Trend – {it}")
                    fig.update_xaxes(title_text="Month")
                    fig.update_yaxes(title_text=y_label, tickformat=",.0f")
                    apply_month_order(fig)
                    st.plotly_chart(fig, use_container_width=True)
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
                df_sel = monthly_df[monthly_df["Item Description"].isin(selected)]
                if df_sel.empty:
                    st.info("Tidak ada data untuk item yang dipilih.")
                else:
                    if isinstance(year_option, str) and " vs " in year_option:
                        fig = px.line(df_sel, x="MonthName", y=metric_col, color="Item Description", 
                            line_dash="Year", markers=True, title=f"Monthly Trend Comparison – {item_option}")
                    else:
                        fig = px.line(df_sel, x="MonthName", y=metric_col, color="Item Description",
                            markers=True, title=f"Monthly Trend Comparison – {item_option}")
                    fig.update_xaxes(title_text="Month")
                    fig.update_yaxes(title_text=y_label, tickformat=",.0f")
                    apply_month_order(fig)
                    st.plotly_chart(fig, use_container_width=True)

        #region Tren Mingguan
        st.markdown("---")
        st.header("📅 Weekly Sales Trend")

        MONTH_NAMES_ID = ["Januari","Februari","Maret","April","Mei","Juni",
                          "Juli","Agustus","September","Oktober","November","Desember"]
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

        weekly_trend_mode = st.radio(
            "Select Weekly Trend Mode:",
            ["Item Trend", "Item Comparison"],
            horizontal=True,key="weekly_trend_mode_radio")

        def _month_week_starts(year: int, month: int, week_start: int = 0):
            first_day = datetime(year, month, 1)
            last_day = datetime(year, month, calendar.monthrange(year, month)[1])
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

            agg = (
                df.groupby(["Year","WeekLabel","Item Description"], dropna=False)[metric_col]
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
                df_m = prep_weekly_calendar(weekly_source_df, mnum, week_start=0)
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
                                fig_i = px.line(dfi, x="WeekName", y=metric_col, color="Year",
                                    line_dash="Year", facet_col="MonthName", facet_col_wrap=2, markers=True, title=None)
                            else:
                                fig_i = px.line(dfi, x="WeekName", y=metric_col, color="MonthName", markers=True, title=None)
                        else:
                            if isinstance(year_option, str) and " vs " in year_option:
                                fig_i = px.line(dfi, x="WeekName", y=metric_col, color="Year", line_dash="Year", markers=True, title=None)
                            else:
                                fig_i = px.line(dfi, x="WeekName", y=metric_col, markers=True, title=None)

                        fig_i.update_xaxes(title_text="Minggu ke-")
                        fig_i.update_yaxes(title_text=y_label, tickformat=",.0f")
                        fig_i.update_layout(yaxis=dict(tickformat=",.0f"))
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
                                    facet_col="MonthName" if len(selected_months) > 1 else None,
                                    facet_col_wrap=2 if len(selected_months) > 1 else None,
                                    markers=True, title="Weekly Trend Comparison")
                            else:
                                if len(selected_months) > 1:
                                    fig_c = px.line(df_sel_w, x="WeekName", y=metric_col, color="Item Description",
                                        line_dash="MonthName", markers=True, title="Weekly Trend Comparison")
                                else:
                                    fig_c = px.line(df_sel_w, x="WeekName", y=metric_col, color="Item Description",
                                        markers=True, title=f"Weekly Trend Comparison ({selected_months[0]})")

                            fig_c.update_xaxes(title_text="Minggu ke-")
                            fig_c.update_yaxes(title_text=y_label, tickformat=",.0f")
                            fig_c.update_layout(yaxis=dict(tickformat=",.0f"))
                            st.plotly_chart(fig_c, use_container_width=True)

        #region Top Customers
        st.subheader("📊 Top Customer")
        
        def top_customer_dual_full_nominal(df: pd.DataFrame, title_suffix: str = ""):
            # Cek apakah kolom Customer Name ada
            if 'Customer Name' not in df.columns:
                st.warning("Kolom 'Customer Name' tidak ditemukan di data ARIN.")
                return
            
            # Groupby HANYA berdasarkan Customer Name (bukan Customer Code)
            # Tapi hitung berapa banyak Customer Code per Customer Name
            sup = (
                df.groupby('Customer Name', dropna=False)
                .agg({
                    'Netto Quantity': 'sum',
                    'Total LC': 'sum',
                    'Customer Code': 'nunique'})  # Hitung jumlah unique code
                .reset_index()
                .rename(columns={'Customer Code': 'Code Count'}))
            
            # Ambil satu Customer Code untuk yang cuma punya 1 code
            code_map = (
                df.groupby('Customer Name')['Customer Code']
                .agg(lambda x: x.mode()[0] if len(x.mode()) > 0 else x.iloc[0])
                .to_dict())
            
            sup['Single Code'] = sup['Customer Name'].map(code_map)
            
            # Buat label X-axis: kalau Code Count > 1, pakai Customer Name, kalau = 1 pakai Customer Code
            sup['X_Label'] = sup.apply(
                lambda row: row['Customer Name'] if row['Code Count'] > 1 else str(row['Single Code']),
                axis=1)

            # URUTKAN BERDASARKAN TOTAL LC
            sup = sup.sort_values(by='Total LC', ascending=False).head(10)
            if sup.empty:
                st.info("Tidak ada data customer untuk periode ini.")
                return

            fig = go.Figure()
            # BAR OREN = TOTAL LC (di yaxis kiri)
            fig.add_trace(go.Bar(
                x=sup['X_Label'], y=sup['Total LC'], name='Total (IDR)', marker_color='orange',
                hovertemplate=(
                    '<b>Customer Name: %{customdata[0]}</b>'
                    '<br>Customer Codes: %{customdata[1]}'
                    '<br>Total LC: %{y:,}'
                    '<extra></extra>'),
                customdata=sup[['Customer Name', 'Code Count']],
                yaxis='y', offsetgroup='money'))

            # BAR BIRU = NETTO QUANTITY (di yaxis2 kanan)
            fig.add_trace(go.Bar(
                x=sup['X_Label'], y=sup['Netto Quantity'], name='Quantity (Netto)', marker_color='royalblue',
                hovertemplate=(
                    '<b>Customer Name: %{customdata[0]}</b>'
                    '<br>Customer Codes: %{customdata[1]}'
                    '<br>Quantity: %{y:,}'
                    '<extra></extra>'),
                customdata=sup[['Customer Name', 'Code Count']],
                yaxis='y2', offsetgroup='qty'))

            fig.update_layout(
                title=f"Top Customers {title_suffix}".strip(),
                xaxis_title="Customer Code / Name",
                yaxis=dict(title="Total LC", tickformat=",.0f", separatethousands=True, showgrid=True),
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
                top_customer_dual_full_nominal(d1, f"({y1})")
            with c2:
                top_customer_dual_full_nominal(d2, f"({y2})")
        else:
            y = int(year_option)
            d = combined_filtered[combined_filtered['Posting Date'].dt.year == y]
            top_customer_dual_full_nominal(d, f"({y})")

        #region Customer Activity Status
        if customerdata_file is not None:
            st.subheader("🧑‍🤝‍🧑 Customer Activity Status")
            df_master_customer = pd.read_excel(customerdata_file)

            def customer_activity_section(df_trans: pd.DataFrame, df_master: pd.DataFrame, year_option):
                if not pd.api.types.is_datetime64_any_dtype(df_trans['Posting Date']):
                    df_trans = df_trans.copy()
                    df_trans['Posting Date'] = pd.to_datetime(df_trans['Posting Date'], errors='coerce')
                if isinstance(year_option, str) and " vs " in year_option:
                    y1, y2 = map(int, year_option.split(" vs "))
                    df_year = df_trans[df_trans['Posting Date'].dt.year.isin([y1, y2])].copy()
                    title_suffix = f"({y1} & {y2})"
                else:
                    y = int(year_option)
                    df_year = df_trans[df_trans['Posting Date'].dt.year == y].copy()
                    title_suffix = f"({y})"

                # PERBAIKAN: Pakai Customer Name untuk menghindari duplikasi
                active_names_raw = (
                    df_year['Customer Name']
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .str.upper()  # Normalisasi huruf besar
                    .unique())
                active_names_set = set(active_names_raw)
                
                df_master_local = df_master.copy()
                df_master_local['Customer Name'] = (
                    df_master_local['Customer Name']
                    .astype(str)
                    .str.strip()
                    .str.upper())  # Normalisasi huruf besar

                master_names_set = set(df_master_local['Customer Name'].unique())
                active_customers = sorted(master_names_set.intersection(active_names_set))
                inactive_customers = sorted(master_names_set.difference(active_names_set))

                n_active = len(active_customers)
                n_inactive = len(inactive_customers)
                n_total_master = len(master_names_set)

                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric(
                        label=f"Active Customers {title_suffix}",
                        value=n_active)
                with c2:
                    st.metric(
                        label=f"Inactive Customers {title_suffix}",
                        value=n_inactive)
                with c3:
                    st.metric(
                        label="Total Registered Customers",
                        value=n_total_master)

            customer_activity_section(df_trans=combined_filtered, df_master=df_master_customer, year_option=year_option)

            #region Top City
            st.subheader("🌆 Top Cities by Sales")
            def top_city_chart(df_trans: pd.DataFrame, df_master: pd.DataFrame, title_suffix: str = ""):
                df_work = df_trans.copy()
                df_work['Customer Code'] = df_work['Customer Code'].astype(str).str.strip()
                df_master_local = df_master.copy()
                df_master_local['Customer Code'] = df_master_local['Customer Code'].astype(str).str.strip()
                df_merged = df_work.merge(
                    df_master_local[['Customer Code', 'Ship-to City', 'Bill-to City']], 
                    on='Customer Code', how='left')
                
                def get_city(row):
                    ship_city = str(row.get('Ship-to City', '')).strip()
                    bill_city = str(row.get('Bill-to City', '')).strip()
                    
                    if ship_city and ship_city not in ['nan', 'None', '']:
                        return ship_city.lower()  # Konversi ke lowercase
                    elif bill_city and bill_city not in ['nan', 'None', '']:
                        return bill_city.lower()  # Konversi ke lowercase
                    else:
                        return "unknown"  # lowercase juga untuk konsistensi
                
                df_merged['City'] = df_merged.apply(get_city, axis=1)
                city_sales = (
                    df_merged.groupby('City', dropna=False)['Netto Quantity']
                    .sum()
                    .sort_values(ascending=True) 
                    .tail(10)
                    .reset_index())
                
                if city_sales.empty:
                    st.info("Tidak ada data kota untuk periode ini.")
                    return
                
                # Capitalize untuk display yang lebih rapi (opsional)
                city_sales['City_Display'] = city_sales['City'].str.title()
                
                n_cities = len(city_sales)
                colors = px.colors.sequential.Purp
                city_sales['color'] = [colors[int(i * (len(colors)-1) / max(1, n_cities-1))] for i in range(n_cities)]
                city_sales['DisplayQty'] = city_sales['Netto Quantity'].apply(lambda x: f"{x:,.0f}")
                
                # Gunakan City_Display untuk tampilan
                fig = px.bar(city_sales, x='Netto Quantity', y='City_Display', orientation='h', text='DisplayQty',
                    color='color', color_discrete_map='identity', title=f'Cities by Quantity {title_suffix}',
                    labels={'Netto Quantity': 'Total Quantity (Netto)', 'City_Display': 'Kota'})
                
                fig.update_layout(
                    yaxis=dict(title="Kota"),
                    xaxis=dict(title="Total Quantity (Netto)", tickformat=",.0f"), showlegend=False, height=500)
                
                fig.update_traces(textposition='outside')
                st.plotly_chart(fig, use_container_width=True)

            if isinstance(year_option, str) and " vs " in year_option:
                y1, y2 = map(int, year_option.split(" vs "))
                d1 = combined_filtered[combined_filtered['Posting Date'].dt.year == y1]
                d2 = combined_filtered[combined_filtered['Posting Date'].dt.year == y2]
                c1, c2 = st.columns(2)
                with c1:
                    top_city_chart(d1, df_master_customer, f"({y1})")
                with c2:
                    top_city_chart(d2, df_master_customer, f"({y2})")
            else:
                y = int(year_option)
                d = combined_filtered[combined_filtered['Posting Date'].dt.year == y]
                top_city_chart(d, df_master_customer, f"({y})")

        #region Rata-rata Uang per Quantity per Bulan (Fokus hanya pada IDR)
        st.markdown("---")
        st.header("📅 Monthly Average Sales Value by Quantity")

        # Cek apakah DO sudah diproses
        if 'do_data1' not in st.session_state:
            st.info("DO data belum diproses. Silakan upload dan proses file terlebih dahulu.")
        else:
            # Gabungkan DO data
            do1 = st.session_state.get("do_data1", pd.DataFrame()).copy()
            do2 = st.session_state.get("do_data2", pd.DataFrame()).copy() if 'do_data2' in st.session_state else pd.DataFrame()

            if not do2.empty:
                do_df = pd.concat([do1, do2], ignore_index=True)
            else:
                do_df = do1.copy()

            # Tambahkan kategori ke DO jika belum ada
            if 'Grub' not in do_df.columns and 'Item Description' in do_df.columns:
                do_df['Grub'] = do_df['Item Description'].apply(categorize_item)

            # Ambil invoice data
            invoice_df = combined_filtered.copy()

            if invoice_df.empty:
                st.info("Tidak ada data Invoice untuk ditampilkan.")
            elif do_df.empty:
                st.info("Tidak ada data DO untuk ditampilkan.")
            else:
                # Pastikan kolom tanggal sudah datetime
                if not pd.api.types.is_datetime64_any_dtype(invoice_df["Posting Date"]):
                    invoice_df["Posting Date"] = pd.to_datetime(invoice_df["Posting Date"], errors="coerce")

                if "Posting Date" in do_df.columns:
                    if not pd.api.types.is_datetime64_any_dtype(do_df["Posting Date"]):
                        do_df["Posting Date"] = pd.to_datetime(do_df["Posting Date"], errors="coerce")

                # Drop NaT
                invoice_df = invoice_df.dropna(subset=["Posting Date"]).copy()
                if "Posting Date" in do_df.columns:
                    do_df = do_df.dropna(subset=["Posting Date"]).copy()

                # Filter berdasarkan item yang dipilih
                invoice_df = invoice_df[invoice_df["Grub"] == item_option].copy()
                if "Grub" in do_df.columns:
                    do_df = do_df[do_df["Grub"] == item_option].copy()

                # Nama bulan
                MONTH_NAMES_ID = [
                    "January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"
                ]

                invoice_df["Month"] = invoice_df["Posting Date"].dt.month.apply(lambda x: MONTH_NAMES_ID[x - 1])
                if "Posting Date" in do_df.columns:
                    do_df["Month"] = do_df["Posting Date"].dt.month.apply(lambda x: MONTH_NAMES_ID[x - 1])

                # Fungsi untuk menghitung rata-rata bulanan (IDR Only)
                def calc_monthly_avg(inv_df, do_df, year):
                    """Menghitung rata-rata pembelian per bulan untuk tahun tertentu"""
                    # Filter per tahun
                    inv_year = inv_df[inv_df["Posting Date"].dt.year == year].copy()
                    do_year = do_df[do_df["Posting Date"].dt.year == year].copy() if "Posting Date" in do_df.columns else do_df.copy()

                    # Discount unique
                    discount_unique = inv_year.drop_duplicates(subset=["Doc Number"], keep="first")[
                        ["Doc Number", "Month", "Discount"]
                    ].copy()

                    # Isi NaN dengan 0
                    discount_unique["Discount"] = discount_unique["Discount"].fillna(0) if "Discount" in discount_unique.columns else 0

                    # Aggregate Discount per bulan
                    discount_monthly = discount_unique.groupby("Month").agg({
                        "Discount": "sum",
                    }).reset_index()

                    # Aggregate kolom lainnya dari ARIN
                    inv_monthly = inv_year.groupby("Month").agg({
                        "Netto Quantity": "sum",
                        "Total LC": "sum",
                    }).reset_index()

                    # Merge dengan discount
                    inv_monthly = pd.merge(inv_monthly, discount_monthly, on="Month", how="left").fillna(0)

                    # Aggregate DO per bulan
                    if not do_year.empty and "Month" in do_year.columns:
                        do_monthly = do_year.groupby("Month").agg({
                            "Freight": "sum",
                        }).reset_index()
                        monthly = pd.merge(inv_monthly, do_monthly, on="Month", how="left").fillna(0)
                    else:
                        monthly = inv_monthly.copy()
                        monthly["Freight"] = 0

                    # HITUNG RATA-RATA IDR
                    monthly["IDR"] = monthly.apply(
                        lambda row: (
                            (row["Total LC"] + row["Freight"] - row["Discount"]) / row["Netto Quantity"]
                            if (row["Total LC"] > 0 or row["Freight"] > 0) and row["Netto Quantity"] > 0
                            else 0
                        ),
                        axis=1
                    )

                    # Replace inf/nan dengan 0
                    monthly["IDR"] = monthly["IDR"].replace([float('inf'), -float('inf')], 0).fillna(0)

                    return monthly[["Month", "Netto Quantity", "Total LC", "Freight", "Discount", "IDR"]]

                # Fungsi untuk menghitung rata-rata tahunan (IDR Only)
                def calc_yearly_avg(inv_df, do_df, year):
                    """Menghitung rata-rata pembelian tahunan"""
                    inv_year = inv_df[inv_df["Posting Date"].dt.year == year].copy()
                    do_year = do_df[do_df["Posting Date"].dt.year == year].copy() if "Posting Date" in do_df.columns else do_df.copy()

                    # Total Quantity dan LC
                    total_qty = inv_year["Netto Quantity"].sum()
                    total_doc_idr = inv_year["Total LC"].sum()

                    # Discount UNIK
                    discount_unique = inv_year.drop_duplicates(subset=["Doc Number"], keep="first")
                    total_discount_idr = discount_unique["Discount"].fillna(0).sum() if "Discount" in discount_unique.columns else 0

                    # Total DO
                    total_freight_idr = do_year["Freight"].sum() if "Freight" in do_year.columns and not do_year.empty else 0

                    # HITUNG AVERAGE IDR
                    if (total_doc_idr > 0 or total_freight_idr > 0) and total_qty > 0:
                        avg_idr = (total_doc_idr + total_freight_idr - total_discount_idr) / total_qty
                    else:
                        avg_idr = 0

                    return {
                        "Netto Quantity": total_qty,
                        "Total LC": total_doc_idr,
                        "Freight": total_freight_idr,
                        "Discount": total_discount_idr,
                        "IDR": avg_idr
                    }

                # MODE PERBANDINGAN TAHUN (VS)
                if isinstance(year_option, str) and " vs " in year_option:
                    y1, y2 = map(int, year_option.split(" vs "))

                    # Filter data per tahun
                    invoice_df_filtered = invoice_df[invoice_df["Posting Date"].dt.year.isin([y1, y2])]
                    do_df_filtered = do_df[do_df["Posting Date"].dt.year.isin([y1, y2])] if "Posting Date" in do_df.columns else do_df

                    # Hitung untuk kedua tahun
                    detail_y1 = calc_monthly_avg(invoice_df_filtered, do_df_filtered, y1)
                    detail_y2 = calc_monthly_avg(invoice_df_filtered, do_df_filtered, y2)

                    # TABEL DETAIL (SUM) - Gabungkan kedua tahun
                    detail_y1_display = detail_y1[["Month", "Netto Quantity", "Total LC", "Freight", "Discount"]].rename(columns={
                        "Netto Quantity": f"Sum Netto Quantity ({y1})",
                        "Total LC": f"Sum Total LC ({y1})",
                        "Freight": f"Sum Freight ({y1})",
                        "Discount": f"Discount ({y1})"
                    })
                    
                    detail_y2_display = detail_y2[["Month", "Netto Quantity", "Total LC", "Freight", "Discount"]].rename(columns={
                        "Netto Quantity": f"Sum Netto Quantity ({y2})",
                        "Total LC": f"Sum Total LC ({y2})",
                        "Freight": f"Sum Freight ({y2})",
                        "Discount": f"Discount ({y2})"
                    })

                    # Merge kedua tahun
                    detail_compare = pd.merge(detail_y1_display, detail_y2_display, on="Month", how="outer")

                    # Urutkan berdasarkan bulan
                    detail_compare["Month"] = pd.Categorical(detail_compare["Month"], categories=MONTH_NAMES_ID, ordered=True)
                    detail_compare = detail_compare.sort_values("Month")

                    # YEARLY TOTAL untuk detail
                    yearly_y1 = calc_yearly_avg(invoice_df_filtered, do_df_filtered, y1)
                    yearly_y2 = calc_yearly_avg(invoice_df_filtered, do_df_filtered, y2)

                    total_row_detail = pd.DataFrame([{
                        "Month": "Yearly Total",
                        f"Sum Netto Quantity ({y1})": yearly_y1["Netto Quantity"],
                        f"Sum Total LC ({y1})": yearly_y1["Total LC"],
                        f"Sum Freight ({y1})": yearly_y1["Freight"],
                        f"Discount ({y1})": yearly_y1["Discount"],
                        f"Sum Netto Quantity ({y2})": yearly_y2["Netto Quantity"],
                        f"Sum Total LC ({y2})": yearly_y2["Total LC"],
                        f"Sum Freight ({y2})": yearly_y2["Freight"],
                        f"Discount ({y2})": yearly_y2["Discount"]
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

                    # TABEL IDR AVERAGE
                    avg_y1 = detail_y1[["Month", "IDR"]].rename(columns={"IDR": f"IDR {y1}"})
                    avg_y2 = detail_y2[["Month", "IDR"]].rename(columns={"IDR": f"IDR {y2}"})

                    monthly_compare = pd.merge(avg_y1, avg_y2, on="Month", how="outer")

                    # Urutkan berdasarkan bulan
                    monthly_compare["Month"] = pd.Categorical(monthly_compare["Month"], categories=MONTH_NAMES_ID, ordered=True)
                    monthly_compare = monthly_compare.sort_values("Month")

                    total_row_avg = pd.DataFrame([{
                        "Month": "Yearly Average",
                        f"IDR {y1}": yearly_y1["IDR"],
                        f"IDR {y2}": yearly_y2["IDR"]
                    }])

                    monthly_compare = pd.concat([monthly_compare, total_row_avg], ignore_index=True).fillna("-")

                    # Format angka dengan 2 desimal
                    for col in [f"IDR {y1}", f"IDR {y2}"]:
                        monthly_compare[col] = monthly_compare[col].apply(
                            lambda x: f"Rp{x:,.2f}" if x != "-" and pd.notna(x) and x > 0 else "-"
                        )

                    st.subheader("💰 Average IDR by Month")
                    st.dataframe(monthly_compare, use_container_width=True, hide_index=True)

                # MODE SATU TAHUN
                else:
                    y = int(year_option)

                    # Filter per tahun
                    invoice_year = invoice_df[invoice_df["Posting Date"].dt.year == y].copy()
                    do_year = do_df[do_df["Posting Date"].dt.year == y].copy() if "Posting Date" in do_df.columns else do_df.copy()

                    if invoice_year.empty:
                        st.info("Tidak ada data Invoice untuk item dan tahun yang dipilih.")
                    else:
                        # Hitung monthly average (dengan detail)
                        monthly_detail = calc_monthly_avg(invoice_df, do_df, y)

                        # Urutkan
                        monthly_detail["Month"] = pd.Categorical(monthly_detail["Month"], categories=MONTH_NAMES_ID, ordered=True)
                        monthly_detail = monthly_detail.sort_values("Month")

                        # YEARLY AVERAGE
                        yearly_avg = calc_yearly_avg(invoice_df, do_df, y)

                        # TABEL DETAIL (SUM)
                        detail_table = monthly_detail[["Month", "Netto Quantity", "Total LC", "Freight", "Discount"]].copy()
                        detail_table = detail_table.rename(columns={
                            "Netto Quantity": "Sum Netto Quantity",
                            "Total LC": "Sum Total LC",
                            "Freight": "Sum Freight"
                        })

                        # Tambah row total
                        total_row_detail = pd.DataFrame([{
                            "Month": "Yearly Total",
                            "Sum Netto Quantity": yearly_avg["Netto Quantity"],
                            "Sum Total LC": yearly_avg["Total LC"],
                            "Sum Freight": yearly_avg["Freight"],
                            "Discount": yearly_avg["Discount"]
                        }])

                        detail_table = pd.concat([detail_table, total_row_detail], ignore_index=True).fillna(0)

                        # Format angka
                        for col in ["Sum Netto Quantity", "Sum Total LC", "Sum Freight", "Discount"]:
                            detail_table[col] = detail_table[col].apply(
                                lambda x: f"{x:,.2f}" if pd.notna(x) and x != 0 else "0.00"
                            )

                        st.subheader("📊 Detail Sum per Month")
                        st.dataframe(detail_table, use_container_width=True, hide_index=True)

                        # TABEL IDR AVERAGE
                        tampil_tabel = monthly_detail[["Month", "IDR"]].copy()

                        total_row = pd.DataFrame([{
                            "Month": "Yearly Average",
                            "IDR": yearly_avg["IDR"]
                        }])

                        tampil_tabel = pd.concat([tampil_tabel, total_row], ignore_index=True).fillna("-")

                        # Format angka dengan 2 desimal
                        tampil_tabel["IDR"] = tampil_tabel["IDR"].apply(
                            lambda x: f"Rp{x:,.2f}" if x != "-" and pd.notna(x) and x > 0 else "-"
                        )

                        st.subheader("💰 Average IDR per Month")
                        st.dataframe(tampil_tabel, use_container_width=True, hide_index=True)