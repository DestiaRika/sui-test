
import streamlit as st


st.set_page_config(page_title="Data Visualization Web", layout="centered")
st.title("Data Visualization Web")
st.write("This is a simple web application for data visualization using Streamlit.")
st.divider()

with st.container():
    if(st.button("Click Me")):
        st.switch_page("pages/gabung-beli.py")

# import io
# import re
# import pandas as pd
# import streamlit as st

# # =====================
# # Config & Header
# # =====================
# st.set_page_config(page_title="Gabung Data Jual & Beli", layout="wide")
# st.title("🔗 Gabung Data Jual & Beli")
# st.caption("Upload file mentah kamu satu-satu. Nama file bebas (nggak perlu tahun).")

# # =====================
# # Helpers
# # =====================
# def _read_xlsx(uploaded_file: st.runtime.uploaded_file_manager.UploadedFile, sheet_name=0):
#     if uploaded_file is None:
#         return None
#     try:
#         return pd.read_excel(uploaded_file, sheet_name=sheet_name)
#     except Exception as e:
#         st.error(f"Gagal baca Excel {getattr(uploaded_file, 'name', '')}: {e}")
#         return None

# def _download_df(df: pd.DataFrame, filename: str, label: str):
#     buf = io.BytesIO()
#     with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
#         df.to_excel(writer, index=False)
#     st.download_button(label=label, data=buf.getvalue(),
#         file_name=filename,
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# def _norm_key(s: pd.Series) -> pd.Series:
#     return (s.astype(str).str.strip()
#             .str.replace(r"\.0+$", "", regex=True)
#             .str.replace(" ", ""))

# # ======================
# # Column ordering helpers (URUTAN TAMPILAN)
# # ======================
# DOC_PRIORITY = [
#     "Doc Number","Base On","Remarks","Customer Code","Vendor Code","Item No","Product Code",
#     "Bag Quantity","Quantity","Unit Price","Whse","Warehouse","Instruction Number",
#     "Base Number","License Number","License No.","Weight Difference","Details","Date","Posting Date"
# ]

# def _order_group(df: pd.DataFrame, suffix: str) -> list:
#     cols = [c for c in df.columns if c.endswith(suffix)]
#     def key(c):
#         base = c[: -len(suffix)]
#         try:
#             idx = DOC_PRIORITY.index(base)
#         except ValueError:
#             idx = len(DOC_PRIORITY)
#         return (idx, base.lower())
#     return sorted(cols, key=key)

# def reorder_columns(df: pd.DataFrame, group_suffixes: list[str]) -> pd.DataFrame:
#     ordered = []
#     for suf in group_suffixes:
#         ordered += _order_group(df, suf)
#     tail = [c for c in df.columns if c not in ordered]
#     return df[ordered + tail]

# # =====================
# # BELI PIPELINE
# # =====================
# def _clean_val(v):
#     return re.sub(r"\.0+$", "", str(v).strip().upper())

# def _extract_grpo_from_remarks(txt: str):
#     if pd.isna(txt): return []
#     m = re.findall(r"Based On Goods Receipt(?: PO)?\s*([\d\. ]+)", str(txt))
#     if not m: return []
#     return re.findall(r"\d{8}", m[0])

# def pipeline_beli(df_ap: pd.DataFrame, df_grpo: pd.DataFrame, df_timbangan: pd.DataFrame|None):
#     df_ap = df_ap.copy(); df_grpo = df_grpo.copy()
#     df_ap.columns = df_ap.columns.astype(str).str.strip()
#     df_grpo.columns = df_grpo.columns.astype(str).str.strip()

#     imp = ['Doc Number','Vendor Code','Item No','Quantity','Unit Price','Whse']
#     for c in imp:
#         if c not in df_grpo.columns: df_grpo[c] = ''
#         df_grpo[c] = df_grpo[c].astype(str).str.strip().str.replace(r"\.0+$","",regex=True).str.upper()

#     if 'Remarks' not in df_ap.columns: df_ap['Remarks'] = ''
#     df_ap['GRPO_List'] = df_ap['Remarks'].apply(_extract_grpo_from_remarks)

#     used = set()
#     def choose_grpo(row):
#         grpos = row['GRPO_List']
#         if not grpos: return None
#         best,score = None,-1
#         apv = {c:_clean_val(row.get(c,'')) for c in imp if c in df_ap.columns}
#         for gnum in grpos:
#             if gnum in used: continue
#             m = df_grpo[df_grpo['Doc Number']==gnum]
#             if m.empty: continue
#             g = m.iloc[0]
#             gv = {c:_clean_val(g.get(c,'')) for c in imp}
#             same = sum(apv.get(c,'')==gv.get(c,'') for c in imp if c!='Doc Number')
#             if same>score:
#                 score, best = same, gnum
#         if best:
#             used.add(best)
#             return best
#         return None

#     df_ap['Base On GRPO'] = df_ap.apply(choose_grpo, axis=1)

#     # fallback by 5 key cols
#     if 'Base On GRPO' in df_ap.columns:
#         for i, r in df_ap[df_ap['Base On GRPO'].isna()].iterrows():
#             apv = {c:_clean_val(r.get(c,'')) for c in imp if c in df_ap.columns}
#             poss = df_grpo.copy()
#             for c in ['Vendor Code','Item No','Quantity','Unit Price','Whse']:
#                 poss = poss[poss[c].apply(_clean_val)==apv.get(c,'')]
#                 if poss.empty: break
#             if not poss.empty:
#                 poss = poss[~poss['Doc Number'].isin(used)]
#                 if not poss.empty:
#                     dn = poss.iloc[0]['Doc Number']
#                     df_ap.at[i,'Base On GRPO'] = dn; used.add(dn)

#     out_beli = df_ap

#     # optional: Timbangan vs GRPO
#     if isinstance(df_timbangan,pd.DataFrame) and not df_timbangan.empty:
#         t = df_timbangan.copy(); t.columns = t.columns.astype(str).str.strip()
#         for c in ['Vendor Code','Product Code','Warehouse','License No.','Weight Difference']:
#             if c not in t.columns: t[c]=''
#         t['merged_key'] = t[['Vendor Code','Product Code','Warehouse','License No.','Weight Difference']].astype(str).agg('_'.join,axis=1)

#         g = df_grpo.copy()
#         for c in ['Vendor Code','Item No','Whse','License Number','Quantity']:
#             if c not in g.columns: g[c]=''
#         g['merged_key'] = g[['Vendor Code','Item No','Whse','License Number','Quantity']].astype(str).agg('_'.join,axis=1)

#         t = t.add_suffix('_TIMB').rename(columns={'merged_key_TIMB':'merged_key'})
#         g = g.add_suffix('_GRPO').rename(columns={'merged_key_GRPO':'merged_key'})

#         tg = pd.merge(g, t, on='merged_key', how='left')
#         out_beli = pd.merge(
#             tg,
#             df_ap.add_suffix('_APIN').rename(columns={'Base On GRPO_APIN':'Doc Number_GRPO'}),
#             on='Doc Number_GRPO', how='left'
#         )

#     return out_beli

# # =====================
# # JUAL PIPELINE
# # =====================
# def _extract_do_list(txt: str):
#     if pd.isna(txt): return []
#     m = re.findall(r"Deliveries[^\d]*(.*)", str(txt))
#     if not m: return []
#     return re.findall(r"\d{8}", m[0])

# def pipeline_jual(
#     df_ar: pd.DataFrame, df_do: pd.DataFrame,
#     df_return: pd.DataFrame|None, df_timbangan_jual: pd.DataFrame|None,
#     df_ardp: pd.DataFrame|None, df_arcm: pd.DataFrame|None,
#     df_so: pd.DataFrame|None, df_cmd: pd.DataFrame|None
# ):
#     if df_ar is None or df_do is None:
#         raise ValueError("AR INVOICE dan DO wajib diunggah.")

#     df_ar = df_ar.copy(); df_do = df_do.copy()
#     for d in (df_ar, df_do): d.columns = d.columns.astype(str).str.strip()

#     need_ar = ['Remarks','Doc Number','Customer Code','Item No','Bag Quantity','Quantity','Unit Price','Whse']
#     need_do = ['Doc Number','Customer Code','Item No','Bag Quantity','Quantity','Unit Price','Whse','Instruction Number','Remarks']
#     for c in need_ar:
#         if c not in df_ar.columns: df_ar[c]=''
#     for c in need_do:
#         if c not in df_do.columns: df_do[c]=''

#     df_ar['DO_List'] = df_ar['Remarks'].apply(_extract_do_list)

#     imp = ['Doc Number','Customer Code','Item No','Bag Quantity','Quantity','Unit Price','Whse']
#     for c in imp:
#         df_do[c] = df_do[c].astype(str).str.strip().str.replace(r"\.0+$","",regex=True)

#     used = set()
#     def pick_do(row):
#         lst = row['DO_List']
#         if not lst: return None
#         if len(lst)==1:
#             used.add(lst[0]); return lst[0]
#         for num in lst:
#             if num in used: continue
#             m = df_do[df_do['Doc Number']==num]
#             if m.empty: continue
#             g = m.iloc[0]; ok=True
#             for c in imp:
#                 if c=='Doc Number': continue
#                 va = re.sub(r"\.0+$","",str(row.get(c,'')).strip())
#                 vd = re.sub(r"\.0+$","",str(g.get(c,'')).strip())
#                 if va!=vd: ok=False; break
#             if ok:
#                 used.add(num); return num
#         return None

#     df_ar['Base On DO'] = df_ar.apply(pick_do, axis=1)

#     # DO ↔ AR
#     do_s = df_do.add_suffix('_DO')
#     ar_s = df_ar.add_suffix('_ARIN')
#     if 'Doc Number_DO' not in do_s.columns: do_s['Doc Number_DO']=''
#     if 'Base On DO_ARIN' not in ar_s.columns: ar_s['Base On DO_ARIN']=''
#     merged = pd.merge(do_s, ar_s, left_on='Doc Number_DO', right_on='Base On DO_ARIN', how='left')

#     # RETURN (opsional)
#     if isinstance(df_return,pd.DataFrame) and not df_return.empty:
#         r = df_return.copy(); r.columns = r.columns.astype(str).str.strip()
#         if 'Base on DO' not in r.columns:
#             r['Base on DO'] = r['Remarks'].str.extract(r'Based On Deliveries (\d{8})') if 'Remarks' in r.columns else ''
#         for c in ['Customer Code','Quantity','Base on DO']:
#             if c not in r.columns: r[c]=''
#         r['Customer Code'] = _norm_key(r['Customer Code'])
#         r['Quantity'] = _norm_key(r['Quantity'])
#         r['Base on DO'] = _norm_key(r['Base on DO'])
#         merged['Customer Code_DO'] = _norm_key(merged['Customer Code_DO'])
#         merged['Quantity_DO'] = _norm_key(merged['Quantity_DO'])
#         merged['Doc Number_DO'] = _norm_key(merged['Doc Number_DO'])
#         r['key_return'] = r[['Customer Code','Quantity','Base on DO']].astype(str).agg('_'.join,axis=1)
#         merged['key_return'] = merged[['Customer Code_DO','Quantity_DO','Doc Number_DO']].astype(str).agg('_'.join,axis=1)
#         r = r.rename(columns={c:f"{c}_RETURN" for c in r.columns if c!='key_return'})
#         merged = pd.merge(merged, r, on='key_return', how='left')

#     # TIMBANGAN JUAL (opsional) Instruction Number_DO ↔ Base Number_DL
#     if isinstance(df_timbangan_jual,pd.DataFrame) and not df_timbangan_jual.empty:
#         dl = df_timbangan_jual.copy(); dl.columns = dl.columns.astype(str).str.strip()
#         if 'Base Number' in dl.columns:
#             dl = dl.drop_duplicates(subset=['Base Number'], keep='first')
#         dl = dl.rename(columns={c:f"{c}_DL" for c in dl.columns})
#         if 'Instruction Number_DO' not in merged.columns: merged['Instruction Number_DO']=''
#         if 'Base Number_DL' not in dl.columns: dl['Base Number_DL']=''
#         merged['Instruction Number_DO'] = _norm_key(merged['Instruction Number_DO'])
#         dl['Base Number_DL'] = _norm_key(dl['Base Number_DL'])
#         merged = pd.merge(merged, dl, left_on='Instruction Number_DO', right_on='Base Number_DL', how='left')

#     # ARDP / ARCM / SO (opsional, side-join)
#     a=c=s=None
#     if isinstance(df_ardp,pd.DataFrame) and not df_ardp.empty:
#         a = df_ardp.copy(); a.columns = a.columns.astype(str).str.strip()
#         if 'Base On SO' not in a.columns and 'Details' in a.columns:
#             a['Base On SO'] = a['Details'].str.extract(r'Sales Orders (\d{8})')
#         a = a.add_suffix('_ARDP')
#         if 'Doc Number_ARDP' in a.columns: a['Doc Number_ARDP']=_norm_key(a['Doc Number_ARDP'])
#         if 'Base On SO_ARDP' in a.columns: a['Base On SO_ARDP']=_norm_key(a['Base On SO_ARDP'])

#     if isinstance(df_arcm,pd.DataFrame) and not df_arcm.empty:
#         c = df_arcm.copy(); c.columns = c.columns.astype(str).str.strip()
#         if 'Base On ARDP' not in c.columns and 'Remarks' in c.columns:
#             c['Base On ARDP'] = c['Remarks'].str.extract(r'A/R Down Payment (\d{8})')
#         c = c.add_suffix('_CM')
#         if 'Base On ARDP_CM' in c.columns: c['Base On ARDP_CM']=_norm_key(c['Base On ARDP_CM'])

#     if isinstance(df_so,pd.DataFrame) and not df_so.empty:
#         s = df_so.copy(); s.columns = s.columns.astype(str).str.strip()
#         s = s.add_suffix('_SO')
#         if 'Doc Number_SO' in s.columns: s['Doc Number_SO']=_norm_key(s['Doc Number_SO'])

#     if 'Remarks_DO' in merged.columns and 'Base On_SO' not in merged.columns:
#         merged['Base On_SO'] = merged['Remarks_DO'].str.extract(r'Sales Orders (\d{8})')
#     if 'Base On_SO' in merged.columns: merged['Base On_SO']=_norm_key(merged['Base On_SO'])

#     if a is not None:
#         join_ac = a if c is None else pd.merge(c, a, left_on='Base On ARDP_CM', right_on='Doc Number_ARDP', how='outer')
#         join_acs = join_ac if s is None else pd.merge(join_ac, s, left_on='Base On SO_ARDP', right_on='Doc Number_SO', how='outer')
#         merged = pd.merge(merged, join_acs, left_on='Base On_SO', right_on='Doc Number_SO', how='left')

#     # CUSTOMER MASTER (opsional)
#     if isinstance(df_cmd,pd.DataFrame) and not df_cmd.empty:
#         cmd = df_cmd.copy(); cmd.columns = cmd.columns.astype(str).str.strip()
#         cmd = cmd.add_suffix('_CMD')
#         left_key = next((c for c in merged.columns if c.endswith('Customer Code_ARIN')), None)
#         if left_key is None:
#             left_key = next((c for c in merged.columns if c.endswith('Customer Code_DO')), None)
#         right_key = next((c for c in cmd.columns if c.endswith('Customer Code_CMD')), None)
#         if left_key and right_key:
#             merged[left_key] = _norm_key(merged[left_key])
#             cmd[right_key] = _norm_key(cmd[right_key])
#             merged = pd.merge(merged, cmd, left_on=left_key, right_on=right_key, how='left')

#     return merged

# # =====================
# # UI
# # =====================
# mode = st.sidebar.radio("Pilih Mode", ["Gabung Data Jual", "Gabung Data Beli"], index=0)

# if mode == "Gabung Data Beli":
#     st.header("📥 Gabung Data **Beli**")
#     c1, c2, c3 = st.columns(3)
#     with c1:
#         up_ap   = st.file_uploader("AP INVOICE.xlsx",        type=["xlsx"], key="ap")
#         up_grpo = st.file_uploader("GRPO.xlsx",              type=["xlsx"], key="grpo")
#     with c2:
#         up_timb = st.file_uploader("TIMBANGAN BELI.xlsx",    type=["xlsx"], key="timb_beli")
#     with c3:
#         st.info("Minimal: AP INVOICE + GRPO. Timbangan opsional.")

#     if st.button("🔄 Proses Beli"):
#         if not up_ap or not up_grpo:
#             st.warning("Mohon unggah **AP INVOICE** dan **GRPO**.")
#         else:
#             df_ap   = _read_xlsx(up_ap)
#             df_grpo = _read_xlsx(up_grpo)
#             df_timb = _read_xlsx(up_timb) if up_timb else None

#             with st.spinner("Memproses data beli…"):
#                 out_beli = pipeline_beli(df_ap, df_grpo, df_timb)

#             # URUTAN TAMPILAN: PO → APDP → TIMBANGAN → GRPO → APIN
#             BELI_ORDER = ['_PO','_ARDP','_TIMB','_GRPO','_APIN']
#             out_beli = reorder_columns(out_beli, BELI_ORDER)

#             st.success("Selesai!")
#             st.dataframe(out_beli.head(100))
#             _download_df(out_beli, "DATA_PEMBELIAN.xlsx", "⬇️ Unduh DATA_PEMBELIAN.xlsx")

# else:
#     st.header("📤 Gabung Data **Jual**")
#     st.markdown("Unggah file kamu. Yang wajib minimal **AR INVOICE** dan **DO**. Sisanya opsional.")

#     c1, c2, c3 = st.columns(3)
#     with c1:
#         up_ar     = st.file_uploader("AR INVOICE.xlsx (wajib)", type=["xlsx"], key="arin")
#         up_do     = st.file_uploader("DO.xlsx (wajib)",          type=["xlsx"], key="do")
#         up_return = st.file_uploader("RETURN.xlsx",              type=["xlsx"], key="ret")
#     with c2:
#         up_timbjl = st.file_uploader("TIMBANGAN JUAL.xlsx",      type=["xlsx"], key="timb_jual")
#         up_ardp   = st.file_uploader("ARDP.xlsx",                type=["xlsx"], key="ardp")
#         up_arcm   = st.file_uploader("ARCM.xlsx",                type=["xlsx"], key="arcm")
#     with c3:
#         up_so     = st.file_uploader("SO.xlsx",                  type=["xlsx"], key="so")
#         up_cmd    = st.file_uploader("CUSTOMER MASTER DATA.xlsx",type=["xlsx"], key="cmd")

#     if st.button("🔄 Proses Jual"):
#         if not up_ar or not up_do:
#             st.warning("Minimal unggah **AR INVOICE** dan **DO** dulu.")
#         else:
#             df_ar   = _read_xlsx(up_ar)
#             df_do   = _read_xlsx(up_do)
#             df_ret  = _read_xlsx(up_return) if up_return else None
#             df_tjl  = _read_xlsx(up_timbjl) if up_timbjl else None
#             df_ardp = _read_xlsx(up_ardp) if up_ardp else None
#             df_cm   = _read_xlsx(up_arcm) if up_arcm else None
#             df_so   = _read_xlsx(up_so) if up_so else None
#             df_cmd  = _read_xlsx(up_cmd) if up_cmd else None

#             with st.spinner("Memproses data jual…"):
#                 out_jual = pipeline_jual(df_ar, df_do, df_ret, df_tjl, df_ardp, df_cm, df_so, df_cmd)

#             # URUTAN TAMPILAN: SO → ARDP → RCM → TIMBANGAN → DO → RETURN → ARIN
#             JUAL_ORDER = ['_SO','_ARDP','_CM','_DL','_DO','_RETURN','_ARIN']
#             out_jual = reorder_columns(out_jual, JUAL_ORDER)

#             st.success("Selesai!")
#             st.dataframe(out_jual.head(100))
#             _download_df(out_jual, "DATA_PENJUALAN.xlsx", "⬇️ Unduh DATA_PENJUALAN.xlsx")

# st.divider()
# st.markdown("**Catatan**: Urutan kolom dioutput sudah disusun sesuai preferensi: "
#             "Beli = PO→APDP→Timbang→GRPO→APIN, "
#             "Jual = SO→ARDP→RCM→Timbang→DO→Return→ARIN.")
