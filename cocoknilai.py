# cocoknilai.py
"""
MB Tools - Pencocokan Nilai Siswa (Streamlit minimal)
Mode: Upload → Proses → Download
Tanpa tkinter, tanpa rapidfuzz. Fuzzy matching menggunakan difflib (built-in).
Jika openpyxl tersedia, hasil ditulis ke .xlsx. Jika tidak, ditulis .csv.
"""

import os
import tempfile
from datetime import datetime

import streamlit as st
import pandas as pd
from difflib import SequenceMatcher

# Try to import openpyxl for nicer .xlsx output; if not available we'll fallback to CSV
try:
    import openpyxl  # only to detect availability
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

st.set_page_config(page_title="MB Tools — Cocok Nilai (Minimal)", layout="wide")

st.markdown(
    "<h2>📊 MB Tools — Pencocokan Nilai Siswa (Minimal)</h2>"
    "<div style='color:gray'>Upload 2 file Excel (Respons & Hasil), klik Proses, lalu download hasil.</div>",
    unsafe_allow_html=True,
)

st.write("")  # spacer

col1, col2 = st.columns(2)
with col1:
    uploaded_resp = st.file_uploader("Upload file Respons (Excel)", type=["xlsx", "xls"], key="resp")
with col2:
    uploaded_hasil = st.file_uploader("Upload file Hasil (Excel) — template daftar siswa", type=["xlsx", "xls"], key="hasil")

st.write("")  # spacer

# Simple options
rename_opt = st.checkbox("Auto-rename output with timestamp (recommended)", value=True)
process_btn = st.button("🚀 Proses Data")

# Helpers (same logic)
def parse_score(raw):
    if pd.isna(raw):
        return None
    s = str(raw).strip()
    if "/" in s:
        s = s.split("/")[0].strip()
    if s.endswith("%"):
        s = s[:-1].strip()
    cleaned = "".join(ch for ch in s if (ch.isdigit() or ch in ".,"))
    cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except:
        return None

def normalize_text(s):
    if pd.isna(s):
        return ""
    return " ".join(str(s).strip().lower().split())

def cari_kolom_otomatis(df, keywords):
    cols = list(df.columns)
    for kw in keywords:
        for col in cols:
            name = str(col).strip().lower()
            if kw in name:
                return col
    return None

def similarity_pct(a, b):
    if not a and not b:
        return 100.0
    return SequenceMatcher(None, a, b).ratio() * 100.0

def match_df_bytes(df_resp, df_hasil, log_fn=None):
    # core logic adapted to use DataFrame objects (no file IO here)
    def log(msg):
        if log_fn:
            log_fn(msg)

    name_keys = ["nama", "name"]
    score_keys = ["score", "nilai", "skor"]
    absen_keys = ["absen", "no", "nomor", "nis", "id"]
    time_keys = ["time", "timestamp", "tgl", "waktu"]

    # safe fallbacks
    def col_or(index, df):
        try:
            return df.columns[index]
        except:
            return df.columns[0]

    col_name = cari_kolom_otomatis(df_resp, name_keys) or col_or(2, df_resp)
    col_score = cari_kolom_otomatis(df_resp, score_keys) or col_or(1, df_resp)
    col_time = cari_kolom_otomatis(df_resp, time_keys) or col_or(0, df_resp)
    col_absen = cari_kolom_otomatis(df_resp, absen_keys)

    kolom_nama_hasil = cari_kolom_otomatis(df_hasil, name_keys) or col_or(1, df_hasil)
    kolom_absen_hasil = cari_kolom_otomatis(df_hasil, absen_keys) or col_or(0, df_hasil)

    # normalize
    df_resp = df_resp.copy()
    df_hasil = df_hasil.copy()
    df_resp["_name_norm"] = df_resp[col_name].apply(normalize_text)
    df_hasil["_name_norm"] = df_hasil[kolom_nama_hasil].apply(normalize_text)
    if col_absen in df_resp.columns:
        df_resp["_absen_str"] = df_resp[col_absen].astype(str).fillna("").str.strip()
    else:
        df_resp["_absen_str"] = ""
    df_hasil["_absen_str"] = df_hasil[kolom_absen_hasil].astype(str).fillna("").str.strip()

    # ensure Score_1..6 exist
    for i in range(1, 7):
        colname = f"Score_{i}"
        if colname not in df_hasil.columns:
            df_hasil[colname] = pd.NA

    total_resp = len(df_resp)
    processed = 0

    # matching loop
    for ridx, rrow in df_resp.iterrows():
        processed += 1
        raw_name = rrow["_name_norm"]
        raw_absen = str(rrow["_absen_str"])
        raw_score = parse_score(rrow[col_score])

        matched_idx = None

        # exact/substring name
        for hid, target_norm in df_hasil["_name_norm"].items():
            if raw_name == target_norm or raw_name in target_norm or target_norm in raw_name:
                matched_idx = hid
                break

        # match by absen
        if matched_idx is None and raw_absen:
            for hid, a in df_hasil["_absen_str"].items():
                if str(a).strip() == raw_absen.strip():
                    matched_idx = hid
                    break

        # fuzzy by difflib
        if matched_idx is None and raw_name:
            best_score = 0
            best_idx = None
            for hid, target_norm in df_hasil["_name_norm"].items():
                sc = similarity_pct(raw_name, target_norm)
                if sc > best_score:
                    best_score = sc
                    best_idx = hid
            if best_score >= 65:
                matched_idx = best_idx

        if matched_idx is None:
            log(f"⚠️ Tidak ditemukan: {rrow[col_name]}")
            continue

        # place into next empty Score_1..6
        for i in range(1, 7):
            coln = f"Score_{i}"
            val = df_hasil.at[matched_idx, coln]
            if pd.isna(val) or str(val).strip() == "":
                df_hasil.at[matched_idx, coln] = raw_score
                break

    # compute SCORE per Bayu logic
    def safe_float(x):
        try:
            return float(x)
        except:
            return None

    def hitung_score(row):
        skor_list = [safe_float(row.get(f"Score_{i}")) for i in range(1, 7)]
        s1 = skor_list[0]
        if s1 and s1 >= 80:
            return s1
        elif any(s and s >= 80 for s in skor_list[1:]):
            return 78
        return None

    df_hasil["SCORE"] = df_hasil.apply(hitung_score, axis=1)

    # drop helper cols
    for c in ["_name_norm", "_absen_str"]:
        if c in df_hasil.columns:
            df_hasil.drop(columns=[c], inplace=True)

    return df_hasil

# Processing action
if process_btn:
    if not uploaded_resp or not uploaded_hasil:
        st.warning("Silakan upload kedua file (Respons & Hasil) terlebih dahulu.")
    else:
        try:
            # read uploaded files to DataFrames
            df_resp = pd.read_excel(uploaded_resp, engine="openpyxl") if uploaded_resp.name.lower().endswith(("xlsx","xls")) else pd.read_csv(uploaded_resp)
            df_hasil = pd.read_excel(uploaded_hasil, engine="openpyxl") if uploaded_hasil.name.lower().endswith(("xlsx","xls")) else pd.read_csv(uploaded_hasil)

            st.info("Memproses... (ini mungkin memakan beberapa detik tergantung ukuran file)")

            # simple logger collector
            log_msgs = []
            def log_fn(msg):
                log_msgs.append(msg)

            result_df = match_df_bytes(df_resp, df_hasil, log_fn=log_fn)

            # prepare output file (xlsx if openpyxl available, else csv)
            out_name = "hasil_pencocokan"
            if rename_opt:
                tstamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_name = f"{out_name}_{tstamp}"
            tmpdir = tempfile.mkdtemp(prefix="mbtools_")
            out_path_xlsx = os.path.join(tmpdir, out_name + ".xlsx")
            out_path_csv = os.path.join(tmpdir, out_name + ".csv")

            if HAS_OPENPYXL:
                # write xlsx
                result_df.to_excel(out_path_xlsx, index=False, engine="openpyxl")
                with open(out_path_xlsx, "rb") as f:
                    data = f.read()
                st.success("✅ Proses selesai! (output: .xlsx)")
                st.download_button("⬇️ Download hasil_pencocokan.xlsx", data, file_name=os.path.basename(out_path_xlsx), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                # fallback CSV
                result_df.to_csv(out_path_csv, index=False)
                with open(out_path_csv, "rb") as f:
                    data = f.read()
                st.success("✅ Proses selesai! (openpyxl tidak terpasang — output: .csv)")
                st.download_button("⬇️ Download hasil_pencocokan.csv", data, file_name=os.path.basename(out_path_csv), mime="text/csv")

            # show short preview of results
            st.markdown("**Preview hasil (5 baris pertama)**")
            st.dataframe(result_df.head(5))

            # show logs (if any)
            if log_msgs:
                st.markdown("**Log**")
                for m in log_msgs[-50:]:
                    st.write(m)

        except Exception as e:
            st.error(f"Terjadi error saat memproses: {e}")
            raise
