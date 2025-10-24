# cocoknilai.py
"""
MB Tools - Pencocokan Nilai Siswa (Streamlit)
Converted for Streamlit by ChatGPT for Bayu
- Menggunakan difflib (built-in) sebagai pengganti rapidfuzz
- Upload 2 file Excel (Respons & Hasil), proses, dan download hasil_pencocokan.xlsx
"""

import os
import tempfile
from datetime import datetime
import io

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from difflib import SequenceMatcher

# -------------------------
# Helpers
# -------------------------
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

def apply_color(cell, value):
    if value is None:
        return
    try:
        val = float(value)
    except:
        return
    if val < 78:
        cell.fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")  # merah
    elif val < 80:
        cell.fill = PatternFill(start_color="FFF58C", end_color="FFF58C", fill_type="solid")  # kuning
    else:
        cell.fill = PatternFill(start_color="B7E1A1", end_color="B7E1A1", fill_type="solid")  # hijau

# -------------------------
# Core logic (file paths)
# -------------------------
def match_and_write(respons_path, hasil_path, log_fn=None, progress_fn=None):
    def log(msg):
        if log_fn:
            log_fn(msg)
        else:
            print(msg)

    log("🔁 Membaca file Excel...")
    df_resp = pd.read_excel(respons_path, engine="openpyxl")
    df_hasil = pd.read_excel(hasil_path, engine="openpyxl")

    # deteksi kolom
    resp_cols_original = list(df_resp.columns)
    hasil_cols_original = list(df_hasil.columns)

    name_keys = ["nama", "name"]
    score_keys = ["score", "nilai", "skor"]
    absen_keys = ["absen", "no", "nomor", "nis", "id"]
    time_keys = ["time", "timestamp", "tgl", "waktu"]

    # safe fallback
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
    df_resp["_name_norm"] = df_resp[col_name].apply(normalize_text)
    df_hasil["_name_norm"] = df_hasil[kolom_nama_hasil].apply(normalize_text)
    if col_absen in df_resp.columns:
        df_resp["_absen_str"] = df_resp[col_absen].astype(str).fillna("").str.strip()
    else:
        df_resp["_absen_str"] = ""
    df_hasil["_absen_str"] = df_hasil[kolom_absen_hasil].astype(str).fillna("").str.strip()

    # siapkan kolom Score_1..6 jika belum ada
    for i in range(1, 7):
        colname = f"Score_{i}"
        if colname not in df_hasil.columns:
            df_hasil[colname] = pd.NA

    total_resp = len(df_resp)
    processed = 0

    log("🔁 Mencocokkan data siswa...")
    for ridx, rrow in df_resp.iterrows():
        processed += 1
        if progress_fn:
            progress_fn(int(processed / max(total_resp, 1) * 100))

        raw_name = rrow["_name_norm"]
        raw_absen = str(rrow["_absen_str"])
        raw_score = parse_score(rrow[col_score])

        matched_idx = None

        # 1) nama exact / substring
        for hid, target_norm in df_hasil["_name_norm"].items():
            if raw_name == target_norm or raw_name in target_norm or target_norm in raw_name:
                matched_idx = hid
                break

        # 2) cocokkan absen
        if matched_idx is None and raw_absen:
            for hid, a in df_hasil["_absen_str"].items():
                if str(a).strip() == raw_absen.strip():
                    matched_idx = hid
                    break

        # 3) fuzzy using difflib
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

        # isi ke Score_1..6 pada baris hasil
        for i in range(1, 7):
            coln = f"Score_{i}"
            val = df_hasil.at[matched_idx, coln]
            if pd.isna(val) or str(val).strip() == "":
                df_hasil.at[matched_idx, coln] = raw_score
                break

    # hitung SCORE sesuai aturan Bayu
    log("🧮 Menghitung kolom SCORE otomatis...")
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

    # simpan hasil ke temp file
    out_dir = os.path.dirname(os.path.abspath(hasil_path))
    out_name = os.path.join(out_dir, "hasil_pencocokan.xlsx")
    df_save = df_hasil.copy()
    for c in ["_name_norm", "_absen_str"]:
        if c in df_save.columns:
            df_save.drop(columns=[c], inplace=True)
    df_save.to_excel(out_name, index=False)

    # pewarnaan Score_1..Score_6 (openpyxl)
    try:
        wb = load_workbook(out_name)
        ws = wb.active
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        score_cols_idx = []
        for i, h in enumerate(headers, start=1):
            if h and isinstance(h, str) and h.strip().lower().startswith("score_"):
                score_cols_idx.append(i)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for idx in score_cols_idx:
                cell = row[idx - 1]
                apply_color(cell, cell.value)
        wb.save(out_name)
    except Exception as e:
        log(f"⚠️ Pewarnaan gagal: {e}")

    log(f"✅ Selesai! File disimpan di: {out_name}")
    if progress_fn:
        progress_fn(100)
    return out_name

# -------------------------
# Streamlit UI
# -------------------------
st.set_page_config(page_title="MB Tools — Pencocokan Nilai Siswa", layout="wide")

st.markdown(
    """
    <div style="background:#1f6feb;padding:14px;border-radius:8px">
      <h2 style="color:white;margin:0">📊 MB Tools — Pencocokan Nilai Siswa</h2>
      <div style="color:#e6f0ff">Credit: Apps by MB — Donasi: wa.me/628522939579</div>
    </div>
    """,
    unsafe_allow_html=True
)

st.write("")  # spacer

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📂 Upload File Respons")
    uploaded_resp = st.file_uploader("Pilih file respons (Excel)", type=["xlsx", "xls"], key="resp")
    if uploaded_resp:
        try:
            df_preview_resp = pd.read_excel(uploaded_resp, engine="openpyxl", nrows=5)
            st.markdown("**Preview (5 baris pertama)**")
            st.dataframe(df_preview_resp)
        except Exception as e:
            st.warning(f"Gagal membaca preview respons: {e}")

with col2:
    st.subheader("📘 Upload File Hasil (template daftar siswa)")
    uploaded_hasil = st.file_uploader("Pilih file hasil (Excel)", type=["xlsx", "xls"], key="hasil")
    if uploaded_hasil:
        try:
            df_preview_hasil = pd.read_excel(uploaded_hasil, engine="openpyxl", nrows=5)
            st.markdown("**Preview (5 baris pertama)**")
            st.dataframe(df_preview_hasil)
        except Exception as e:
            st.warning(f"Gagal membaca preview hasil: {e}")

st.write("")

colA, colB = st.columns([1, 2])
with colA:
    rename_opt = st.checkbox("Auto-rename output dengan timestamp", value=True)
with colB:
    st.write("")  # alignment
    process_btn = st.button("🚀 Proses Data", use_container_width=True)

# placeholders
log_box = st.empty()
prog_placeholder = st.empty()

def save_uploaded_to_temp(uploaded_file, prefix):
    if uploaded_file is None:
        return None
    tmpdir = tempfile.mkdtemp(prefix="mbtools_")
    ext = os.path.splitext(uploaded_file.name)[1] or ".xlsx"
    tmp_path = os.path.join(tmpdir, prefix + ext)
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return tmp_path

if process_btn:
    if not uploaded_resp or not uploaded_hasil:
        st.warning("Silakan upload kedua file (Respons & Hasil) terlebih dahulu.")
    else:
        logs = []
        def log_fn(msg):
            ts = datetime.now().strftime("%H:%M:%S")
            entry = f"[{ts}] {msg}"
            logs.append(entry)
            # show last 200 lines
            log_box.text("\n".join(logs[-200:]))

        prog = prog_placeholder.progress(0)
        pct_text = prog_placeholder.empty()

        def progress_fn(p):
            try:
                prog.progress(p)
                pct_text.text(f"{p}%")
            except:
                pass

        try:
            path_resp = save_uploaded_to_temp(uploaded_resp, "respons")
            path_hasil = save_uploaded_to_temp(uploaded_hasil, "hasil")
            out_path = match_and_write(path_resp, path_hasil, log_fn=log_fn, progress_fn=progress_fn)

            # prepare download
            with open(out_path, "rb") as f:
                data = f.read()
            fname = "hasil_pencocokan.xlsx"
            if rename_opt:
                tstamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                fname = f"hasil_pencocokan_{tstamp}.xlsx"

            st.success("✅ Proses selesai!")
            st.download_button("⬇️ Download Hasil Pencocokan", data, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            log_fn(f"File output: {out_path}")
        except Exception as e:
            st.error(f"❌ Terjadi error: {e}")
            log_fn(f"ERROR: {e}")
        finally:
            try:
                prog_placeholder.empty()
            except:
                pass

st.write("")
st.markdown("<small>Apps by MB — Donasi/Support: wa.me/628522939579</small>", unsafe_allow_html=True)
