# cocoknilai_csv.py
"""
MB Tools - Cocok Nilai (CSV)
- Input: 2x CSV (jawaban siswa & kunci)
- Deteksi kolom otomatis (sesuai format contoh)
- 1 poin per jawaban benar (case-insensitive, trim)
- Preview, hasil, dan download CSV hasil
"""

import io
import tempfile
from datetime import datetime

import streamlit as st
import pandas as pd

st.set_page_config(page_title="MB Tools - Cocok Nilai (CSV)", layout="wide")

st.markdown(
    """
    <div style="background:#1f6feb;padding:12px;border-radius:8px">
      <h2 style="color:white;margin:0">📊 MB Tools — Cocok Nilai (CSV)</h2>
      <div style="color:#e6f0ff">Upload CSV Jawaban Siswa & CSV Kunci → proses → download hasil</div>
    </div>
    """,
    unsafe_allow_html=True
)

st.write("")
col1, col2 = st.columns([1, 1])
with col1:
    uploaded_siswa = st.file_uploader("Upload CSV Jawaban Siswa (contoh: NIS, Nama, No1, No2, ...)", type=["csv"], key="siswa")
with col2:
    uploaded_kunci = st.file_uploader("Upload CSV Kunci Jawaban (contoh: No, Kunci) atau single-row Kunci", type=["csv"], key="kunci")

st.write("")
opts_col1, opts_col2 = st.columns([1, 2])
with opts_col1:
    rename_opt = st.checkbox("Auto-rename output (timestamp)", value=True)
with opts_col2:
    st.caption("Deteksi otomatis kolom pertanyaan. Sistem memberi 1 poin per jawaban yang cocok (case-insensitive).")

process_btn = st.button("🚀 Proses & Hitung Nilai", use_container_width=True)

def normalize_answer(x):
    try:
        if pd.isna(x):
            return ""
        return str(x).strip().upper()
    except:
        return str(x).strip().upper()

def detect_question_columns(df_siswa):
    # Heuristik: kolom yang bukan 'nis','nama','no','id' dianggap kolom soal
    low = [str(c).strip().lower() for c in df_siswa.columns]
    ignore = set(["nis", "nama", "name", "no", "id", "nomor", "kelas"])
    q_cols = [c for c, l in zip(df_siswa.columns, low) if l not in ignore]
    # if first two columns are NIS & Nama, skip them; otherwise if many columns, take columns from 3..end
    if len(df_siswa.columns) >= 3 and (low[0] in ["nis", "id", "no"] or low[1] in ["nama", "name"]):
        # assume question cols start at index 2 (0-based)
        q_cols = list(df_siswa.columns[2:])
    return q_cols

def load_kunci_from_df(df_kunci, q_cols):
    """
    Support two kunci formats:
    1) Two-column table: 'No' | 'Kunci' (No numeric)
    2) Single row header: columns named 1,2,3 or No1,No2 ...
    3) One-row where first row contains kunci values (e.g., header generic)
    """
    # 1) try two-column detection
    cols = [str(c).strip().lower() for c in df_kunci.columns]
    if len(df_kunci.columns) >= 2 and any("kunci" in c for c in cols) or ("jawaban" in "".join(cols)):
        # try to find kunci column
        # find index of "kunci" like column
        kidx = None
        noidx = None
        for i, c in enumerate(cols):
            if "kunci" in c or "jawaban" in c or "answer" in c:
                kidx = i
            if c in ("no", "nomor", "index", "no."):
                noidx = i
        if kidx is not None:
            # build dict: no -> kunci
            try:
                ser_no = df_kunci.iloc[:, noidx] if noidx is not None else pd.Series(range(1, len(df_kunci)+1))
                ser_k = df_kunci.iloc[:, kidx]
                mapping = {}
                for i, (n, k) in enumerate(zip(ser_no, ser_k)):
                    keynum = str(int(n)) if pd.notna(n) and str(n).strip().isdigit() else str(i+1)
                    mapping[keynum] = normalize_answer(k)
                return mapping
            except Exception:
                pass

    # 2) try single row with question-number headers matching q_cols length
    # If df_kunci has one row and many columns, treat as answers row
    if df_kunci.shape[0] == 1 and df_kunci.shape[1] >= 1:
        mapping = {}
        # if headers look numeric or "No1" style, use headers; else use ordinal 1..n
        headers = list(df_kunci.columns)
        for i, h in enumerate(headers):
            num = None
            hn = str(h).strip()
            # try extract number
            import re
            m = re.search(r"(\d+)", hn)
            if m:
                num = m.group(1)
            else:
                num = str(i+1)
            mapping[num] = normalize_answer(df_kunci.iloc[0, i])
        return mapping

    # 3) fallback: if q_cols provided, try to match by position
    if len(q_cols) > 0:
        mapping = {}
        # if kunci file has single column of values, map by order
        flat = []
        # flatten all cells row-wise
        for _, row in df_kunci.iterrows():
            for v in row:
                flat.append(v)
        for i, v in enumerate(flat):
            mapping[str(i+1)] = normalize_answer(v)
        return mapping

    # final fallback: empty mapping
    return {}

def build_result(df_siswa, kunci_map, q_cols):
    # kunci_map keys likely are "1","2",... ; q_cols are column names in df_siswa
    results = []
    # create per-question correctness columns
    for _, row in df_siswa.iterrows():
        row_result = row.to_dict()
        total = 0
        perq = {}
        # iterate over question columns in order
        for i, col in enumerate(q_cols, start=1):
            qnum = str(i)
            student_ans = normalize_answer(row.get(col, ""))
            key = kunci_map.get(qnum, "")
            correct = 1 if (student_ans != "" and key != "" and student_ans == key) else 0
            perq[col] = correct
            total += correct
            # also store normalized answers and key for debugging if needed
            row_result[f"ans_{qnum}"] = student_ans
            row_result[f"key_{qnum}"] = key
            row_result[f"ok_{qnum}"] = correct
        row_result["TOTAL_BENAR"] = total
        row_result["PERSENTASE"] = round((total / max(len(q_cols),1)) * 100, 2)
        results.append(row_result)
    res_df = pd.DataFrame(results)
    # order columns: keep original NIS/Nama first if present, then per-question columns, then totals
    # We will not reorder aggressively; just return the DF
    return res_df

if process_btn:
    if (not uploaded_siswa) or (not uploaded_kunci):
        st.warning("Upload kedua file CSV (siswa & kunci) terlebih dahulu.")
    else:
        try:
            # read CSVs
            df_siswa = pd.read_csv(uploaded_siswa)
            df_kunci = pd.read_csv(uploaded_kunci)

            st.info("File terbaca. Mendeteksi kolom soal...")

            q_cols = detect_question_columns(df_siswa)
            if not q_cols:
                st.error("Gagal mendeteksi kolom soal pada file siswa. Pastikan format: NIS, Nama, No1, No2, ... atau set kolom soal dimulai kolom ke-3.")
                st.stop()

            st.success(f"Terdeteksi {len(q_cols)} kolom soal. Contoh kolom: {q_cols[:6]}")

            # load kunci
            kmap = load_kunci_from_df(df_kunci, q_cols)
            if not kmap:
                st.warning("Gagal mendeteksi kunci otomatis. Pastikan kunci dalam format 'No,Kunci' atau single-row kunci.")
                # show preview of kunci for debugging
                st.markdown("**Preview Kunci (file kunci)**")
                st.dataframe(df_kunci.head(8))
            else:
                st.success(f"Kunci terdeteksi untuk {len(kmap)} soal (menggunakan nomor soal sebagai urutan).")

            # build result
            res_df = build_result(df_siswa, kmap, q_cols)

            # prepare download
            out_name = "hasil_pencocokan"
            if rename_opt:
                tstamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_name = f"{out_name}_{tstamp}"
            csv_bytes = res_df.to_csv(index=False).encode("utf-8")

            st.success("✅ Proses selesai — hasil siap diunduh.")
            st.download_button("⬇️ Download Hasil (CSV)", csv_bytes, file_name=out_name + ".csv", mime="text/csv")

            st.markdown("**Preview hasil (5 baris pertama)**")
            st.dataframe(res_df.head(10))

            # also show simple summary (class average)
            avg = res_df["PERSENTASE"].mean() if "PERSENTASE" in res_df.columns else None
            st.markdown(f"**Rata-rata kelas:** {round(avg,2)}%") if avg is not None else None

        except Exception as e:
            st.error(f"Terjadi error saat memproses: {e}")
            raise
