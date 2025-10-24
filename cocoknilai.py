"""
MB Tools - Pencocokan Nilai Siswa (FINAL+)
- Deteksi kolom respons & hasil (toleran)
- Fuzzy match nama & absen
- Multi respons (Score_1 ... Score_6)
- Output: hasil_pencocokan.xlsx
- Otomatis menambahkan kolom SCORE sesuai logika Bayu:
    - Jika Score_1 >= 80 → SCORE = Score_1
    - Jika Score_1 < 80 tapi ada Score_2..Score_6 >= 80 → SCORE = 78
    - Jika semua < 80 → SCORE kosong
Credit: Apps by MB (Donasi: wa.me/628522939579)
"""

import os
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from rapidfuzz import fuzz

# -------------------------
# Util helpers
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

# -------------------------
# Core matching logic
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

    col_name = cari_kolom_otomatis(df_resp, name_keys) or resp_cols_original[2]
    col_score = cari_kolom_otomatis(df_resp, score_keys) or resp_cols_original[1]
    col_time = cari_kolom_otomatis(df_resp, time_keys) or resp_cols_original[0]
    col_absen = cari_kolom_otomatis(df_resp, absen_keys)

    kolom_nama_hasil = cari_kolom_otomatis(df_hasil, name_keys) or hasil_cols_original[1]
    kolom_absen_hasil = cari_kolom_otomatis(df_hasil, absen_keys) or hasil_cols_original[0]

    # normalize
    df_resp["_name_norm"] = df_resp[col_name].apply(normalize_text)
    df_hasil["_name_norm"] = df_hasil[kolom_nama_hasil].apply(normalize_text)
    df_resp["_absen_str"] = df_resp[col_absen].astype(str).fillna("").str.strip() if col_absen in df_resp.columns else ""
    df_hasil["_absen_str"] = df_hasil[kolom_absen_hasil].astype(str).fillna("").str.strip()

    # siapkan kolom Score_1..6 jika belum ada
    for i in range(1, 7):
        colname = f"Score_{i}"
        if colname not in df_hasil.columns:
            df_hasil[colname] = pd.NA

    total_resp = len(df_resp)
    processed = 0

    # proses pencocokan
    log("🔁 Mencocokkan data siswa...")
    for ridx, rrow in df_resp.iterrows():
        processed += 1
        if progress_fn:
            progress_fn(int(processed / max(total_resp, 1) * 100))

        raw_name = rrow["_name_norm"]
        raw_absen = str(rrow["_absen_str"])
        raw_score = parse_score(rrow[col_score])

        matched_idx = None

        # 1) cocokkan nama langsung
        for hid, target_norm in df_hasil["_name_norm"].items():
            if raw_name == target_norm or raw_name in target_norm or target_norm in raw_name:
                matched_idx = hid
                break

        # 2) kalau belum, cocokkan absen
        if matched_idx is None and raw_absen:
            for hid, a in df_hasil["_absen_str"].items():
                if str(a).strip() == raw_absen.strip():
                    matched_idx = hid
                    break

        # 3) fuzzy match (nama mirip)
        if matched_idx is None and raw_name:
            best_score = 0
            best_idx = None
            for hid, target_norm in df_hasil["_name_norm"].items():
                sc = fuzz.token_set_ratio(raw_name, target_norm)
                if sc > best_score:
                    best_score = sc
                    best_idx = hid
            if best_score >= 65:
                matched_idx = best_idx

        if matched_idx is None:
            log(f"⚠️ Tidak ditemukan: {rrow[col_name]}")
            continue

        # tempatkan ke kolom kosong berikutnya
        for i in range(1, 7):
            coln = f"Score_{i}"
            if pd.isna(df_hasil.at[matched_idx, coln]) or str(df_hasil.at[matched_idx, coln]).strip() == "":
                df_hasil.at[matched_idx, coln] = raw_score
                break

    # -----------------------------
    # Hitung kolom SCORE otomatis
    # -----------------------------
    log("🧮 Menghitung kolom SCORE otomatis...")
    def hitung_score(row):
        s1 = row.get("Score_1")
        skor_list = [row.get(f"Score_{i}") for i in range(1, 7)]
        skor_valid = [v for v in skor_list if pd.notna(v)]
        if not skor_valid:
            return None
        if s1 is not None and s1 >= 80:
            return s1
        for s in skor_list[1:]:
            if s is not None and s >= 80:
                return 78
        return None

    df_hasil["SCORE"] = df_hasil.apply(hitung_score, axis=1)

    # simpan hasil
    out_name = os.path.join(os.path.dirname(os.path.abspath(hasil_path)), "hasil_pencocokan.xlsx")
    for c in ["_name_norm", "_absen_str"]:
        if c in df_hasil.columns:
            df_hasil.drop(columns=[c], inplace=True)
    df_hasil.to_excel(out_name, index=False)
    log(f"✅ Selesai. File disimpan: {out_name}")
    if progress_fn:
        progress_fn(100)
    return out_name

# -------------------------
# GUI (Modern)
# -------------------------
class MatcherGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MB Tools — Pencocokan Nilai Siswa (final+)")
        self.geometry("760x520")
        self.configure(bg="#f3f7fb")

        header = tk.Frame(self, bg="#1f6feb", height=90)
        header.pack(fill="x")
        tk.Label(header, text="Pencocokan Nilai Siswa — MB Tools", bg="#1f6feb", fg="white",
                 font=("Segoe UI", 18, "bold")).place(x=20, y=18)
        tk.Label(header, text="Credit Apps by MB — Donasi: 08522939579", bg="#1f6feb", fg="white",
                 font=("Segoe UI", 9)).place(x=22, y=50)

        body = tk.Frame(self, bg=self["bg"])
        body.pack(fill="both", expand=True, padx=18, pady=12)

        tk.Button(body, text="📂 Pilih File Respons", command=self.pick_respons).grid(row=0, column=0, sticky="w", pady=8)
        self.lbl_respons = tk.Entry(body, width=72)
        self.lbl_respons.grid(row=0, column=1, padx=8, pady=8)

        tk.Button(body, text="📂 Pilih File Hasil", command=self.pick_hasil).grid(row=1, column=0, sticky="w", pady=8)
        self.lbl_hasil = tk.Entry(body, width=72)
        self.lbl_hasil.grid(row=1, column=1, padx=8, pady=8)

        self.btn_process = tk.Button(body, text="🔁 Proses Data", bg="#0b78e3", fg="white", command=self.start_process)
        self.btn_process.grid(row=2, column=1, sticky="w", pady=10)

        self.progress = ttk.Progressbar(body, orient="horizontal", length=640, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=2, pady=8)
        self.lbl_pct = tk.Label(body, text="0%", bg=self["bg"])
        self.lbl_pct.grid(row=4, column=0, sticky="w", pady=2)

        tk.Label(body, text="Log:", bg=self["bg"]).grid(row=5, column=0, sticky="w")
        self.txt_log = tk.Text(body, height=14, width=92)
        self.txt_log.grid(row=6, column=0, columnspan=2, pady=6)

    def pick_respons(self):
        p = filedialog.askopenfilename(title="Pilih file respons", filetypes=[("Excel files", "*.xlsx *.xls")])
        if p:
            self.lbl_respons.delete(0, tk.END)
            self.lbl_respons.insert(0, p)

    def pick_hasil(self):
        p = filedialog.askopenfilename(title="Pilih file hasil", filetypes=[("Excel files", "*.xlsx *.xls")])
        if p:
            self.lbl_hasil.delete(0, tk.END)
            self.lbl_hasil.insert(0, p)

    def log(self, msg):
        timestamp = time.strftime("%H:%M:%S")
        self.txt_log.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.txt_log.see(tk.END)
        self.update_idletasks()

    def set_progress(self, pct):
        self.progress["value"] = pct
        self.lbl_pct.config(text=f"{pct}%")
        self.update_idletasks()

    def start_process(self):
        respons = self.lbl_respons.get().strip()
        hasil = self.lbl_hasil.get().strip()
        if not respons or not hasil:
            messagebox.showwarning("File belum lengkap", "Silakan pilih file respons dan hasil terlebih dahulu.")
            return

        self.btn_process.config(state="disabled")
        self.txt_log.delete("1.0", tk.END)
        threading.Thread(target=self._run_match, args=(respons, hasil), daemon=True).start()

    def _run_match(self, respons, hasil):
        try:
            out = match_and_write(respons, hasil, log_fn=self.log, progress_fn=self.set_progress)
            messagebox.showinfo("Selesai", f"Proses selesai.\nFile tersimpan: {out}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.log(f"ERROR: {e}")
        finally:
            self.btn_process.config(state="normal")
            self.set_progress(0)

if __name__ == "__main__":
    app = MatcherGUI()
    app.mainloop()
