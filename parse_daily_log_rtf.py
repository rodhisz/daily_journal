# -*- coding: utf-8 -*-
"""
Parse catatan aktivitas harian dari file RTF menjadi Excel.

Fitur:
- Baca dari RTF (striprtf)
- Header tanggal:
    "Rabu, 10 December 2025"  atau  "Rabu 10 December 2025"
- Format tanggal output: MM/DD/YYYY (contoh: 10 December 2025 -> 12/10/2025)
- List bernomor boleh "1." atau "1,"
- Split "Project -> Issue" atau "Project - Issue"
- Jika tidak ada project -> Project = "Other"
- Status:
    ada 🟡 / ❌  -> "In Progress"
    tidak ada   -> "Done"
- Cleaning karakter ilegal supaya Excel tidak corrupt
"""

from datetime import date
import re
import pandas as pd

# =================== KONFIGURASI FILE ===================
INPUT_FILE = "raw_notes.rtf"          # ganti kalau nama file beda
OUTPUT_FILE = "daily_activities_clean.xlsx"


# =================== CLEANING SUPER-KETAT ===================

def clean_text(s):
    """
    Bersihkan string dari karakter yang bisa bikin file Excel rusak.
    - Hapus ASCII control (0–31 kecuali tab/newline/CR)
    - Hapus Unicode surrogate (D800–DFFF)
    - Hapus karakter non-printable lainnya
    """
    if not isinstance(s, str):
        return s

    # 1. Hapus ASCII control chars (0–31) kecuali tab(9), newline(10), carriage return(13)
    s = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', s)

    # 2. Hapus Unicode surrogate pairs
    s = re.sub(r'[\uD800-\uDFFF]', '', s)

    # 3. Buang karakter non-printable lain
    s = ''.join(ch for ch in s if ch.isprintable())

    # 4. Rapikan spasi
    return s.strip()


# =================== MAPPING HARI & BULAN ===================

day_map = {
    'senin': 'Monday',
    'selasa': 'Tuesday',
    'selesa': 'Tuesday',   # typo yang sering muncul
    'rabu': 'Wednesday',
    'kamis': 'Thursday',
    'kami': 'Thursday',    # typo
    'jumat': 'Friday',
    "jum'at": 'Friday',
    'sabtu': 'Saturday',
    'minggu': 'Sunday',
}

month_map = {
    'januari': 1, 'january': 1, 'jan': 1,
    'februari': 2, 'february': 2, 'feb': 2,
    'maret': 3, 'march': 3, 'mar': 3,
    'april': 4, 'apr': 4,
    'mei': 5, 'may': 5,
    'juni': 6, 'june': 6, 'jun': 6,
    'juli': 7, 'july': 7, 'jul': 7,
    'agustus': 8, 'august': 8, 'aug': 8,
    'september': 9, 'sept': 9, 'sep': 9,
    'oktober': 10, 'october': 10, 'oct': 10,
    'november': 11, 'nov': 11,
    'desember': 12, 'december': 12, 'dec': 12,
}

# Contoh yang ditangkap:
# "Rabu, 10 December 2025"
# "Rabu 10 December 2025"
header_re = re.compile(
    r'^\s*([A-Za-zÀ-ÿ]+)[,]?\s+(\d{1,2})\s+([A-Za-zÀ-ÿ]+)\s+(\d{4})\s*$'
)


def parse_header(line: str):
    """Parse baris header tanggal menjadi (date_obj, day_english)."""
    m = header_re.match(line.strip())
    if not m:
        return None
    dname, d_str, mon_str, y_str = m.groups()

    dname = dname.lower()
    mon_str = mon_str.lower()

    if dname not in day_map:
        return None
    if mon_str not in month_map:
        return None

    dt = date(int(y_str), month_map[mon_str], int(d_str))
    return dt, day_map[dname]


# =================== RTF -> TEXT ===================

def read_rtf(path: str) -> str:
    """Baca file RTF dan convert ke plain text."""
    try:
        from striprtf.striprtf import rtf_to_text
    except ImportError:
        raise ImportError(
            "Module 'striprtf' belum terinstall.\n"
            "Install dulu dengan perintah:\n"
            "  pip install striprtf\n"
        )

    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        raw = f.read()
    txt = rtf_to_text(raw)
    return txt


# =================== PARSER UTAMA ===================

def parse_all(text: str):
    """
    Parse seluruh isi text menjadi list of dict:
    {Date, Day, Project, Issue, Status}
    """
    rows = []
    cur_date = None
    cur_day = None

    for raw_line in text.splitlines():
        line = clean_text(raw_line)
        if not line:
            continue

        # 1) Cek apakah ini header tanggal
        hdr = parse_header(line)
        if hdr:
            cur_date, cur_day = hdr
            continue

        # Kalau belum ada tanggal aktif, skip
        if cur_date is None:
            continue

        # 2) Cek numbering: "1. text" atau "1, text"
        m = re.match(r'^\s*\d+[\.,]\s*(.*)$', line)
        content = m.group(1) if m else line

        if not content:
            continue

        # 3) Split Project & Issue
        if "->" in content:
            p, i = content.split("->", 1)
        elif " - " in content:
            p, i = content.split(" - ", 1)
        else:
            p, i = "", content

        p = clean_text(p)
        i = clean_text(i)

        # Project kosong -> Other
        if not p.strip():
            p = "Other"

        # 4) Status dari emoji
        status = "In Progress" if ("🟡" in i or "❌" in i) else "Done"
        i = i.replace("🟡", "").replace("❌", "").strip()

        # 5) Simpan row, tanggal format MM/DD/YYYY
        rows.append({
            "Date": cur_date.strftime("%m/%d/%Y"),  # MM/DD/YYYY
            "Day": cur_day,
            "Project": p,
            "Issue": i,
            "Status": status,
        })

    return rows


# =================== MAIN ===================

def main():
    print(f"Reading RTF from: {INPUT_FILE}")
    text = read_rtf(INPUT_FILE)

    print("Parsing text...")
    rows = parse_all(text)
    print(f"Total rows parsed: {len(rows)}")

    if not rows:
        print("Tidak ada baris yang berhasil diparsing. Cek lagi format input.")
        return

    df = pd.DataFrame(rows, columns=["Date", "Day", "Project", "Issue", "Status"])

    # Cleaning extra safety sebelum masuk Excel
    df = df.applymap(clean_text)

    print(f"Saving to Excel: {OUTPUT_FILE}")
    df.to_excel(OUTPUT_FILE, index=False)

    print("Done.\nPreview 10 baris pertama:")
    try:
        print(df.head(10).to_string(index=False))
    except Exception:
        print(df.head(10))


if __name__ == "__main__":
    main()
