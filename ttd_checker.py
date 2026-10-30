# -*- coding: utf-8 -*-
"""
ttd_checker.py - Cek TANDA TANGAN DOKTER pada hasil Lab PK, hasil Lab PA, dan Resep.

Hasilnya HANYA untuk kolom khusus di Excel (tidak memengaruhi status LENGKAP / TIDAK LENGKAP).

Tanda tangan dianggap ADA bila di area tanda tangan dokter terdapat:
  - gambar (tanda tangan scan, QR, atau tanda tangan digital berupa gambar), atau
  - coretan vektor (kurva), atau
  - anotasi / kolom tanda tangan digital pada PDF.
Teks nama yang tercetak saja TIDAK dihitung.

Letak area (dari contoh PDF asli):
  RESEP   : kanan atas, di bawah tulisan "Nama dan Tanda Tangan"
  LAB PA  : di bawah tulisan "Ahli Patologi Anatomi"
  LAB PK  : halaman TERAKHIR, di bawah tulisan "Pemeriksa"
"""
import os
import re

import warnings
warnings.filterwarnings("ignore")
try:
    import pymupdf as fitz
except ImportError:  # pragma: no cover
    import fitz

# jenis: (label, pola nama file, kata penanda berurutan, area relatif terhadap penanda (kiri, atas, kanan, bawah), halaman)
JENIS = [
    ("RESEP", re.compile(r"^\d+\s*-\s*RESEP DOKTER", re.I), ["Tanda", "Tangan"], (-120, -6, 40, 75), "pertama"),
    ("LAB PK", re.compile(r"^\d+\s*-\s*HASIL LABORATORIUM PK", re.I), ["Pemeriksa"], (-110, 0, 110, 100), "terakhir"),
    ("LAB PA", re.compile(r"^\d+\s*-\s*HASIL LABORATORIUM PA", re.I), ["Ahli", "Patologi"], (-70, 0, 200, 115), "pertama"),
]
MIN_UKURAN = 12          # gambar lebih kecil dari ini (titik) diabaikan
MIN_OVERLAP = 0.25       # bagian gambar yang harus masuk area


def _cari_penanda(page, kata):
    """Rect penanda (gabungan kata berurutan) atau None; kemunculan TERAKHIR di halaman."""
    w = page.get_text("words")
    n = len(kata)
    ketemu = None
    for i in range(len(w) - n + 1):
        if all(w[i + j][4].strip(":,.").lower() == kata[j].lower() for j in range(n)):
            ketemu = fitz.Rect(w[i][0], min(x[1] for x in w[i:i + n]), w[i + n - 1][2], max(x[3] for x in w[i:i + n]))
    return ketemu


def _overlap(r, area):
    i = r & area
    if i.is_empty or r.get_area() <= 0:
        return 0.0
    return i.get_area() / r.get_area()


def _ada_ttd(page, area):
    for im in page.get_image_info():
        r = fitz.Rect(im["bbox"])
        if r.width < MIN_UKURAN or r.height < MIN_UKURAN:
            continue
        if _overlap(r, area) >= MIN_OVERLAP:
            return True
    try:
        for dr in page.get_drawings():
            r = fitz.Rect(dr["rect"])
            if r.width < 8 and r.height < 8:
                continue
            if any(it[0] == "c" for it in dr["items"]) and _overlap(r, area) >= 0.6:
                return True       # coretan vektor
    except Exception:
        pass
    try:
        for a in page.annots() or []:
            if _overlap(fitz.Rect(a.rect), area) > 0.2:
                return True
        for wd in page.widgets() or []:
            if (wd.field_type_string or "").lower() == "signature" and _overlap(fitz.Rect(wd.rect), area) > 0.2:
                return True
    except Exception:
        pass
    return False


def periksa_file(path, kata, rel, halaman):
    """'ada' | 'tidak' | 'tidak_terbaca' (penanda tidak ditemukan)."""
    with fitz.open(path) as d:
        urut = range(len(d) - 1, -1, -1) if halaman == "terakhir" else range(len(d))
        for i in urut:
            pg = d[i]
            m = _cari_penanda(pg, kata)
            if m is None:
                continue
            area = fitz.Rect(m.x0 + rel[0], m.y0 + rel[1] + 0, m.x1 + rel[2], m.y1 + rel[3])
            return "ada" if _ada_ttd(pg, area) else "tidak"
    return "tidak_terbaca"


def periksa_ttd(folder_pasien):
    """Kembalikan teks untuk kolom TTD (baris dipisah \\n), kosong bila tidak ada Lab PK / Lab PA / Resep."""
    try:
        semua = sorted(f for f in os.listdir(folder_pasien) if f.lower().endswith(".pdf"))
    except OSError:
        return ""
    masalah, lengkap = [], []
    for label, pola, kata, rel, halaman in JENIS:
        files = [f for f in semua if pola.match(f)]
        for f in files:
            suf = f" [{f[:-4].split(' - ')[-1]}]" if len(files) > 1 else ""
            try:
                h = periksa_file(os.path.join(folder_pasien, f), kata, rel, halaman)
            except Exception as e:
                masalah.append(f"CEK MANUAL: {label}{suf} tidak bisa dibaca ({type(e).__name__})")
                continue
            if h == "ada":
                if label not in lengkap:
                    lengkap.append(label)
            elif h == "tidak":
                masalah.append(f"{label}{suf}: TIDAK ADA TANDA TANGAN DOKTER")
            else:
                masalah.append(f"CEK MANUAL: {label}{suf} area tanda tangan tidak ditemukan")
    if lengkap and not masalah:
        return "TTD DOKTER ADA: " + ", ".join(lengkap)
    if lengkap:
        masalah.append("TTD ADA: " + ", ".join(lengkap))
    return "\n".join(masalah)
