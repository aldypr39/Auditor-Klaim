# -*- coding: utf-8 -*-
"""
hd_checker.py
Pemeriksaan khusus pasien HEMODIALISIS (HD) + pencocokan obat rincian <-> resep.

    from hd_checker import periksa_hd
    h = periksa_hd(folder_pasien, sesi, tgl)
    h["catatan"] -> masuk kolom CATATAN KELENGKAPAN BUKTI (bila ada = status TIDAK LENGKAP)
    h["tanda"]   -> masuk kolom TANDA
    h["detail"]  -> baris untuk sheet Detail Bukti  [kategori, unit, item, status, file]

Aturan HD:
  1. Pasien HD = nama sesi mengandung "hemodialisis" (cadangan: rincian punya unit Hemodialisis).
  2. Wajib ada file "98 - PROTOKOL HD" (boleh >1 file/lembar).
  3. Pada baris protokol dengan tanggal = tanggal kunjungan, TTD Kepala Ruangan HD dan TTD DPJP wajib ada.
  4. Rincian wajib memuat tindakan HD (item apa pun di unit Hemodialisis).
  5. Epodion di rincian harus ada di resep (cukup 1 file resep); Epodion di resep tapi tidak di rincian -> TANDA.
"""
import os
import re

import warnings
warnings.filterwarnings("ignore")
try:
    import pymupdf as fitz
except ImportError:  # pragma: no cover
    import fitz

from bukti_checker import baca_rincian_unit, _daftar_file, _norm, AWALAN_RINCIAN, UNIT_DIABAIKAN

# ----------------------------------------------------------------------------- pengaturan
AWALAN_PROTOKOL = "98 - PROTOKOL HD"
AWALAN_RESEP = "50 - RESEP DOKTER"
KATA_SESI_HD = "hemodialisis"
KATA_UNIT_HD = "hemodialisis"

# Obat yang wajib ada di resep bila tercantum di rincian.
OBAT_WAJIB = ("epodion",)

# ---- DISIAPKAN, BELUM AKTIF -------------------------------------------------------------------
# Aturan sebenarnya: SEMUA obat di rincian harus ada di resep (dan sebaliknya), untuk semua poli.
# Untuk mengaktifkannya nanti: ubah  OBAT_SEMUA_POLI = True  (lihat juga fungsi _nama_obat di bawah).
OBAT_SEMUA_POLI = False
# -------------------------------------------------------------------------------------------------

_TGL = re.compile(r"^\d{2}-\d{2}-\d{4}$")
_JAM = re.compile(r"^\d{2}:\d{2}")


# ----------------------------------------------------------------------------- umum
def sesi_hd(sesi):
    return KATA_SESI_HD in (sesi or "").lower()


def tanggal_kunjungan(tgl_folder, rincian_files):
    """'dd-mm-yyyy' dari penanda folder; cadangan: 'Tgl. Masuk' di rincian (dd/mm/yyyy)."""
    m = re.match(r"^(\d{2}-\d{2}-\d{4})", tgl_folder or "")
    if m:
        return m.group(1)
    for p in rincian_files:
        try:
            with fitz.open(p) as d:
                teks = "\n".join(pg.get_text("text") for pg in d)
        except Exception:
            continue
        m = re.search(r"(\d{2})/(\d{2})/(\d{4})\s+\d{2}\.\d{2}\.\d{2}", teks)
        if m:
            return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return ""


# ----------------------------------------------------------------------------- protokol HD
def _ada_gambar_ttd(page, x0, x1, y0, y1):
    """True bila ada gambar tanda tangan (bukan placeholder 1x1) di dalam sel."""
    for im in page.get_image_info():
        b = im["bbox"]
        if im.get("width", 0) <= 1 or im.get("height", 0) <= 1:
            continue
        if (b[2] - b[0]) < 15 or (b[3] - b[1]) < 10:
            continue
        cx, cy = (b[0] + b[2]) / 2, (b[1] + b[3]) / 2
        if x0 <= cx < x1 and y0 <= cy < y1:
            return True
    # cadangan: tanda tangan berupa goresan vektor
    for dr in page.get_drawings():
        r = dr["rect"]
        if r.width < 8 or r.height < 8 or r.width > (x1 - x0) - 4 or r.height > (y1 - y0) - 4:
            continue
        if len(dr["items"]) < 3 or all(it[0] == "re" for it in dr["items"]):
            continue
        cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
        if x0 <= cx < x1 and y0 <= cy < y1:
            return True
    return False


def baca_protokol(path_pdf):
    """List baris protokol: dict(tgl, kr(bool), dpjp(bool)) dari semua halaman."""
    hasil = []
    with fitz.open(path_pdf) as d:
        for page in d:
            words = page.get_text("words")
            # pusat kolom dari judul tabel; cadangan: posisi standar
            kr_c, dp_c = 457.0, 530.0
            for w in words:
                if w[4] == "Kepala":
                    kr_c = (w[0] + w[2]) / 2 + 0.0
                    kr_c = max(kr_c, 440)   # judul 3 kata: "Kepala" ada di tengah-kiri
                if w[4] == "DPJP":
                    dp_c = (w[0] + w[2]) / 2
            dp_c = dp_c if 480 < dp_c < 560 else 530.0
            kr_c = 457.0 if not (430 < kr_c < 480) else kr_c
            batas = (kr_c + dp_c) / 2
            kiri, kanan = kr_c - 45, dp_c + 45

            # baris tanggal: dd-mm-yyyy di kolom paling kiri + jam tepat di bawahnya
            baris = []
            for w in words:
                if w[0] < 110 and _TGL.match(w[4]):
                    ada_jam = any(_JAM.match(v[4]) and abs(v[0] - w[0]) < 30 and 3 < v[1] - w[1] < 20 for v in words)
                    if ada_jam:
                        baris.append((w[1], w[4]))
            baris.sort()
            for i, (y, tgl) in enumerate(baris):
                atas = y - 20
                bawah = (baris[i + 1][0] - 20) if i + 1 < len(baris) else y + 44
                hasil.append(dict(
                    tgl=tgl,
                    kr=_ada_gambar_ttd(page, kiri, batas, atas, bawah),
                    dpjp=_ada_gambar_ttd(page, batas, kanan, atas, bawah)))
    return hasil


def _periksa_protokol(folder, files, tgl, hasil):
    prot = [f for f in files if f["file"].upper().startswith(AWALAN_PROTOKOL)]
    if not prot:
        hasil["catatan"].append("PROTOKOL HD TIDAK ADA")
        hasil["detail"].append(["HEMODIALISIS", "Protokol HD", "Formulir Protokol HD", "BELUM ADA", ""])
        return
    baris, terbaca = [], 0
    for f in prot:
        try:
            baris += baca_protokol(os.path.join(folder, f["file"]))
            terbaca += 1
        except Exception as e:
            hasil["tanda"].append(f"CEK MANUAL: {f['file']} tidak terbaca ({type(e).__name__})")
    nama_file = prot[0]["file"]
    if not terbaca:
        return
    if not tgl:
        hasil["tanda"].append("CEK MANUAL: tanggal kunjungan tidak diketahui, TTD Protokol HD belum bisa dicek")
        return
    if not baris:
        hasil["tanda"].append("CEK MANUAL: baris tanggal Protokol HD tidak terbaca (format berubah / hasil scan)")
        return
    cocok = [b for b in baris if b["tgl"] == tgl]
    if not cocok:
        hasil["catatan"].append(f"PROTOKOL HD {tgl}: TANGGAL TIDAK ADA DI PROTOKOL")
        hasil["detail"].append(["HEMODIALISIS", "Protokol HD", f"Tanggal {tgl}", "BELUM ADA", nama_file])
        return
    if any(b["kr"] and b["dpjp"] for b in cocok):
        hasil["detail"].append(["HEMODIALISIS", "Protokol HD", f"TTD Kepala Ruangan & DPJP {tgl}", "ADA", nama_file])
        return
    kr = any(b["kr"] for b in cocok)
    dp = any(b["dpjp"] for b in cocok)
    if not kr and not dp:
        ket = "TTD KEPALA RUANGAN DAN TTD DPJP TIDAK ADA"
    elif not kr:
        ket = "TTD KEPALA RUANGAN TIDAK ADA"
    else:
        ket = "TTD DPJP TIDAK ADA"
    hasil["catatan"].append(f"PROTOKOL HD {tgl}: {ket}")
    hasil["detail"].append(["HEMODIALISIS", "Protokol HD", f"{ket} ({tgl})", "BELUM ADA", nama_file])


# ----------------------------------------------------------------------------- obat: rincian vs resep
def _nama_obat(item):
    """Nama obat tanpa kekuatan/sediaan: 'Epodion 3000IU/0.3ml inj' -> 'epodion'.
    (dipakai bila OBAT_SEMUA_POLI diaktifkan)"""
    m = re.match(r"^\s*([A-Za-z][A-Za-z\- ]*?)\s*(?:\d|$)", item or "")
    return _norm(m.group(1)) if m else _norm(item)


def _teks_resep(folder, files):
    """(teks gabungan semua file resep, jumlah file, ada_yang_tak_terbaca)"""
    res = [f for f in files if f["file"].upper().startswith(AWALAN_RESEP)]
    teks, kosong = [], False
    for f in res:
        try:
            with fitz.open(os.path.join(folder, f["file"])) as d:
                t = "\n".join(p.get_text("text") for p in d)
        except Exception:
            kosong = True
            continue
        if not t.strip():
            kosong = True
        teks.append(t)
    return _norm(" ".join(teks)), res, kosong


def _periksa_obat(folder, files, baris_rincian, hasil):
    obat_rincian = [it for u, it in baris_rincian if any(x in u.lower() for x in UNIT_DIABAIKAN)]
    if OBAT_SEMUA_POLI:      # --- belum aktif: semua obat di rincian wajib ada di resep
        wajib = sorted({_nama_obat(o) for o in obat_rincian if _nama_obat(o)})
    else:
        wajib = [k for k in OBAT_WAJIB if any(k in _norm(o) for o in obat_rincian)]
    teks, res, kosong = _teks_resep(folder, files)

    for k in wajib:
        if not res:
            hasil["catatan"].append(f"{k.upper()} ADA DI RINCIAN TAGIHAN TAPI FILE RESEP TIDAK ADA")
            hasil["detail"].append(["OBAT", "Resep", k.capitalize(), "BELUM ADA", ""])
        elif f" {k} " in f" {teks} ":
            hasil["detail"].append(["OBAT", "Resep", k.capitalize(), "ADA", res[0]["file"]])
        elif kosong:
            hasil["tanda"].append(f"CEK MANUAL: resep tidak terbaca (scan/kosong), {k.capitalize()} belum bisa dicocokkan")
        else:
            hasil["catatan"].append(f"{k.upper()} ADA DI RINCIAN TAGIHAN TAPI TIDAK ADA DI RESEP")
            hasil["detail"].append(["OBAT", "Resep", k.capitalize(), "BELUM ADA", res[0]["file"]])

    # kebalikannya (hanya daftar OBAT_WAJIB): ada di resep, tidak ada di rincian -> TANDA saja
    for k in OBAT_WAJIB:
        if f" {k} " in f" {teks} " and not any(k in _norm(o) for o in obat_rincian):
            hasil["tanda"].append(f"{k.upper()} ADA DI RESEP TAPI TIDAK ADA DI RINCIAN TAGIHAN")


# ----------------------------------------------------------------------------- utama
def periksa_hd(folder_pasien, sesi, tgl=""):
    """Periksa pasien HD. Bukan sesi HD (dan rincian tanpa unit HD) -> hasil kosong."""
    hasil = dict(catatan=[], tanda=[], detail=[], hd=False)
    try:
        files = _daftar_file(folder_pasien)
    except OSError:
        return dict(catatan="", tanda="", detail=[], hd=False)

    rincian = [f for f in files if f["file"].upper().startswith(AWALAN_RINCIAN)]
    baris, n_unit = [], 0
    for r in rincian:
        try:
            b_, u_ = baca_rincian_unit(os.path.join(folder_pasien, r["file"]))
            baris += b_
            n_unit += u_
        except Exception:
            pass
    hd = sesi_hd(sesi) or any(KATA_UNIT_HD in u.lower() for u, _ in baris)
    if not hd and not OBAT_SEMUA_POLI:
        return dict(catatan="", tanda="", detail=[], hd=False)
    if hd:
        hasil["hd"] = True
        tgl = tanggal_kunjungan(tgl, [os.path.join(folder_pasien, r["file"]) for r in rincian])
        _periksa_protokol(folder_pasien, files, tgl, hasil)
        if n_unit:   # rincian terbaca -> cek tindakan HD
            tind = [it for u, it in baris if KATA_UNIT_HD in u.lower()]
            if tind:
                hasil["detail"].append(["HEMODIALISIS", "Hemodialisis", "; ".join(tind), "ADA", rincian[0]["file"]])
            else:
                hasil["catatan"].append("TINDAKAN HD TIDAK ADA DI RINCIAN TAGIHAN")
                hasil["detail"].append(["HEMODIALISIS", "Hemodialisis", "Tindakan HD", "BELUM ADA", rincian[0]["file"]])
    if n_unit:
        _periksa_obat(folder_pasien, files, baris, hasil)
    return dict(catatan="\n".join(hasil["catatan"]), tanda="\n".join(hasil["tanda"]),
                detail=hasil["detail"], hd=hasil["hd"])
