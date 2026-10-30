# -*- coding: utf-8 -*-
"""
identitas_checker.py
Mencocokkan TANGGAL PENDAFTARAN dan NAMA PASIEN (acuan: Resume Medis) dengan dokumen lain:

  SEP     (00 - SEP*)             : "Tgl. SEP" dan "Nama Peserta"
  RESEP   (50 - RESEP DOKTER*)    : "Tgl Pendaftaran" (cadangan: "Tgl" di kepala resep)      -> tanggal saja
  RINCIAN (80 - RINCIAN TAGIHAN*) : "Tgl. Masuk" dan "Nama"

    from identitas_checker import periksa_identitas
    h = periksa_identitas(folder_pasien, tgl_acuan, nama_acuan)
    h["catatan"] -> kolom CATATAN KELENGKAPAN BUKTI (status TIDAK LENGKAP)
    h["tanda"]   -> kolom TANDA

Aturan bila satu jenis dokumen punya lebih dari satu file:
  - semua file tidak cocok  -> CATATAN (tidak lengkap)
  - sebagian saja tidak cocok -> TANDA (CEK MANUAL)
Pasien bayi (nama memuat "BY"/"BAYI", mis. "BAYI NYONYA X") dikecualikan dari pengecekan nama.
Nama mirip tapi tidak persis (selisih huruf / kurang satu kata) -> TANDA saja.
"""
import os
import re
from difflib import SequenceMatcher

import warnings
warnings.filterwarnings("ignore")
try:
    import pymupdf as fitz
except ImportError:  # pragma: no cover
    import fitz

# Awalan nama file
AWALAN_SEP = "00 - SEP"
AWALAN_RESEP = "50 - RESEP DOKTER"
AWALAN_RINCIAN = "80 - RINCIAN TAGIHAN"

# Sebutan di depan/belakang nama yang diabaikan saat mencocokkan
SEBUTAN = set("""tn ny nn an by bayi sdr sdri tuan nyonya nona anak mr mrs ms bpk bapak ibu h hj haji hajjah
dr drg ir prof""".split())
# Gelar akademik/profesi: dibuang di posisi mana pun (mis. "SITI AMINAH S KEP NS")
GELAR = set("""dr drg ns kep skep skm sked skp amd amk amkg ak akp sh se st sp spd spog spa spb spjp mkes mkm mm mpd mh msi
apt sarjana ahli madya""".split())
MIN_KATA_SAMA = 3   # satu kata yang sama (min. 3 huruf) sudah cukup dianggap cocok
# Penanda nama ayah (bin/binti): nama di belakangnya tidak dibandingkan, mis. "ITUN BINTI SYUKUR" = "ITUN"
PENANDA_AYAH = {"bin", "binti", "bte", "bt", "ibnu"}
KEMIRIPAN_NAMA = 0.85   # di atas ini = nama dianggap mirip (CEK MANUAL), di bawahnya = tidak sesuai

_TGL_ISO = re.compile(r"(\d{4})[-/](\d{2})[-/](\d{2})")
_TGL_STRIP = re.compile(r"(\d{2})[-/](\d{2})[-/](\d{4})")


# ----------------------------------------------------------------------------- util
def _tgl(teks):
    """'15-10-2026' / '27/10/2026 07.31.30' -> '15-10-2026'; tidak ada -> ''."""
    m = _TGL_STRIP.search(teks or "")
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = _TGL_ISO.search(teks or "")      # format tahun-dulu, mis. '2026-10-01 08:19:09' (Rincian Tagihan)
    return f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else ""


def _token_nama(nama):
    t = re.sub(r"[^a-z0-9]+", " ", (nama or "").lower()).split()
    for i, k in enumerate(t):                 # potong di "bin"/"binti" (bukan di kata pertama)
        if i > 0 and k in PENANDA_AYAH:
            t = t[:i]
            break
    while t and t[0] in SEBUTAN:
        t = t[1:]
    while t and t[-1] in SEBUTAN:
        t = t[:-1]
    bersih = [k for k in t if k not in GELAR]
    return bersih or t


def bandingkan_nama(a, b):
    """'sama' | 'mirip' | 'beda' | 'tidak_ada' (salah satu kosong)."""
    ta, tb = _token_nama(a), _token_nama(b)
    if not ta or not tb:
        return "tidak_ada"
    if ta == tb or sorted(ta) == sorted(tb):
        return "sama"
    if len(ta) == len(tb) and all(x == y or (len(x) == 1 and y.startswith(x)) or (len(y) == 1 and x.startswith(y))
                                  for x, y in zip(ta, tb)):
        return "sama"     # singkatan huruf depan, mis. "M SYAIHAN NOOR" = "MOHAMMAD SYAIHAN NOOR"
    sa, sb = set(ta), set(tb)
    if any(len(k) >= MIN_KATA_SAMA for k in sa & sb):
        return "sama"         # toleransi gelar / nama panjang: cukup ada satu kata yang sama
    if (sa <= sb or sb <= sa) and min(len(sa), len(sb)) >= 2:
        return "mirip"        # salah satu kekurangan kata (mis. nama belakang tidak tercetak)
    r = SequenceMatcher(None, " ".join(ta), " ".join(tb)).ratio()
    return "mirip" if r >= KEMIRIPAN_NAMA else "beda"


def _bayi(nama):
    """Nama bayi: 'BAYI NYONYA X', 'BY NY X', 'X, BY' -> nama tidak dicek (nama bayi di dokumen sering beda)."""
    t = re.sub(r"[^a-z0-9]+", " ", (nama or "").lower()).split()
    return any(k in ("by", "bayi", "baby") for k in t)


def _daftar(folder, awalan):
    try:
        return sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf") and f.upper().startswith(awalan))
    except OSError:
        return []


def _nilai_label(page, label, xmin, xmax, y_toleransi=4.0):
    """Teks di sebelah kanan sebuah label (urutan kata persis, mis. ['Tgl.', 'SEP']) pada baris yang sama.
    Hanya kata dengan xmin <= x0 < xmax yang diambil. Kosong bila label tidak ketemu."""
    words = page.get_text("words")
    n = len(label)
    for i in range(len(words) - n + 1):
        if all(words[i + j][4].lower().rstrip(":") == label[j].lower() for j in range(n)):
            w0 = words[i]
            yc = (w0[1] + w0[3]) / 2
            x_akhir = words[i + n - 1][2]
            nilai = [w for w in words
                     if abs((w[1] + w[3]) / 2 - yc) <= y_toleransi and w[0] >= max(xmin, x_akhir - 1) and w[0] < xmax
                     and w[4] not in (":",)]
            nilai.sort(key=lambda w: w[0])
            return " ".join(w[4].lstrip(":") for w in nilai).strip()
    return ""


def _halaman1(path):
    with fitz.open(path) as d:
        return d[0] if len(d) else None, d


# ----------------------------------------------------------------------------- pembaca per dokumen
def baca_sep(path):
    with fitz.open(path) as d:
        pg = d[0]
        return dict(tgl=_tgl(_nilai_label(pg, ["Tgl.", "SEP"], 60, 310)),
                    nama=_nilai_label(pg, ["Nama", "Peserta"], 60, 310))


def baca_resep(path):
    with fitz.open(path) as d:
        pg = d[0]
        t = _tgl(_nilai_label(pg, ["Tgl", "Pendaftaran"], 200, 600))
        if not t:        # cadangan: tanggal di kepala resep
            t = _tgl(_nilai_label(pg, ["Tgl"], 0, 200))
        return dict(tgl=t, nama="")


def baca_rincian_id(path):
    with fitz.open(path) as d:
        pg = d[0]
        return dict(tgl=_tgl(_nilai_label(pg, ["Tgl.", "Masuk"], 340, 600)),
                    nama=_nilai_label(pg, ["Nama"], 60, 340))


# ----------------------------------------------------------------------------- utama
def _periksa_jenis(folder, awalan, label, baca, tgl_ref, nama_ref, cek_nama, hasil):
    files = _daftar(folder, awalan)
    if not files:
        return
    ok_tgl = ok_nama = 0
    salah_tgl, salah_nama, tak_terbaca, mirip = [], [], [], []
    for f in files:
        try:
            v = baca(os.path.join(folder, f))
        except Exception:
            tak_terbaca.append(f"{label} {f[:-4]}")
            continue
        suf = f" [{f[:-4].split(' - ')[-1]}]" if len(files) > 1 else ""
        # --- tanggal
        if not v["tgl"]:
            tak_terbaca.append(f"{label}{suf} tanggal")
        elif tgl_ref and v["tgl"] != tgl_ref:
            salah_tgl.append((suf, v["tgl"]))
        else:
            ok_tgl += 1
        # --- nama
        if cek_nama and nama_ref and not (_bayi(nama_ref) or _bayi(v["nama"])):
            k = bandingkan_nama(v["nama"], nama_ref)
            if k == "tidak_ada":
                tak_terbaca.append(f"{label}{suf} nama")
            elif k == "beda":
                salah_nama.append((suf, v["nama"]))
            elif k == "mirip":
                mirip.append(f"{label}{suf}: nama '{v['nama']}' mirip dengan '{nama_ref}'")
                ok_nama += 1
            else:
                ok_nama += 1

    if salah_tgl:
        ket = ", ".join(f"{label}{s} {t}" for s, t in salah_tgl)
        if ok_tgl:     # sebagian cocok -> cukup tanda
            hasil["tanda"].append(f"CEK MANUAL: tanggal tidak sesuai pendaftaran {tgl_ref} pada {ket}")
        else:
            hasil["catatan"].append(f"{label}: TANGGAL TIDAK SESUAI PENDAFTARAN ({label} "
                                    f"{', '.join(t for _, t in salah_tgl)}, DAFTAR {tgl_ref})")
    if salah_nama:      # nama TIDAK memengaruhi status lengkap/tidak: cukup ditandai
        hasil["tanda"].append("CEK MANUAL: nama tidak sesuai dengan Resume Medis (" + nama_ref + ") pada " +
                              ", ".join(f"{label}{s} ({n})" for s, n in salah_nama))
    hasil["tanda"] += [f"CEK MANUAL: {m}" for m in mirip]
    if tak_terbaca:
        hasil["tanda"].append("CEK MANUAL: tidak terbaca - " + "; ".join(tak_terbaca))


def periksa_identitas(folder_pasien, tgl_ref, nama_ref):
    """tgl_ref 'dd-mm-yyyy' (tanggal pendaftaran), nama_ref nama pasien (dari Resume Medis)."""
    hasil = dict(catatan=[], tanda=[])
    tgl_ref = _tgl(tgl_ref)
    if not tgl_ref and not nama_ref:
        return dict(catatan="", tanda="")
    for awalan, label, baca, cek_nama in (
            (AWALAN_SEP, "SEP", baca_sep, True),
            (AWALAN_RESEP, "RESEP", baca_resep, False),
            (AWALAN_RINCIAN, "RINCIAN TAGIHAN", baca_rincian_id, True)):
        try:
            _periksa_jenis(folder_pasien, awalan, label, baca, tgl_ref, nama_ref, cek_nama, hasil)
        except Exception as e:      # satu dokumen bermasalah tidak boleh menghentikan rekap
            hasil["tanda"].append(f"CEK MANUAL: pemeriksaan {label} gagal ({type(e).__name__})")
    return dict(catatan="\n".join(hasil["catatan"]), tanda="\n".join(hasil["tanda"]))
