# -*- coding: utf-8 -*-
"""
bukti_checker.py
Memeriksa kelengkapan BUKTI penunjang: item tindakan di Rincian Biaya (80 - RINCIAN TAGIHAN)
dicocokkan dengan file hasil/expertise di folder pasien (HASIL RADIOLOGI, HASIL LABORATORIUM PK,
UPLOAD DOKUMEN, dst).

Menambah item pemeriksaan baru nanti: cukup tambah satu entri di daftar KATEGORI di bawah.

    from bukti_checker import periksa_bukti
    h = periksa_bukti(folder_pasien)
    h["catatan"]  -> teks untuk kolom CATATAN KELENGKAPAN BUKTI (kosong bila lengkap)
    h["tanda"]    -> teks untuk kolom TANDA (cek manual / tidak ada tindakan) atau ""
    h["detail"]   -> list baris per item (untuk sheet Detail Bukti)
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

# ---------------------------------------------------------------------------------------------
# DAFTAR KATEGORI  (ubah / tambah di sini)
#   nama      : istilah di kolom catatan
#   unit_kw   : kata kunci pada NAMA UNIT di rincian (mis. "Radiodiagnostik")
#   item_kw   : kata kunci pada NAMA ITEM di rincian (mis. "Elektrokardiogram")
#   hasil_awalan : awalan jenis file hasil khusus (nama file: "40 - HASIL RADIOLOGI - ...")
#   file_kw   : kata kunci pada nama file (khususnya UPLOAD DOKUMEN yang ditulis manual)
#   sebut_item: False -> di sheet Detail Bukti cukup ditulis nama kategorinya (tanpa daftar item), mis. lab
#   per_unit  : True -> tiap UNIT layanan di rincian butuh 1 file hasil (mis. 2 unit lab PK = 2 file)
#   mode      : "per_item"  -> tiap item di rincian butuh 1 file
#               "per_kategori" -> cukup ada >= 1 file untuk seluruh kategori
# Item diperiksa lewat item_kw dulu, baru unit_kw.
# ---------------------------------------------------------------------------------------------
KATEGORI = [
    dict(nama="EKG", item_kw=["elektrokardiogram", "ecg", "ekg"], unit_kw=[],
         hasil_awalan=["HASIL EKG", "HASIL ECG"], file_kw=["ekg", "ecg", "elektrokardiogram"], mode="per_item"),
    dict(nama="ECHO", item_kw=["echo"], unit_kw=["echo"],
         hasil_awalan=["HASIL ECHO"], file_kw=["echo", "ekokardio"], mode="per_item"),
    dict(nama="USG", item_kw=["usg", "ultrasonografi", "ultrasound"], unit_kw=["usg", "ultrasonografi"],
         hasil_awalan=["HASIL USG"], file_kw=["usg", "ultrasonografi", "ultrasound"], mode="per_item"),
    dict(nama="RADIOLOGI", item_kw=[], unit_kw=["radiodiagnostik", "radiologi"],
         hasil_awalan=["HASIL RADIOLOGI"],
         file_kw=["radiologi", "rontgen", "röntgen", "thorax", "toraks", "thoraks", "ct scan", "ctscan", "mri"],
         mode="per_item"),
    dict(nama="SPIROMETRI", item_kw=["spirometri", "spirometer"], unit_kw=[],
         hasil_awalan=["HASIL SPIROMETRI"], file_kw=["spirometri", "spirometer", "spiro"], mode="per_item"),
    # Endoskopi: satu laporan hasil cukup untuk seluruh tindakan endoskopi di rincian
    dict(nama="ENDOSKOPI", item_kw=["esofagogastro", "endoskopi", "endoscopy", "egd", "kolonoskopi", "bronkoskopi"],
         unit_kw=["endoskopi", "endoscopy"],
         hasil_awalan=["HASIL ENDOSKOPI", "HASIL EGD"],
         file_kw=["egd", "endoskopi", "endoscopy", "kolonoskopi", "bronkoskopi", "esofagogastro"], mode="per_kategori"),
    dict(nama="LAB PK", item_kw=[], unit_kw=["patologi klinik", "laboratorium pk"],
         hasil_awalan=["HASIL LABORATORIUM PK"], file_kw=["lab pk", "laboratorium", "darah", "lab"], mode="per_kategori",
         sebut_item=False, per_unit=True),
    dict(nama="LAB PA", item_kw=[], unit_kw=["patologi anatomi", "laboratorium pa"],
         hasil_awalan=["HASIL LABORATORIUM PA", "HASIL PATOLOGI ANATOMI"],
         file_kw=["patologi anatomi", "lab pa", "laboratorium pa", "hasil pa", "histopatologi", "sitologi", "biopsi"], mode="per_kategori",
         sebut_item=False, per_unit=True),
    # ---- Radioterapi rawat jalan
    # Radiasi Eksternal 2D/3D/IMRT di rincian -> harus ada upload "Protokol Radiasi"
    dict(nama="PROTOKOL RADIASI", tampil="RADIASI EKSTERNAL", kata_kurang="BUKTI",
         item_kw=["radiasi eksternal", "radiasi external"], unit_kw=[],
         hasil_awalan=[], file_kw=["protokol radiasi"], mode="per_kategori", sebut_item=False),
    # CT Simulator di rincian -> harus ada bukti permintaan pasien CT Sim DAN posisi pasien CT Sim
    # (file_kw berisi tuple = SEMUA kata dalam tuple harus ada pada nama file)
    dict(nama="CT SIM (PERMINTAAN PASIEN)", tampil="CT SIMULATOR", kata_kurang="BUKTI",
         item_kw=["ct simulator", "ct simulasi", "ct sim"], unit_kw=[],
         hasil_awalan=[], file_kw=[("permintaan", "ct sim")], mode="per_kategori", sebut_item=False),
    dict(nama="CT SIM (POSISI PASIEN)", tampil="CT SIMULATOR", kata_kurang="BUKTI",
         item_kw=["ct simulator", "ct simulasi", "ct sim"], unit_kw=[],
         hasil_awalan=[], file_kw=[("posisi", "ct sim")], mode="per_kategori", sebut_item=False),
]

# Jenis file yang BUKAN bukti penunjang
BUKAN_BUKTI = ("SEP", "PENGKAJIAN HARIAN", "RESEP DOKTER", "RINCIAN TAGIHAN")

# Unit di rincian yang memang bukan tindakan (obat) -> seluruh isinya diabaikan
UNIT_DIABAIKAN = ("depo", "farmasi", "apotek")
# Item yang diabaikan (konsultasi/administrasi), selain yang sudah dibuang pembaca rincian
# File upload yang bukan hasil pemeriksaan (tidak dilaporkan sebagai "belum terpakai")
UPLOAD_DIABAIKAN = ("formulir", "persetujuan", "informed", "consent", "ktp", "kartu", "rujukan", "surat")
ITEM_DIABAIKAN = ("paket poli", "karcis", "administrasi", "pendaftaran", "materai")

AWALAN_RINCIAN = "80 - RINCIAN TAGIHAN"

_ANGKA = re.compile(r"^[\d.,]+$")
_QTY = re.compile(r"^\d+,\d{2}$")
_DOKTER = re.compile(r"^Pemeriksaan\s+(Klinik|Dokter)", re.I)


# ----------------------------------------------------------------------------- rincian biaya
def baca_rincian(path_pdf):
    """Kembalikan list (unit, item) dari satu PDF Rincian Biaya."""
    return baca_rincian_unit(path_pdf)[0]


def baca_rincian_unit(path_pdf):
    """Kembalikan (list (unit, item), jumlah unit layanan yang terbaca)."""
    with fitz.open(path_pdf) as d:
        teks = [l.strip() for p in d for l in p.get_text("text").splitlines() if l.strip()]
    units = {l[7:].strip() for l in teks if l.startswith("Jumlah ")}
    hasil, unit = [], None
    kepala = False   # True selama membaca kop halaman lanjutan (tabel yang terpotong antar halaman)
    for i, l in enumerate(teks):
        if l == "RINCIAN BIAYA PELAYANAN PASIEN":
            kepala = True
            continue
        if kepala:
            if l == "Total Tagihan":   # judul kolom terakhir pada kop tabel
                kepala = False
            continue
        if l in units:
            unit = l
            continue
        if l.startswith("Jumlah ") and l[7:].strip() == unit:
            unit = None
            continue
        if not unit:
            continue
        if _ANGKA.match(l) or l in ("Rp.", ": Rp.") or l.startswith("-") or _DOKTER.match(l):
            continue
        nxt = teks[i + 1] if i + 1 < len(teks) else ""
        if l.endswith(":") or _QTY.match(nxt):
            nama = l.rstrip(":").strip()
            if re.search(r"[A-Za-z0-9]", nama):   # buang baris tanda baca/kosong
                hasil.append((unit, nama))
    return hasil, len(units)


# Rincian wajib memuat minimal satu: paket poli / pemeriksaan dokter / klinik, spesialis atau subspesialis
_POLA_SPESIALIS = re.compile(r"(paket\s+poli|pemeriksaan\s+(?:dokter|klinik))\s+(sub\s*)?spesialis", re.I)
CATATAN_SPESIALIS = "RINCIAN TAGIHAN: TIDAK ADA PAKET POLI / PEMERIKSAAN DOKTER SPESIALIS/SUBSPESIALIS"


def ada_dokter_spesialis(daftar_pdf):
    """True bila salah satu PDF rincian memuat paket poli / pemeriksaan dokter (klinik) spesialis atau subspesialis."""
    for p in daftar_pdf:
        with fitz.open(p) as d:
            teks = " ".join(" ".join(pg.get_text("text").split()) for pg in d)
        if _POLA_SPESIALIS.search(teks):
            return True
    return False


def unit_spesialis(daftar_pdf):
    """{unit layanan: {"spes": ada paket/pemeriksaan spesialis, "sub": ada paket/pemeriksaan SUBspesialis}}"""
    hasil = {}
    for p in daftar_pdf:
        with fitz.open(p) as d:
            teks = [l.strip() for pg in d for l in pg.get_text("text").splitlines() if l.strip()]
        units = {l[7:].strip() for l in teks if l.startswith("Jumlah ")}
        unit, kepala = None, False
        for l in teks:
            if l == "RINCIAN BIAYA PELAYANAN PASIEN":
                kepala = True
                continue
            if kepala:
                kepala = l != "Total Tagihan"
                continue
            if l in units:
                unit = l
                hasil.setdefault(unit, dict(spes=False, sub=False))
                continue
            if l.startswith("Jumlah ") and l[7:].strip() == unit:
                unit = None
                continue
            m = _POLA_SPESIALIS.search(l) if unit else None
            if m:
                hasil[unit]["sub" if m.group(2) else "spes"] = True
    return hasil


_KATA_POLI = {"poliklinik", "poli", "klinik", "sub", "sp", "spesialis", "subspesialis", "dokter", "instalasi", "rawat", "jalan"}


def _kunci_poli(nama):
    return {t for t in _norm(nama).split() if t not in _KATA_POLI}


def _unit_poli(unit):
    return any(k in unit.lower() for k in ("poli", "klinik")) and not any(
        k in unit.lower() for k in ("laboratorium", "radio", "pelayanan"))


def catatan_poli_spesialis(daftar_pdf, klinik=""):
    """Teks catatan bila rincian TIDAK memuat paket poli/pemeriksaan dokter spesialis di poli kunjungan
    (atau di poli mana pun bila poli kunjungan tidak diketahui). '' bila lengkap."""
    unit = unit_spesialis(daftar_pdf)
    ada = [u for u, v in unit.items() if v["spes"] or v["sub"]]
    if not ada:
        return CATATAN_SPESIALIS
    kk = _kunci_poli(klinik)
    if not kk:
        return ""
    jiwa = "jiwa" in kk       # poli Jiwa: harus paket/pemeriksaan SPESIALIS (tertulis subspesialis = salah)
    for u in ada:
        ku = _kunci_poli(u)
        if ku and kk <= ku:       # semua kata poli kunjungan harus ada pada nama unit di rincian
            if jiwa and not unit[u]["spes"]:
                return (f"RINCIAN TAGIHAN: {u.upper()} HARUS PAKET POLI / PEMERIKSAAN DOKTER SPESIALIS, "
                        "TERTULIS SUBSPESIALIS")
            return ""
    return ("RINCIAN TAGIHAN: TIDAK ADA PAKET POLI / PEMERIKSAAN DOKTER SPESIALIS/SUBSPESIALIS DI "
            f"{klinik.upper()} (YANG ADA: {', '.join(u.upper() for u in ada)})")


def _pdf_kosong(path_pdf):
    """True bila PDF tidak punya teks sama sekali (halaman putih)."""
    try:
        with fitz.open(path_pdf) as d:
            return not any(p.get_text("text").strip() for p in d)
    except Exception:
        return False


def _kategori(unit, item):
    """Daftar kategori yang cocok untuk satu item rincian (umumnya satu; CT Sim butuh dua bukti)."""
    ul, il = unit.lower(), item.lower()
    cocok = [k for k in KATEGORI if any(re.search(r"\b" + re.escape(w), il) for w in k["item_kw"])]
    if cocok:
        return cocok
    return [k for k in KATEGORI if any(w in ul for w in k["unit_kw"])][:1]


# ----------------------------------------------------------------------------- file bukti
_POLA_FILE = re.compile(r"^(\d{2}) - (.+?) - (.*?)(?:_\d+)?$")


def _norm(s):
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def _kata_ada(teks, kw):
    """kw: daftar kata; elemen bertipe tuple/list = semua kata di dalamnya harus ada."""
    t = " " + _norm(teks) + " "

    def ada(w):
        return " " + _norm(w) + " " in t or (len(_norm(w)) > 3 and _norm(w) in t)
    return any(all(ada(x) for x in w) if isinstance(w, (tuple, list)) else ada(w) for w in kw)


def _daftar_file(folder):
    out = []
    for f in sorted(os.listdir(folder)):
        if not f.lower().endswith(".pdf"):
            continue
        nama = f[:-4]
        m = _POLA_FILE.match(nama)
        jenis, detail = (m.group(2), m.group(3)) if m else ("", nama)
        out.append(dict(file=f, jenis=jenis.strip(), detail=detail.strip()))
    return out


def _kandidat(kat, files):
    """File yang bisa menjadi bukti untuk satu kategori. Hasil khusus didahulukan."""
    khusus, upload = [], []
    for f in files:
        if f["jenis"].upper().startswith(BUKAN_BUKTI):
            continue
        if any(f["jenis"].upper().startswith(a) for a in kat["hasil_awalan"]):
            khusus.append(f)
        elif f["jenis"].upper().startswith(("UPLOAD", "HASIL")) and _kata_ada(f["detail"], kat["file_kw"]):
            upload.append(f)
    return khusus, upload


def _cocokkan_per_item(items, khusus, upload):
    """Pasangkan item -> file. Kembalikan (pasangan {idx: file|None}, tidak_pasti)."""
    kand = list(khusus) + list(upload)
    skor = []
    for i, it in enumerate(items):
        for j, f in enumerate(kand):
            skor.append((SequenceMatcher(None, _norm(it), _norm(f["detail"])).ratio(), i, j))
    skor.sort(reverse=True)
    pas, pakai = {}, set()
    for r, i, j in skor:
        if i in pas or j in pakai or r < 0.6:
            continue
        pas[i] = (kand[j], True)
        pakai.add(j)
    tidak_pasti = False
    # sisa item: pasangkan dengan sisa file (urut: file hasil khusus dulu) -> nama tidak cocok = ditandai
    sisa = [j for j in range(len(kand)) if j not in pakai]
    for i in range(len(items)):
        if i not in pas and sisa:
            j = sisa.pop(0)
            pas[i] = (kand[j], False)
            tidak_pasti = len(items) > 1   # 1 item + 1 file tidak ambigu
    return pas, tidak_pasti


# ----------------------------------------------------------------------------- utama
def periksa_bukti(folder_pasien, hd=False, klinik=""):
    """Periksa bukti penunjang satu folder pasien."""
    hasil = dict(catatan="", tanda="", detail=[], rincian_ada=False)
    try:
        files = _daftar_file(folder_pasien)
    except OSError as e:
        hasil["tanda"] = f"CEK MANUAL: FOLDER TIDAK BISA DIBACA ({type(e).__name__})"
        return hasil

    rincian = [f for f in files if f["file"].upper().startswith(AWALAN_RINCIAN)]
    if not rincian:
        hasil["tanda"] = "CEK MANUAL: FILE RINCIAN TAGIHAN TIDAK ADA"
        return hasil

    baris, n_unit = [], 0
    for r in rincian:
        try:
            b_, u_ = baca_rincian_unit(os.path.join(folder_pasien, r["file"]))
            baris += b_
            n_unit += u_
        except Exception as e:
            hasil["tanda"] = f"CEK MANUAL: RINCIAN TAGIHAN TIDAK TERBACA ({type(e).__name__})"
            return hasil
    hasil["rincian_ada"] = True
    if n_unit == 0:
        if all(_pdf_kosong(os.path.join(folder_pasien, r["file"])) for r in rincian):
            # halaman putih: tagihan belum diproses kasir / pasien belum selesai dilayani -> bukti belum bisa dinilai
            hasil["tanda"] = ("CEK MANUAL: RINCIAN TAGIHAN KOSONG (tagihan belum terbit / pasien belum selesai), "
                              "bukti penunjang belum bisa dinilai")
        else:   # ada teks, tetapi tidak satu pun unit layanan terbaca -> format rincian tidak dikenali
            hasil["tanda"] = "CEK MANUAL: ITEM RINCIAN TAGIHAN TIDAK TERBACA"
        return hasil
    tanda_poli = []
    # aturan: rincian harus memuat paket poli / pemeriksaan dokter spesialis (pasien HD dikecualikan)
    if not hd and not any("hemodialisis" in u.lower() for u, _ in baris):
        try:
            hasil["catatan"] = catatan_poli_spesialis([os.path.join(folder_pasien, r["file"]) for r in rincian], klinik)
            if not _kunci_poli(klinik):
                tanda_poli.append("poli kunjungan (Klinik Tujuan) tidak terbaca di resume, "
                                  "kecocokan poli pada rincian tagihan belum dicek")
        except Exception:
            pass
    if not baris:     # unit terbaca tetapi hanya konsultasi dokter -> tidak ada tindakan
        if tanda_poli:
            hasil["tanda"] = "CEK MANUAL: " + "; ".join(tanda_poli)
        return hasil

    # kelompokkan per kategori
    per_kat, tak_dikenal = {}, []
    for unit, item in baris:
        kk = _kategori(unit, item)
        if kk:
            for k in kk:
                per_kat.setdefault(k["nama"], (k, []))[1].append((unit, item))
        elif not any(u in unit.lower() for u in UNIT_DIABAIKAN + ("hemodialisis",)) \
                and not any(w in item.lower() for w in ITEM_DIABAIKAN):
            teks_item = f"{item} ({unit})"
            if teks_item not in tak_dikenal:
                tak_dikenal.append(teks_item)

    kurang, tanda, dipakai = [], list(tanda_poli), set()
    for nama, (k, daftar) in per_kat.items():
        khusus, upload = _kandidat(k, files)
        if k["mode"] == "per_kategori":
            unit_daftar = list(dict.fromkeys(u for u, _ in daftar))
            perlu = len(unit_daftar) if k.get("per_unit") else 1      # jumlah file hasil yang dibutuhkan
            file_ada = khusus + upload
            ada = len(file_ada) >= perlu if file_ada else False
            for f in file_ada[:perlu]:
                dipakai.add(f["file"])
            teks_item = "; ".join(it for _, it in daftar) if k.get("sebut_item", True) else nama
            if perlu > 1:
                teks_item = f"{nama} ({len(file_ada)} file dari {perlu} unit: {'; '.join(unit_daftar)})"
            hasil["detail"].append([nama, daftar[0][0], teks_item, "ADA" if ada else "BELUM ADA",
                                    file_ada[0]["file"] if file_ada else ""])
            if not ada:
                kurang.append((k.get("kata_kurang", "HASIL/EXPERTISE"),
                               f"{nama} ({len(file_ada)} DARI {perlu} UNIT)" if perlu > 1 else nama))
            continue
        items = [it for _, it in daftar]
        pas, tidak_pasti = _cocokkan_per_item(items, khusus, upload)
        belum = []
        for i, (unit, it) in enumerate(daftar):
            f = pas.get(i)
            if f:
                dipakai.add(f[0]["file"])
            hasil["detail"].append([nama, unit, it, "ADA" if f else "BELUM ADA", f[0]["file"] if f else ""])
            if not f:
                belum.append(it)
        if tidak_pasti:
            tanda.append(f"nama file {nama} tidak persis sama dengan item rincian")
        if belum:
            kurang.append((k.get("kata_kurang", "HASIL/EXPERTISE"),
                           f"RADIOLOGI ({', '.join(belum)})" if nama == "RADIOLOGI" else nama))

    if kurang:
        kalimat = []
        for kata in dict.fromkeys(kt for kt, _ in kurang):      # urut kemunculan: HASIL/EXPERTISE dulu, lalu BUKTI
            kalimat.append(f"BELUM ADA {kata} " + ", ".join(n for kt, n in kurang if kt == kata))
        hasil["catatan"] = "\n".join(x for x in [hasil["catatan"]] + kalimat if x)

    # file upload yang tidak terpakai oleh kategori manapun: kemungkinan nama upload ditulis lain
    asing = [f["file"] for f in files
             if f["jenis"].upper().startswith("UPLOAD") and f["file"] not in dipakai
             and not any(_kata_ada(f["detail"], k["file_kw"]) for k, _ in per_kat.values())
             and not any(w in f["detail"].lower() for w in UPLOAD_DIABAIKAN)]

    # ---- isi kolom TANDA: tindakan yang ada ditulis walau sudah lengkap
    bagian = []
    if per_kat:
        ringkas = ", ".join(dict.fromkeys(
            k.get("tampil") or (f"{n} ({len(d)})" if k["mode"] == "per_item" and len(d) > 1 else n)
            for n, (k, d) in per_kat.items()))
        bagian.append("ADA TINDAKAN: " + ringkas)
    if tanda:
        bagian.append("CEK MANUAL: " + "; ".join(tanda))
    if tak_dikenal:
        bagian.append("ITEM BELUM DIKENAL (cek buktinya manual): " + "; ".join(tak_dikenal))
    if asing and (kurang or tak_dikenal):
        bagian.append("file upload belum terpakai: " + "; ".join(a[:-4] for a in asing))
    hasil["tanda"] = "\n".join(bagian)   # kosong bila tidak ada tindakan sama sekali
    return hasil
