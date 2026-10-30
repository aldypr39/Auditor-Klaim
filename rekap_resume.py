# -*- coding: utf-8 -*-
"""
rekap_resume.py
Menelusuri folder hasil unduhan (sesi / {NoRM}_{Nama} / *.pdf), memeriksa Resume Medis
(folder 11 - PENGKAJIAN HARIAN) dengan resume_checker, lalu menulis Excel
"Rekap_Kelengkapan_Resume.xlsx" dengan kolom yang sama seperti rekap manual:

  TANGGAL BEROBAT | NO RM | NAMA | SUDAH DI KODE | KELENGKAPAN | DPJP |
  CATATAN KELENGKAPAN RESUME | CATATAN KELENGKAPAN BUKTI

Sheet kedua "Detail Item" berisi status Terisi/Kosong tiap item per pasien.

Dipanggil dari analyzer.py:
    import rekap_resume
    rekap_resume.buat_rekap(folder_utama, self.callback_log, self.callback_progress)
"""
import os
import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

from resume_checker import periksa_resume, ITEM as ITEM_SEMUA
from bukti_checker import periksa_bukti
from hd_checker import periksa_hd
from identitas_checker import periksa_identitas
from ttd_checker import periksa_ttd

# Tidak ikut menentukan LENGKAP/TIDAK LENGKAP dan tidak masuk catatan:
#   diagnosis sekunder (boleh ada / tidak ada) dan ICD 10 utama (diisi/dikoreksi petugas koding)
ABAIKAN = {"dx_sekunder", "icd_sekunder", "icd_utama"}
ITEM = [(k, l) for k, l in ITEM_SEMUA if k not in ABAIKAN]

# Awalan nama file hasil unduhan scraper untuk Resume Medis
AWALAN_FILE = "11 - PENGKAJIAN HARIAN"

# Nama item di kolom "CATATAN KELENGKAPAN RESUME". Ubah di sini kalau ingin istilah lain.
NAMA_CATATAN = {
    "keluhan_kep": "KELUHAN UTAMA (PERAWAT)",
    "tanda_vital": "TANDA VITAL",
    "nutrisi": "NUTRISI",
    "petugas": "PETUGAS",
    "nyeri": "PENILAIAN NYERI",
    "riwayat_penyakit": "RIWAYAT PENYAKIT (PERAWAT)",
    "edukasi": "KEBUTUHAN EDUKASI",
    "riwayat_obat": "RIWAYAT PENGOBATAN",
    "diagnosis_a": "DIAGNOSIS (A)",
    "s_anamnesis": "ANAMNESIS / RPS",
    "s_keluhan": "KELUHAN UTAMA (S)",
    "s_rpd": "RIWAYAT PENYAKIT TERDAHULU",
    "obyektif": "OBYEKTIF (O)",
    "terapi_p": "RENCANA TERAPI (P)",
    "tindakan": "TINDAKAN",
    "prognosis": "PROGNOSIS",
    "dx_utama": "DIAGNOSIS UTAMA",
    "icd_utama": "ICD 10 UTAMA",
    "dx_sekunder": "DIAGNOSIS SEKUNDER",
    "icd_sekunder": "ICD 10 SEKUNDER",
    "kondisi_pulang": "KONDISI SAAT MENINGGALKAN FASILITAS",
    "ttd_dokter": "TANDA TANGAN DOKTER",
}

HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
BULAN = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]

KOLOM = ["NO", "TANGGAL BEROBAT", "NO RM", "NAMA", "SUDAH DI KODE", "KELENGKAPAN", "DPJP",
         "CATATAN KELENGKAPAN RESUME", "CATATAN KELENGKAPAN BUKTI", "TANDA", "DOKUMEN GAGAL UNDUH", "FOLDER",
         "TTD DOKTER (RESEP, LAB PK, LAB PA)"]


POLA_URUT = re.compile(r"^(\d{4}) - (.*)$")
POLA_TGL = re.compile(r" \[(\d{2}-\d{2}-\d{4}(?: \d{4})?)\]$")


def _punya_pdf(path):
    try:
        return any(f.lower().endswith(".pdf") for f in os.listdir(path))
    except OSError:
        return False


def _subfolder(path):
    try:
        return sorted(d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d)))
    except OSError:
        return []


def daftar_pasien(folder_utama):
    """Kembalikan list pasien berurutan sesuai nomor urut SIMRS (awalan 'NNNN - ').

    Folder yang dipilih boleh:
      - folder induk (berisi folder sesi), atau
      - folder sesi langsung (berisi folder pasien).
    Folder lama tanpa nomor urut ditaruh paling akhir, urut nama.
    """
    anak = _subfolder(folder_utama)
    if any(_punya_pdf(os.path.join(folder_utama, d)) for d in anak):
        sesi_list = [(os.path.basename(os.path.normpath(folder_utama)), folder_utama)]
    else:
        sesi_list = [(d, os.path.join(folder_utama, d)) for d in anak]

    hasil = []
    for sesi, ps in sesi_list:
        pas = []
        for d in _subfolder(ps):
            nm, urut = d, None
            m = POLA_URUT.match(d)
            if m:
                urut, nm = int(m.group(1)), m.group(2)
            tgl = ""
            mt = POLA_TGL.search(nm)          # penanda kunjungan: ' [24-10-2026 1042]'
            if mt:
                tgl, nm = mt.group(1), nm[:mt.start()]
            b = nm.split("_", 1)
            pas.append(dict(sesi=sesi, path_sesi=ps, nama_folder=d, path=os.path.join(ps, d), urutan=urut, tgl=tgl,
                            no_rm=b[0], nama=b[1] if len(b) > 1 else ""))
        pas.sort(key=lambda p: (p["urutan"] is None, p["urutan"] or 0, p["nama_folder"]))
        hasil.extend(pas)
    return hasil


def tanggal_indonesia(teks):
    """'03-10-2026 08:26:07' -> 'Sabtu, 03 Oktober 2026'"""
    m = re.match(r"(\d{2})-(\d{2})-(\d{4})", teks or "")
    if not m:
        return ""
    d, mo, y = (int(x) for x in m.groups())
    try:
        dt = datetime(y, mo, d)
    except ValueError:
        return teks
    return f"{HARI[dt.weekday()]}, {d:02d} {BULAN[mo]} {y}"


NAMA_GAGAL = "_gagal_unduh.txt"


def _baca_baris(path):
    try:
        with open(path, encoding="utf-8") as f:
            return [l.strip() for l in f if l.strip()]
    except OSError:
        return []


def daftar_gagal_unduh(pekerjaan):
    """Kembalikan {path folder pasien: "dokumen A; dokumen B"} dari catatan gagal unduh scraper.

    Sumber 1 (utama): <folder pasien>/_gagal_unduh.txt   (baris: "GRUP | NAMA DOKUMEN")
    Sumber 2 (cadangan, data lama): <folder sesi>/_gagal_unduh.txt  (baris: "NORM_NAMA | GRUP | NAMA DOKUMEN").
      Dipakai hanya bila pasien itu satu-satunya kunjungan dengan RM+nama yang sama di sesi tersebut.
    """
    from collections import Counter

    def kunci_nama(t):
        return re.sub(r"[^a-z0-9]+", "", (t or "").lower())

    jumlah = Counter((p["sesi"], p["no_rm"], p["nama"]) for p in pekerjaan)
    cache_sesi = {}
    hasil = {}
    for p in pekerjaan:
        item = []
        for l in _baca_baris(os.path.join(p["path"], NAMA_GAGAL)):
            bag = [x.strip() for x in l.split(" | ")]
            item.append(" - ".join(bag[-2:]) if len(bag) >= 2 else bag[0])
        if not item:
            # catatan lama (gabungan satu sesi, tanpa penanda kunjungan): dicocokkan lewat RM + nama
            if p["path_sesi"] not in cache_sesi:
                cache_sesi[p["path_sesi"]] = _baca_baris(os.path.join(p["path_sesi"], NAMA_GAGAL))
            kunci = kunci_nama(f"{p['no_rm']}_{p['nama']}")
            ragu = jumlah[(p["sesi"], p["no_rm"], p["nama"])] > 1     # pasien punya >1 kunjungan di sesi ini
            for l in cache_sesi[p["path_sesi"]]:
                bag = [x.strip() for x in l.split(" | ")]
                if len(bag) >= 3 and kunci_nama(bag[0]) == kunci:
                    item.append(f"{bag[1]} - {bag[2]}" + (" (catatan lama, kunjungan tidak pasti)" if ragu else ""))
        unik = list(dict.fromkeys(item))
        if unik:
            hasil[p["path"]] = "\n".join(unik)
    return hasil


def _cari_resume(folder_pasien):
    """Semua PDF resume (nama berawalan 11 - PENGKAJIAN HARIAN) di folder pasien."""
    try:
        semua = sorted(os.listdir(folder_pasien))
    except OSError:
        return []
    return [f for f in semua
            if f.lower().endswith(".pdf") and f.upper().startswith(AWALAN_FILE.upper())]


def _baris_dari_hasil(sesi, no_rm, nama, nama_file, h):
    # Aturan: bila ICD 10 utama terisi, diagnosis utama dianggap terisi walau kolom nama kosong
    semua = dict(ITEM_SEMUA)      # icd_utama tidak ada di ITEM (diabaikan), tetapi tetap dibaca untuk aturan ini
    if h["item"][semua["icd_utama"]]["terisi"]:
        h["item"][semua["dx_utama"]]["terisi"] = True
    tidak = [k for k, _ in ITEM if not h["item"][dict(ITEM)[k]]["terisi"]]

    def nama_catatan(k):
        # tanda vital & nutrisi: sebutkan bagian mana yang kosong, mis. "TANDA VITAL (SUHU, RR)"
        n = NAMA_CATATAN[k]
        if k in ("keluhan_kep", "s_keluhan") and "kontrol" in h["item"][dict(ITEM)[k]].get("keterangan", ""):
            return n + ": ISINYA KONTROL (BUKAN KOSONG)"
        if k == "obyektif":
            m = re.match(r"kosong:\s*(.+)", h["item"][dict(ITEM)[k]].get("keterangan", ""))
            if m:
                return f"{n}: {m.group(1).upper()} TIDAK ADA - Harus diisi di kolom yang tersedia"
            return f"{n}: Harus diisi di kolom yang tersedia"
        if k in ("tanda_vital", "nutrisi"):
            m = re.match(r"kosong:\s*(.+)", h["item"][dict(ITEM)[k]].get("keterangan", ""))
            if m:
                n += f" ({m.group(1).upper()})"
        return n
    catatan = [nama_catatan(k) for k in tidak]
    info = h["info"]
    if h["catatan"]:
        # hanya peringatan yang membuat hasil tidak bisa dipercaya (bukan sekadar posisi cadangan)
        catatan_tambahan = [c for c in h["catatan"] if "judul 'resume medis'" in c.lower()
                            or "tidak punya teks" in c.lower()]
    else:
        catatan_tambahan = []
    # item yang terisi tetapi perlu dicek manual (mis. keluhan "kontrol" dengan keterangan sedikit)
    cek_resume = []
    for k, _ in ITEM:
        ket = h["item"][dict(ITEM)[k]]
        if ket["terisi"] and ket.get("keterangan", "").startswith("cek manual"):
            cek_resume.append(f"CEK MANUAL: {NAMA_CATATAN[k]} berisi kata kontrol dengan keterangan sedikit, "
                              "pastikan keterangannya memadai")
    return dict(cek_resume=cek_resume,
        sesi=sesi, nama_file=nama_file,
        tanggal=tanggal_indonesia(info.get("tanggal_masuk")),
        no_rm=no_rm, nama=nama,
        kode="",   # diisi manual oleh petugas (SUDAH / BELUM)
        lengkap="LENGKAP" if not tidak and not catatan_tambahan else "TIDAK LENGKAP",
        dpjp=info.get("dpjp", ""),
        catatan=", ".join(catatan + [c.upper() for c in catatan_tambahan]),
        catatan_list=catatan + [c.upper() for c in catatan_tambahan],
        terisi=len(ITEM) - len(tidak), total=len(ITEM), hasil=h,
    )


def kumpulkan(folder_utama, log=print, progress=None):
    """Kembalikan list baris (satu per Resume Medis)."""
    pekerjaan = daftar_pasien(folder_utama)

    baris = []
    total = len(pekerjaan)
    gagal_map = daftar_gagal_unduh(pekerjaan)
    ttd_map = {}        # {nama folder pasien: teks kolom TTD DOKTER}
    for i, pk in enumerate(pekerjaan, 1):
        if progress:
            progress(i / max(total, 1))
        sesi, pp, no_rm, nama, urut = pk["sesi"], pk["path"], pk["no_rm"], pk["nama"], pk["urutan"]
        file_resume = _cari_resume(pp)
        hasil_resume = []      # (nama file, hasil, error)
        for f in file_resume:
            try:
                hasil_resume.append((f, periksa_resume(os.path.join(pp, f)), None))
            except Exception as e:  # satu PDF rusak tidak boleh menghentikan semuanya
                hasil_resume.append((f, None, e))
        info_ref = next((h["info"] for _, h, _ in hasil_resume if h), {})
        try:
            bukti = periksa_bukti(pp, hd=("hemodialisis" in sesi.lower()), klinik=info_ref.get("klinik", ""))
        except Exception as e:   # pengecekan bukti tidak boleh menghentikan rekap resume
            bukti = dict(catatan="", tanda=f"CEK MANUAL: PEMERIKSAAN BUKTI GAGAL ({type(e).__name__})", detail=[])
        # Pasien Hemodialisis (sesi HD) + pencocokan obat rincian <-> resep: digabung ke hasil bukti
        try:
            hd = periksa_hd(pp, sesi, pk["tgl"])
            if hd["catatan"]:
                bukti["catatan"] = (bukti["catatan"] + "\n" if bukti["catatan"] else "") + hd["catatan"]
            if hd["tanda"]:
                bukti["tanda"] = (bukti["tanda"] + "\n" if bukti["tanda"] else "") + hd["tanda"]
            bukti["detail"] = list(bukti["detail"]) + hd["detail"]
        except Exception as e:
            bukti["tanda"] = (bukti["tanda"] + "\n" if bukti["tanda"] else "") + \
                f"CEK MANUAL: PEMERIKSAAN HD GAGAL ({type(e).__name__})"
        # SEP / Resep / Rincian: tanggal pendaftaran & nama harus sama dengan Resume Medis
        # (cadangan bila resume tidak ada/tidak terbaca: penanda tanggal & nama pada nama folder)
        try:
            tgl_ref = (info_ref.get("tanggal_masuk") or "")[:10] or pk["tgl"][:10]
            nama_ref = info_ref.get("nama") or nama
            idn = periksa_identitas(pp, tgl_ref, nama_ref)
            if idn["catatan"]:
                bukti["catatan"] = (bukti["catatan"] + "\n" if bukti["catatan"] else "") + idn["catatan"]
            if idn["tanda"]:
                bukti["tanda"] = (bukti["tanda"] + "\n" if bukti["tanda"] else "") + idn["tanda"]
        except Exception as e:
            bukti["tanda"] = (bukti["tanda"] + "\n" if bukti["tanda"] else "") + \
                f"CEK MANUAL: PEMERIKSAAN TANGGAL/NAMA DOKUMEN GAGAL ({type(e).__name__})"
        # Bukti kurang + ada dokumen gagal diunduh -> kemungkinan filenya ada di SIMRS, perlu dicek manual
        if (bukti["catatan"] or "RINCIAN TAGIHAN TIDAK ADA" in bukti["tanda"]) and gagal_map.get(pk["path"]):
            bukti["tanda"] = (bukti["tanda"] + "\n" if bukti["tanda"] else "") + \
                "CEK MANUAL: ada dokumen gagal diunduh, yang kurang mungkin ada di SIMRS (lihat kolom DOKUMEN GAGAL UNDUH)"
        try:        # tanda tangan dokter pada Resep / Lab PK / Lab PA: kolom khusus, tidak memengaruhi status
            ttd_map[pk["nama_folder"]] = periksa_ttd(pp)
        except Exception as e:
            ttd_map[pk["nama_folder"]] = f"CEK MANUAL: pemeriksaan tanda tangan gagal ({type(e).__name__})"
        n0 = len(baris)
        if bukti["catatan"]:
            log(f"🧾 Bukti {nama}: {_satu_baris(bukti['catatan'])}")
        if bukti["tanda"]:
            log(f"🔍 Bukti {nama}: {_satu_baris(bukti['tanda'])}")
        if not bukti["catatan"] and not bukti["tanda"]:
            log(f"🧾 Bukti {nama}: lengkap" if bukti["detail"] else f"🧾 Bukti {nama}: tidak ada tindakan penunjang")
        if not file_resume:
            baris.append(dict(urut=urut, sesi=sesi, nama_file="", tanggal=tanggal_indonesia(pk["tgl"]), no_rm=no_rm, nama=nama,
                              kode="", lengkap="TIDAK LENGKAP", dpjp="",
                              catatan="FILE RESUME MEDIS TIDAK ADA", terisi=0, total=len(ITEM), hasil=None))
            log(f"⚠️ {nama}: file Resume Medis tidak ada")
            baris[-1]["bukti"], baris[-1]["tanda"], baris[-1]["detail_bukti"] = bukti["catatan"], bukti["tanda"], bukti["detail"]
            baris[-1]["folder"] = pk["nama_folder"]
            baris[-1]["gagal"] = gagal_map.get(pk["path"], "")
            continue
        for f, h, e in hasil_resume:
            if e is not None:
                baris.append(dict(urut=urut, sesi=sesi, nama_file=f, tanggal=tanggal_indonesia(pk["tgl"]), no_rm=no_rm, nama=nama,
                                  kode="", lengkap="TIDAK LENGKAP", dpjp="",
                                  catatan=f"PDF TIDAK BISA DIBACA ({type(e).__name__})",
                                  terisi=0, total=len(ITEM), hasil=None))
                log(f"❌ {nama}: {f} gagal dibaca: {e}")
                baris[-1]["bukti"], baris[-1]["tanda"], baris[-1]["detail_bukti"] = bukti["catatan"], bukti["tanda"], bukti["detail"]
                baris[-1]["folder"] = pk["nama_folder"]
                baris[-1]["gagal"] = gagal_map.get(pk["path"], "")
                continue
            b = _baris_dari_hasil(sesi, no_rm, nama, f, h)
            b["urut"] = urut
            if len(file_resume) > 1:
                b["catatan"] = (b["catatan"] + ", " if b["catatan"] else "") + f"ADA {len(file_resume)} FILE RESUME ({f})"
                b["catatan_list"] = list(b.get("catatan_list") or []) + [f"ADA {len(file_resume)} FILE RESUME ({f})"]
            baris.append(b)
            log(f"✔️ Resume {nama}: {b['terisi']}/{b['total']} item terisi")
        for b in baris[n0:]:
            if not b.get("tanggal") and pk["tgl"]:      # tanggal cadangan dari penanda folder
                b["tanggal"] = tanggal_indonesia(pk["tgl"])
            extra = b.get("cek_resume") or []
            b["bukti"], b["detail_bukti"] = bukti["catatan"], bukti["detail"]
            b["tanda"] = "\n".join(x for x in [bukti["tanda"]] + extra if x)
            b["folder"] = pk["nama_folder"]
            b["gagal"] = gagal_map.get(pk["path"], "")
    for b in baris:
        b["ttd"] = ttd_map.get(b.get("folder"), "")
    # Ada hasil/expertise yang belum ada -> status kelengkapan otomatis TIDAK LENGKAP
    # (tanda "CEK MANUAL" saja tidak mengubah status)
    for b in baris:
        if b.get("bukti"):
            b["lengkap"] = "TIDAK LENGKAP"
    # ringkasan akhir per pasien (sama dengan kolom KELENGKAPAN di Excel)
    log("\n===== HASIL AKHIR =====")
    for b in baris:
        if b["lengkap"] == "LENGKAP":
            log(f"✅ {b['nama']} -> LENGKAP")
        else:
            alasan = ", ".join(x for x in (b.get("catatan", ""), _satu_baris(b.get("bukti", ""))) if x)
            log(f"❌ {b['nama']} -> TIDAK LENGKAP: {alasan}")
    return baris


def _satu_baris(teks):
    """Teks bertingkat (satu butir per baris) -> satu baris untuk log."""
    return " | ".join(x for x in (teks or "").split("\n") if x)


def _poin(teks):
    """Teks berbaris-baris -> daftar butir bertanda • (satu butir per baris dalam sel Excel)."""
    butir = [x.strip() for x in (teks or "").split("\n") if x.strip()]
    return "\n".join("• " + x for x in butir)


def _rel(b):
    """Jalur folder pasien relatif terhadap Excel (Excel disimpan di folder sesi), siap dipakai di HYPERLINK."""
    return (b.get("folder") or "").replace('"', '""')


def _link_folder(b):
    """Sel 'buka folder pasien': ikon folder bila folder diketahui."""
    return f'=HYPERLINK("{_rel(b)}","📂 Buka")' if b.get("folder") else ""


def _link_nama(b):
    nama = (b["nama"] or "").replace('"', '""')
    return f'=HYPERLINK("{_rel(b)}","{nama}")' if b.get("folder") else b["nama"]


def tulis_excel(baris, path_excel):
    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap Resume"

    ungu = PatternFill("solid", fgColor="5B3A8C")
    merah = PatternFill("solid", fgColor="F8D7DA")
    hijau = PatternFill("solid", fgColor="D4EDDA")
    merah_muda = PatternFill("solid", fgColor="FDECEA")
    kuning = PatternFill("solid", fgColor="FFF8E1")
    garis = Side(style="thin", color="BBBBBB")
    kotak = Border(left=garis, right=garis, top=garis, bottom=garis)

    ws.append(KOLOM)
    for c in ws[1]:
        c.fill, c.font = ungu, Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = kotak
    ws.row_dimensions[1].height = 32

    for b in baris:
        ws.append([b.get("urut"), b["tanggal"], b["no_rm"], b["nama"], b["kode"], b["lengkap"],
                   b["dpjp"], "\n".join("• " + x for x in b["catatan_list"]) if b.get("catatan_list") else b["catatan"],
                   _poin(b.get("bukti", "")), _poin(b.get("tanda", "")), _poin(b.get("gagal", "")), _link_folder(b), _poin(b.get("ttd", ""))])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = kotak
            c.alignment = Alignment(vertical="top", wrap_text=True)
        row[11].font, row[11].alignment = Font(color="0563C1", underline="single"), Alignment(horizontal="center", vertical="top")
        row[4].fill = hijau if row[4].value == "SUDAH" else (merah if row[4].value == "BELUM" else PatternFill())
        row[5].fill = hijau if row[5].value == "LENGKAP" else merah
        # warna lembut per jenis catatan: merah muda = kekurangan, kuning = perlu dicek manual
        if row[7].value:
            row[7].fill = merah_muda
        if row[8].value:
            row[8].fill = merah_muda
        if row[9].value:
            row[9].fill = kuning
        if row[10].value:
            row[10].fill = kuning
        if row[12].value:
            row[12].fill = hijau if str(row[12].value).startswith("• TTD DOKTER ADA") else kuning

    for i, w in enumerate([7, 24, 12, 28, 16, 16, 30, 48, 58, 58, 40, 10, 40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "E2"
    if baris:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(KOLOM))}{len(baris) + 1}"
        dv1 = DataValidation(type="list", formula1='"SUDAH,BELUM"', allow_blank=True)
        dv2 = DataValidation(type="list", formula1='"LENGKAP,TIDAK LENGKAP"', allow_blank=True)
        ws.add_data_validation(dv1)
        ws.add_data_validation(dv2)
        dv1.add(f"E2:E{len(baris) + 1}")
        dv2.add(f"F2:F{len(baris) + 1}")

    # ---- sheet detail: satu kolom per item
    wd = wb.create_sheet("Detail Item")
    kepala = ["NO URUT", "TANGGAL BEROBAT", "SESI", "NO RM", "NAMA", "FILE", "TERISI"] + [NAMA_CATATAN[k] for k, _ in ITEM]
    wd.append(kepala)
    for c in wd[1]:
        c.fill, c.font = ungu, Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wd.row_dimensions[1].height = 48
    for b in baris:
        h = b["hasil"]
        if h:
            sel = ["Terisi" if h["item"][lab]["terisi"]
                   else "Kosong" + (f" ({h['item'][lab]['keterangan']})" if h["item"][lab].get("keterangan") else "")
                   for _, lab in ITEM]
        else:
            sel = ["-"] * len(ITEM)
        wd.append([b.get("urut"), b["tanggal"], b["sesi"], b["no_rm"], _link_nama(b), b["nama_file"], f"{b['terisi']}/{b['total']}"] + sel)
    for row in wd.iter_rows(min_row=2, min_col=8):
        for c in row:
            c.alignment = Alignment(horizontal="center", wrap_text=True)
            if str(c.value).startswith("Kosong"):
                c.fill = merah
    for i in range(1, len(kepala) + 1):
        wd.column_dimensions[get_column_letter(i)].width = 14 if i > 7 else [9, 24, 22, 12, 28, 40, 9][i - 1]
    for row in wd.iter_rows(min_row=2, min_col=5, max_col=5):
        row[0].font = Font(color="0563C1", underline="single")
    wd.freeze_panes = "H2"

    # ---- sheet detail bukti: satu baris per item penunjang di rincian biaya
    wb_ = wb.create_sheet("Detail Bukti")
    wb_.append(["NO URUT", "TANGGAL BEROBAT", "NO RM", "NAMA", "KATEGORI", "UNIT DI RINCIAN TAGIHAN", "ITEM DI RINCIAN TAGIHAN", "STATUS", "FILE BUKTI"])
    for c in wb_[1]:
        c.fill, c.font = ungu, Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    biru = Font(color="0563C1", underline="single")
    for b in baris:
        for kat, unit, item, status, f in b.get("detail_bukti", []):
            wb_.append([b.get("urut"), b["tanggal"], b["no_rm"], _link_nama(b), kat, unit, item, status, f])
            if f and b.get("folder"):
                # tautan relatif terhadap lokasi Excel (Excel disimpan di folder sesi, sejajar dengan folder pasien)
                rel = (b["folder"] + "\\" + f).replace('"', '""')
                c = wb_.cell(row=wb_.max_row, column=9)
                c.value = f'=HYPERLINK("{rel}","{f.replace(chr(34), chr(34) * 2)}")'
                c.font = biru
    for row in wb_.iter_rows(min_row=2):
        if row[7].value == "BELUM ADA":
            for c in row:
                c.fill = merah
    for i, w in enumerate([9, 24, 12, 28, 14, 34, 50, 12, 60], 1):
        wb_.column_dimensions[get_column_letter(i)].width = w
    for row in wb_.iter_rows(min_row=2, min_col=4, max_col=4):
        row[0].font = Font(color="0563C1", underline="single")
    wb_.freeze_panes = "A2"

    wb.save(path_excel)


def buat_rekap(folder_utama, log=print, progress=None):
    log("Memeriksa kelengkapan Resume Medis...")
    baris = kumpulkan(folder_utama, log, progress)
    if not baris:
        log("⚠️ Tidak ada folder pasien untuk diperiksa.")
        return None
    # Satu file per folder sesi, disimpan SATU FOLDER dengan folder-folder pasien
    peta = {p["sesi"]: p["path_sesi"] for p in daftar_pasien(folder_utama)}
    per_sesi = {}
    for b in baris:
        per_sesi.setdefault(b["sesi"], []).append(b)
    path = None
    for sesi, rows in per_sesi.items():
        tujuan = peta.get(sesi, folder_utama)
        m = re.search(r"(\d{2}-\d{2}-\d{4})\s*sd\s*(\d{2}-\d{2}-\d{4})", sesi)
        periode = f"_{m.group(1)} sd {m.group(2)}" if m else ""
        mp = re.match(r"\s*\[(.+?)\]", sesi)          # nama poli dari nama folder sesi "[Poli] tgl sd tgl"
        poli = "_" + re.sub(r'[\\/*?:"<>|]', "-", mp.group(1)).strip() if mp else ""
        path = os.path.join(tujuan, f"Rekap_Kelengkapan_Resume{poli}{periode}.xlsx")
        tulis_excel(rows, path)
        lengkap = sum(1 for b in rows if b["lengkap"] == "LENGKAP")
        log(f"✅ Rekap Resume dibuat: {path}  ({lengkap} lengkap dari {len(rows)})")
    return path
