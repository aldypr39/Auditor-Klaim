# -*- coding: utf-8 -*-
"""
resume_checker.py
Pemeriksa kuantitatif (terisi / tidak terisi) untuk PDF "Resume Medis Rawat Jalan"
(RS Ulin, folder 11 - PENGKAJIAN HARIAN).

Pemakaian:
    from resume_checker import periksa_resume
    hasil = periksa_resume("path/ke/resume.pdf")
    hasil["item"]["TINDAKAN"]["terisi"]      # True / False
    hasil["tidak_terisi"]                     # ["TINDAKAN", "PROGNOSIS", ...]

CLI:
    python resume_checker.py file.pdf            # ringkasan
    python resume_checker.py file.pdf --debug    # + isi tiap area yang dibaca
    python resume_checker.py file.pdf --json

Prinsip: PDF dibaca dengan koordinat (kata + garis tabel), bukan dengan mencari kata
di teks mentah, sebab label seperti "Keluhan Utama", "Jam", "Diagnosis" muncul lebih dari
sekali. Semua halaman digabung ke satu sumbu-Y global (halaman_ke * 1000 + y).
"""
import re
import sys
import json

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover
    import pymupdf as fitz

STEP = 1000.0                       # jarak sumbu-Y antar halaman
CEKLIS = ("√", "✓", "✔", "☑", "✅")  # glyph tanda centang

# ---------------------------------------------------------------- daftar 20 item
ITEM = [
    ("keluhan_kep", "KELUHAN UTAMA"),
    ("tanda_vital", "TANDA VITAL"),
    ("nutrisi", "NUTRISI"),
    ("petugas", "PETUGAS"),
    ("nyeri", "PENILAIAN NYERI"),
    ("riwayat_penyakit", "RIWAYAT PENYAKIT"),
    ("edukasi", "KEBUTUHAN EDUKASI"),
    ("riwayat_obat", "RIWAYAT PENGOBATAN"),
    ("diagnosis_a", "DIAGNOSIS (A)"),
    ("s_anamnesis", "S: ANAMNESIS / RPS"),
    ("s_keluhan", "S: KELUHAN UTAMA"),
    ("s_rpd", "S: RIWAYAT PENYAKIT TERDAHULU"),
    ("obyektif", "OBYEKTIF (O)"),
    ("terapi_p", "RENCANA TERAPI (P)"),
    ("tindakan", "TINDAKAN"),
    ("prognosis", "PROGNOSIS"),
    ("dx_utama", "DIAGNOSIS UTAMA"),
    ("icd_utama", "ICD 10"),
    ("dx_sekunder", "DIAGNOSIS SEKUNDER"),
    ("icd_sekunder", "ICD 10 DIAGNOSIS SEKUNDER"),
    ("kondisi_pulang", "KONDISI SAAT MENINGGALKAN FASILITAS"),
    ("ttd_dokter", "TANDA TANGAN DOKTER"),
]
LABEL = dict(ITEM)

SATUAN = re.compile(
    r"^(mmhg|x/menit|x/mnt|x/mnit|kg|kg/m2?|kg/m²|cm|m|[º°]?c|celcius|celsius|%)$", re.I)
ICD_RE = re.compile(r"^[A-Z]\d{2}(?:\.\d{1,3})?$")


# ---------------------------------------------------------------- util teks
# Keluhan utama yang memuat kata "kontrol": dianggap TIDAK lengkap bila hanya "kontrol" + kata pengisi saja
# (mis. "kontrol", "kontrol rutin", "kontrol poli paru", "konsul"). Bila masih ada keterangan lain
# (mis. "kontrol post op appendiks", "kontrol DM, ambil obat") -> lengkap, kata apa pun boleh.
# Keterangan sedikit (<= KETERANGAN_SEDIKIT kata) -> tetap lengkap, tetapi ditulis CEK MANUAL di kolom TANDA.
# "terkontrol"/"tidak terkontrol" tidak ikut kena (bukan kata utuh).
KELUHAN_TIDAK_VALID = re.compile(r"\b(kontrol|control)\b")
KETERANGAN_SEDIKIT = 2
# Kata pengisi (bukan keterangan). Tambah / kurangi di sini bila perlu.
KATA_PENGISI = set("""
kontrol control kontrolnya rutin ulang ulangan berkala lanjutan lanjut lagi rawat jalan rj pasien datang ingin mau
poli polik poliklinik klinik ke di dari untuk dan yang yg dgn dengan sesuai jadwal biasa saja hari ini nya
konsul konsultasi konsulan pro cek check up
dr dokter drg sp spesialis subspesialis sub umum
paru jantung interna penyakit dalam bedah saraf anak mata tht kulit kelamin kandungan obgyn gigi mulut jiwa
rehab rehabilitasi medik urologi orthopedi ortopedi onkologi gizi syaraf neurologi kardiologi pulmonologi
""".split())


def _kata_keterangan(s):
    """Kata bermakna pada keluhan setelah kata pengisi (kontrol, poli, nama poli, dst.) dibuang."""
    s = re.sub(r"\b(KU|RPS|RPD|RPK)\s*:", " ", s or "", flags=re.I)
    kata = re.sub(r"[^a-z0-9]+", " ", s.lower()).split()
    return [k for k in kata if k not in KATA_PENGISI]


def _keluhan_valid(s):
    """False bila keluhan memuat 'kontrol' tetapi tidak ada keterangan lain selain kata pengisi."""
    t = re.sub(r"[^a-z0-9]+", " ", (s or "").lower())
    if not KELUHAN_TIDAK_VALID.search(t):
        return True
    return len(_kata_keterangan(s)) > 0


def _keluhan_ragu(s):
    """True bila memuat 'kontrol', dianggap lengkap, tetapi keterangannya sedikit (perlu dicek manual)."""
    t = re.sub(r"[^a-z0-9]+", " ", (s or "").lower())
    if not KELUHAN_TIDAK_VALID.search(t):
        return False
    n = len(_kata_keterangan(s))
    return 0 < n <= KETERANGAN_SEDIKIT


def _isi(s):
    """True bila teks benar-benar berisi (huruf/angka).  '-', ':', 'KU:' saja = kosong.
    'tidak ada', '0', '0.00' = terisi."""
    s = re.sub(r"\b(KU|RPS|RPD|RPK)\s*:", "", s or "", flags=re.I)
    return bool(re.search(r"[A-Za-z0-9]", s))


def _gabung(ws):
    return " ".join(w["t"] for w in ws)


def _ringkas(s, n=70):
    s = re.sub(r"\s+", " ", s or "").strip()
    return s if len(s) <= n else s[: n - 1] + "…"


# ---------------------------------------------------------------- pembaca dokumen
class _Dok:
    def __init__(self, sumber):
        self.doc = sumber if isinstance(sumber, fitz.Document) else fitz.open(sumber)
        self.W = []          # semua kata
        self.imgs = []       # (x0,y0,x1,y1) global
        self.vekt = []       # bentuk vektor non-garis (kandidat tanda tangan): (x0,y0,x1,y1)
        self.garis_tipis = []  # (x0,y0,x1,y1) garis lurus (untuk dikecualikan dari "vektor")
        Hraw, self.V = {}, []
        for pi, pg in enumerate(self.doc):
            off = pi * STEP
            for w in pg.get_text("words"):
                x0, y0, x1, y1, t = w[:5]
                self.W.append(dict(x0=x0, x1=x1, y0=y0 + off, y1=y1 + off,
                                   yc=(y0 + y1) / 2 + off, t=t, pg=pi))
            for im in pg.get_image_info():
                b = im["bbox"]
                self.imgs.append((b[0], b[1] + off, b[2], b[3] + off))
            for dr in pg.get_drawings():
                for it in dr["items"]:
                    op = it[0]
                    if op == "l":
                        self._seg(it[1].x, it[1].y + off, it[2].x, it[2].y + off, Hraw)
                    elif op == "re":
                        r = it[1]
                        self._rect(r.x0, r.y0 + off, r.x1, r.y1 + off, Hraw)
                    elif op == "qu":
                        r = it[1].rect
                        self._rect(r.x0, r.y0 + off, r.x1, r.y1 + off, Hraw)
                    elif op == "c":  # kurva = goresan tangan / tanda tangan vektor
                        xs = [it[k].x for k in range(1, 5)]
                        ys = [it[k].y + off for k in range(1, 5)]
                        self.vekt.append((min(xs), min(ys), max(xs), max(ys)))
        self.W.sort(key=lambda w: (w["yc"], w["x0"]))
        self.baris = self._klaster(self.W)
        self._susun_garis(Hraw)

    # ---- garis
    def _seg(self, x0, y0, x1, y1, H):
        if abs(y0 - y1) < 1.0 and abs(x1 - x0) >= 1:
            H.setdefault(round((y0 + y1) / 2), []).append((min(x0, x1), max(x0, x1)))
            self.garis_tipis.append((min(x0, x1), y0, max(x0, x1), y1))
        elif abs(x0 - x1) < 1.0 and abs(y1 - y0) >= 1:
            self.V.append(((x0 + x1) / 2, min(y0, y1), max(y0, y1)))
            self.garis_tipis.append((x0, min(y0, y1), x1, max(y0, y1)))

    def _rect(self, x0, y0, x1, y1, H):
        w, h = abs(x1 - x0), abs(y1 - y0)
        if h < 1.5:
            self._seg(x0, (y0 + y1) / 2, x1, (y0 + y1) / 2, H)
        elif w < 1.5:
            self._seg((x0 + x1) / 2, y0, (x0 + x1) / 2, y1, H)
        else:
            self._seg(x0, y0, x1, y0, H)
            self._seg(x0, y1, x1, y1, H)
            self._seg(x0, y0, x0, y1, H)
            self._seg(x1, y0, x1, y1, H)

    def _susun_garis(self, Hraw):
        self.HS = []   # (y, x0, x1) setelah digabung
        for y in sorted(Hraw):
            iv = sorted(Hraw[y])
            m = []
            for a, b in iv:
                if m and a <= m[-1][1] + 1.5:
                    m[-1][1] = max(m[-1][1], b)
                else:
                    m.append([a, b])
            for a, b in m:
                self.HS.append((y, a, b))
        # garis selebar tabel (hampir selebar halaman)
        ys = sorted({y for y, a, b in self.HS if a <= 75 and b >= 550})
        self.FH = []
        for y in ys:
            if not self.FH or y - self.FH[-1] > 1.5:
                self.FH.append(y)

    def fh_atas(self, g):
        """garis lebar terdekat DI ATAS/sama dengan g"""
        c = [y for y in self.FH if y <= g]
        return c[-1] if c else None

    def fh_bawah(self, g):
        """garis lebar terdekat DI BAWAH g"""
        c = [y for y in self.FH if y > g]
        return c[0] if c else None

    # ---- kata & baris
    @staticmethod
    def _klaster(words):
        ws = sorted(words, key=lambda w: (w["yc"], w["x0"]))
        out, cur, ref = [], [], 0
        for w in ws:
            if cur and abs(w["yc"] - ref) > 3.5:
                out.append(sorted(cur, key=lambda k: k["x0"]))
                cur = []
            cur.append(w)
            ref = sum(k["yc"] for k in cur) / len(cur)
        if cur:
            out.append(sorted(cur, key=lambda k: k["x0"]))
        return out

    def area(self, xa, xb, ga, gb):
        return [w for w in self.W
                if xa <= (w["x0"] + w["x1"]) / 2 < xb and ga <= w["yc"] < gb]

    def teks(self, ws):
        return "\n".join(_gabung(l) for l in self._klaster(ws))

    def cari(self, token, x_min=-1, x_max=9999, awal=False, setelah=None, sebelum=None):
        """cari urutan kata (regex per kata) dalam satu baris. awal=True: harus kata pertama baris."""
        for ln in self.baris:
            g = ln[0]["yc"]
            if setelah is not None and g <= setelah:
                continue
            if sebelum is not None and g >= sebelum:
                continue
            for i in range(len(ln)):
                if awal and i > 0:
                    break
                if not (x_min <= ln[i]["x0"] <= x_max):
                    continue
                if i + len(token) > len(ln):
                    break
                if all(re.fullmatch(token[k], ln[i + k]["t"], re.I) for k in range(len(token))):
                    m = ln[i:i + len(token)]
                    return dict(x0=m[0]["x0"], x1=m[-1]["x1"], y0=min(w["y0"] for w in m),
                                y1=max(w["y1"] for w in m), yc=m[0]["yc"], pg=m[0]["pg"],
                                words=m, baris=ln)
        return None


# ---------------------------------------------------------------- pemeriksa
class PemeriksaResume:
    def __init__(self, sumber, debug=False):
        self.d = _Dok(sumber)
        self.debug = debug
        self.log = []
        self.item = {}
        self.catatan = []
        self.info = {}

    def _dbg(self, judul, isi):
        if self.debug:
            self.log.append(f"--- {judul}\n{isi}")

    def _set(self, key, terisi, nilai="", ket=""):
        self.item[key] = dict(label=LABEL[key], terisi=bool(terisi),
                              nilai=_ringkas(nilai, 120), keterangan=ket)

    # -- pasangan label : nilai pada satu baris
    @staticmethod
    def _kv(ws, stop=None, buang_satuan=False):
        idx = next((i for i, w in enumerate(ws) if w["t"].startswith(":")), None)
        if idx is None:
            return _gabung(ws), "", False
        label = _gabung(ws[:idx])
        rest = [dict(w) for w in ws[idx:]]
        rest[0]["t"] = rest[0]["t"].lstrip(":")
        rest = [w for w in rest if w["t"]]
        if stop:
            for i, w in enumerate(rest):
                if re.fullmatch(stop, w["t"], re.I):
                    rest = rest[:i]
                    break
        if buang_satuan:
            rest = [w for w in rest if not SATUAN.match(w["t"])]
        return label, _gabung(rest), True

    # ============================================================== jalankan
    def jalankan(self):
        d = self.d
        if not d.W:
            self.catatan.append("PDF tidak punya teks (kemungkinan hasil scan/gambar) - perlu OCR/cek manual")
            for k, _ in ITEM:
                self._set(k, False, "", "tidak terbaca")
            return self._hasil()
        if not d.cari([r"RESUME", r"MEDIS"]):
            self.catatan.append("Judul 'RESUME MEDIS' tidak ditemukan - pastikan ini resume medis")
        if not d.FH:
            self.catatan.append("Garis tabel tidak terdeteksi - hasil Rencana Terapi/ICD perlu dicek manual")

        self._header()
        self._keperawatan()
        self._medis()
        self._icd()
        self._kondisi()
        self._ttd()
        return self._hasil()

    # ============================================================== header
    def _header(self):
        d = self.d
        np_ = d.cari([r"Nama", r"Pasien"], x_max=200)
        info = dict(tanggal_masuk="", no_rm="", nama="", umur=None, umur_teks="", klinik="")
        if np_:
            b = d.area(55, 227, np_["y1"], np_["y1"] + 25)
            info["nama"] = d.teks(b).replace("\n", " ")
            b = d.area(227, 406, np_["y1"], np_["y1"] + 25)
            t = d.teks(b).replace("\n", " ")
            info["umur_teks"] = t
            m = re.search(r"(\d+)\s*Tahun", t, re.I)
            if m:
                info["umur"] = int(m.group(1))
            elif re.search(r"Bulan|Hari|Minggu", t, re.I):
                info["umur"] = 0
        # klinik tujuan (poli kunjungan): teks tepat di bawah label "Klinik Tujuan"
        kt = d.cari([r"Klinik", r"Tujuan"], x_min=300)
        info["klinik"] = d.teks(d.area(kt["x0"] - 5, 9999, kt["y1"] + 1, kt["y1"] + 20)).replace("\n", " ").strip() if kt else ""
        p0 = " ".join(w["t"] for w in d.W if w["pg"] == 0)
        m = re.search(r"\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2}", p0)
        info["tanggal_masuk"] = m.group(0) if m else ""
        m = re.search(r"\b\d{2}-\d{2}-\d{2}-\d{2}\b", p0)
        info["no_rm"] = m.group(0) if m else ""
        self.info.update(info)

    # ============================================================== bagian I (keperawatan)
    def _keperawatan(self):
        d = self.d
        KU = d.cari([r"Keluhan", r"Utama"], x_max=100, awal=True)
        TV = d.cari([r"Tanda", r"Vital"], x_max=200, awal=False)
        PN = d.cari([r"Penilaian", r"Nyeri"], x_max=100, awal=True)
        RPeng = d.cari([r"Riwayat", r"Pengobatan"], x_max=100, awal=True)

        # ---- 1. keluhan utama (keperawatan)
        if KU:
            bawah = (TV["y0"] - 3) if TV else KU["y1"] + 25
            ws = [w for w in d.area(KU["x1"] + 3, 9999, KU["y0"] - 3, bawah)
                  if not w["t"].startswith(":") or len(w["t"]) > 1]
            ws = [dict(w, t=w["t"].lstrip(":")) for w in ws]
            ws = [w for w in ws if w["t"]]
            teks = d.teks(ws)
            self._dbg("1 keluhan utama (keperawatan)", teks)
            ok = _isi(teks) and _keluhan_valid(teks)
            ket = "" if (ok or not _isi(teks)) else "keluhan memuat kata 'kontrol'"
            if ok and _keluhan_ragu(teks):
                ket = "cek manual: keterangan setelah kata kontrol hanya sedikit"
            self._set("keluhan_kep", ok, teks, ket)
        else:
            self._set("keluhan_kep", False, "", "label tidak ditemukan")

        # ---- kolom tabel 4 kolom (Tanda Vital | Nutrisi | Fungsional | Petugas)
        if not TV:
            for k in ("tanda_vital", "nutrisi", "petugas"):
                self._set(k, False, "", "tabel tidak ditemukan")
        else:
            xs = sorted({round(x) for x, a, b in d.V if a <= TV["yc"] <= b and 50 <= x <= 575})
            if len(xs) >= 5:
                kol = xs
            else:  # cadangan: tengah antara judul kolom
                hdr = [TV["x0"] + 20, 262, 387, 515]
                kol = [57] + [(hdr[i] + hdr[i + 1]) / 2 for i in range(3)] + [569]
                self.catatan.append("Garis kolom tabel atas tidak terdeteksi, memakai posisi cadangan")
            g0 = TV["y1"] + 1
            g1 = (PN["y0"] - 2) if PN else g0 + 80
            cols = [d.area(kol[i], kol[i + 1], g0, g1) for i in range(4)]
            self._dbg("kolom: batas x", str(kol))

            # ---- 2. tanda vital
            peta = {"tekanan": "TD", "nadi": "Nadi", "suhu": "Suhu",
                    "nafas|napas": "RR", "kesadaran": "Kesadaran"}
            nilai = {v: "" for v in peta.values()}
            for ln in d._klaster(cols[0]):
                lab, val, _ = self._kv(ln, buang_satuan=True)
                for pola, nm in peta.items():
                    if re.search(pola, lab, re.I):
                        nilai[nm] = val
            kosong = [k for k, v in nilai.items() if not _isi(v)]
            ringkas = "; ".join(f"{k}={v or '-'}" for k, v in nilai.items())
            self._dbg("2 tanda vital", ringkas)
            self._set("tanda_vital", not kosong, ringkas,
                      ("kosong: " + ", ".join(kosong)) if kosong else "")

            # ---- 3. nutrisi
            peta = {"berat": "BB", "tinggi": "TB", "imt": "IMT", "lingkar": "Lingkar Kepala"}
            nilai = {v: "" for v in peta.values()}
            for ln in d._klaster(cols[1]):
                lab, val, _ = self._kv(ln, buang_satuan=True)
                for pola, nm in peta.items():
                    if re.search(pola, lab, re.I):
                        nilai[nm] = val
            umur = self.info.get("umur")
            wajib = ["BB", "TB", "IMT"] + (["Lingkar Kepala"] if (umur is not None and umur < 18) else [])
            kosong = [k for k in wajib if not _isi(nilai[k])]
            ringkas = "; ".join(f"{k}={v or '-'}" for k, v in nilai.items())
            ket = ("kosong: " + ", ".join(kosong)) if kosong else ""
            if umur is None:
                ket = (ket + "; " if ket else "") + "umur tidak terbaca, Lingkar Kepala tidak diwajibkan"
            self._dbg("3 nutrisi", ringkas + f"  (umur={umur})")
            self._set("nutrisi", not kosong, ringkas, ket)

            # ---- 4. petugas
            nama, jam = [], ""
            for ln in d._klaster(cols[3]):
                lab, val, ada = self._kv(ln)
                if ada and re.fullmatch(r"jam", lab.strip(), re.I):
                    jam = val
                elif not ada:
                    nama.append(lab)
            nama = " ".join(nama).strip()
            self._dbg("4 petugas", f"nama={nama!r} jam={jam!r}")
            ket = "" if _isi(jam) else "jam kosong"
            self._set("petugas", _isi(nama), nama, ket if _isi(nama) else "nama petugas kosong")

        # ---- 5. penilaian nyeri
        RP = d.cari([r"Riwayat", r"Penyakit", r":"], x_min=250, setelah=TV["yc"] if TV else None)
        if PN:
            xr = (RP["x0"] - 3) if RP else 315
            g1 = (RPeng["y0"] - 2) if RPeng else PN["y1"] + 100
            ws = d.area(55, xr, PN["y1"] + 1, g1)
            sub = {"skala": "", "onset": "", "pencetus": "", "gambaran": "", "durasi": "", "lokasi": ""}
            nyeri = None
            for ln in d._klaster(ws):
                if (re.fullmatch(r"nyeri", ln[0]["t"], re.I) and ln[0]["x0"] < 100
                        and not any(w["t"].startswith(":") for w in ln)):
                    nyeri = _gabung(ln[1:])
                    continue
                lab, val, _ = self._kv(ln, stop=r"Metode")
                for k in sub:
                    if re.search(k, lab, re.I):
                        sub[k] = val
            ringkas = f"Nyeri={nyeri!r} " + "; ".join(f"{k}={v or '-'}" for k, v in sub.items())
            self._dbg("5 penilaian nyeri", ringkas)
            if nyeri is None or not _isi(nyeri):
                self._set("nyeri", False, ringkas, "pilihan Nyeri (Ya/Tidak) kosong")
            elif re.fullmatch(r"\s*tidak\s*", nyeri, re.I):
                # Nyeri = Tidak sudah cukup, isian lain boleh kosong
                self._set("nyeri", True, ringkas, "Nyeri=Tidak")
            elif re.search(r"ya", nyeri, re.I):
                kosong = [k for k in ("skala", "lokasi") if not _isi(sub[k])]
                self._set("nyeri", not kosong, ringkas,
                          ("Nyeri=Ya, kosong: " + ", ".join(kosong)) if kosong else "")
            else:
                self._set("nyeri", _isi(sub["skala"]), ringkas, "")
        else:
            self._set("nyeri", False, "", "label tidak ditemukan")

        # ---- 6. riwayat penyakit (kotak kanan)
        KE = d.cari([r"Kebutuhan", r"Edukasi"], x_min=250)
        if RP:
            g1 = (KE["y0"] - 2) if KE else RP["y1"] + 40
            ws = d.area(RP["x0"] - 3, 9999, RP["y0"] - 1, g1)
            ws = [w for w in ws
                  if not (RP["y0"] - 1 <= w["yc"] <= RP["y1"] + 1 and w["x1"] <= RP["x1"] + 1)]
            ws = [dict(w, t=w["t"].lstrip(":")) for w in ws]
            ws = [w for w in ws if w["t"]]
            teks = d.teks(ws)
            self._dbg("6 riwayat penyakit", teks)
            self._set("riwayat_penyakit", _isi(teks), teks)
        else:
            self._set("riwayat_penyakit", False, "", "label tidak ditemukan")

        # ---- 7. kebutuhan edukasi (centang)
        if KE:
            g1 = (RPeng["y0"] - 1) if RPeng else KE["y1"] + 80
            ws = d.area(KE["x0"] - 3, 9999, KE["y1"], g1)
            cek = [w for w in ws if any(c in w["t"] for c in CEKLIS)]
            self._dbg("7 edukasi", f"{len(cek)} centang: " + ", ".join(w["t"] for w in cek))
            self._set("edukasi", len(cek) > 0, f"{len(cek)} dicentang",
                      "" if cek else "tidak ada topik edukasi dicentang")
        else:
            self._set("edukasi", False, "", "label tidak ditemukan")

        # ---- 7b. riwayat pengobatan (kotak lebar di bawah; isi mulai dari baris di bawah judul)
        if RPeng:
            mulai = d.cari([r"II\.", r"Pengkajian", r"Medis"], setelah=RPeng["yc"])
            g1 = (mulai["y0"] - 2) if mulai else RPeng["y1"] + 150
            ws = d.area(55, 9999, RPeng["y0"] - 1, g1)
            baris_judul = [w for w in ws if abs(w["yc"] - RPeng["yc"]) <= 3]
            titik2 = [w for w in baris_judul if ":" in w["t"] and w["x0"] >= RPeng["x1"] - 2]
            batas = min((w["x1"] for w in titik2), default=RPeng["x1"])
            ws = [w for w in ws if not (abs(w["yc"] - RPeng["yc"]) <= 3 and w["x1"] <= batas + 1)]
            ws = [dict(w, t=w["t"].lstrip(":")) for w in ws]
            ws = [w for w in ws if w["t"]]
            teks = d.teks(ws)
            self._dbg("7b riwayat pengobatan", teks)
            self._set("riwayat_obat", _isi(teks), teks)
        else:
            self._set("riwayat_obat", False, "", "label tidak ditemukan")

    # ============================================================== bagian II (medis)
    def _medis(self):
        d = self.d
        mulai = d.cari([r"II\.", r"Pengkajian", r"Medis"])
        s0 = mulai["yc"] if mulai else None
        A = d.cari([r"Diagnosis", r"\(A\)"], x_max=110, awal=True, setelah=s0)
        S = d.cari([r"Subyektif|Subjektif", r"\(S\)"], x_max=110, awal=True, setelah=s0)
        O = d.cari([r"Obyektif|Objektif", r"\(O\)"], x_max=110, awal=True, setelah=s0)
        P = d.cari([r"Rencana", r"Terapi", r"\(P\)"], x_max=110, awal=True, setelah=s0)
        T = d.cari([r"Tindakan"], x_max=110, awal=True, setelah=s0)
        G = d.cari([r"Prognosis"], x_max=110, awal=True, setelah=s0)
        DU = d.cari([r"Diagnosis", r"Utama"], x_max=110, awal=True, setelah=s0)
        XC = 140  # isi kolom kanan selalu mulai di x >= ~144

        def isi_area(g0, g1, titik_dua=False):
            ws = d.area(XC, 9999, g0, g1)
            if not titik_dua:
                ws = [w for w in ws if w["t"] != ":"]
            return ws

        # ---- 8. diagnosis (A)
        if A:
            bawah = (S["y0"] - 1) if (S and S["yc"] > A["yc"]) else None
            if bawah is None:
                b = d.fh_bawah(A["yc"] + 2)
                bawah = b if b is not None else A["yc"] + 40
            atas = A["y0"] - 4
            teks = d.teks(isi_area(atas, bawah))
            self._dbg("8 diagnosis (A)", teks)
            self._set("diagnosis_a", _isi(teks), teks)
        else:
            self._set("diagnosis_a", False, "", "label tidak ditemukan")

        # ---- 9-11. subyektif (S)
        if S:
            if O and O["yc"] > S["yc"]:
                akhir = O["y0"] - 2
            elif P:
                akhir = (d.fh_atas(P["yc"]) or P["y0"] - 20)
            else:
                akhir = S["yc"] + 1500
            teks = d.teks(isi_area(S["y0"] - 4, akhir, titik_dua=True))
            self._dbg("S (mentah)", teks)
            self._parse_s(teks)
        else:
            for k in ("s_anamnesis", "s_keluhan", "s_rpd"):
                self._set(k, False, "", "blok Subyektif tidak ditemukan")

        # ---- 12. rencana terapi (P)  -> label di tengah sel, pakai garis atas & bawah sel
        if P:
            atas = d.fh_atas(P["yc"])
            if atas is None:
                atas = P["y0"] - 20
                self.catatan.append("Batas atas sel Rencana Terapi tidak terdeteksi (cek manual)")
            bawah = d.fh_bawah(P["yc"] + 2)
            if T and (bawah is None or T["y0"] - 2 < bawah):
                bawah = T["y0"] - 2
            if bawah is None:
                bawah = P["yc"] + 80
            teks = d.teks(isi_area(atas, bawah))
            self._dbg("12 rencana terapi (P)", f"[{atas:.0f}..{bawah:.0f}]\n{teks}")
            self._set("terapi_p", _isi(teks), teks)
        else:
            self._set("terapi_p", False, "", "label tidak ditemukan")

        # ---- 13. tindakan
        if T:
            bawah = d.fh_bawah(T["yc"] + 1)
            if G and (bawah is None or G["y0"] - 3 < bawah):
                bawah = G["y0"] - 3
            if bawah is None:
                bawah = T["yc"] + 14
            teks = d.teks(isi_area(T["y0"] - 4, bawah))
            self._dbg("13 tindakan", teks)
            self._set("tindakan", _isi(teks), teks)
        else:
            self._set("tindakan", False, "", "label tidak ditemukan")

        # ---- 14. prognosis
        if G:
            bawah = d.fh_bawah(G["yc"] + 1)
            if DU and (bawah is None or DU["y0"] - 3 < bawah):
                bawah = DU["y0"] - 3
            if bawah is None:
                bawah = G["yc"] + 14
            teks = d.teks(isi_area(G["y0"] - 5, bawah))
            self._dbg("14 prognosis", teks)
            self._set("prognosis", _isi(teks), teks)
        else:
            self._set("prognosis", False, "", "label tidak ditemukan")

        # ---- obyektif (O): judul pemeriksaan umum harus ada nilainya
        self._obyektif(O, P)

    # judul pemeriksaan umum pada Obyektif (O): (nama tampil, pola label)
    OBJ_JUDUL = [
        ("KEADAAN UMUM", r"Keadaan\s+Umum"),
        ("KESADARAN", r"Kesadaran"),
        ("TEKANAN DARAH", r"Tekanan\s+Darah"),
        ("NADI", r"Nadi"),
        ("PERNAFASAN", r"Per[nm]a[fp]asan"),
        ("SUHU TUBUH", r"Suhu(?:\s+Tubuh)?"),
    ]

    def _obyektif(self, O, P):
        d = self.d
        if not O:
            self._set("obyektif", False, "", "label tidak ditemukan")
            return
        akhir = (P["y0"] - 2) if (P and P["yc"] > O["yc"]) else O["y0"] + 120
        ws = d.area(140, 9999, O["y0"] - 2, akhir)
        garis = d._klaster(ws)
        # mulai setelah "Pemeriksaan Umum :"; berhenti di baris berlabel lain (mis. "TD :", "Mata :")
        dikenal = re.compile("^(?:" + "|".join(p for _, p in self.OBJ_JUDUL) + r")\s*:", re.I)
        teks_baris, mulai = [], False
        for ln in garis:
            t = _gabung(ln)
            if not mulai:
                if re.search(r"Pemeriksaan\s+Umum", t, re.I):
                    mulai = True
                    t = re.sub(r".*?Pemeriksaan\s+Umum\s*:?", "", t, flags=re.I).strip()
                    if t:
                        teks_baris.append(t)
                continue
            if not dikenal.match(t) and re.match(r"^[A-Za-z][A-Za-z ./]{0,25}\s*:", t):
                break
            teks_baris.append(t)
        if not mulai:      # tidak ada "Pemeriksaan Umum": pakai seluruh blok
            teks_baris = [_gabung(ln) for ln in garis]
        teks = " ".join(teks_baris)
        kurang = []
        for nama, pola in self.OBJ_JUDUL:
            m = re.search(r"(?<![A-Za-z])" + pola + r"\s*:\s*([^;]*)", teks, re.I)
            nilai = m.group(1) if m else ""
            # nilai tidak boleh memuat judul lain (judul kosong diikuti judul berikutnya)
            for _, p2 in self.OBJ_JUDUL:
                m2 = re.search(r"(?<![A-Za-z])" + p2 + r"\s*:", nilai, re.I)
                if m2:
                    nilai = nilai[:m2.start()]
            if not m or not re.search(r"[A-Za-z0-9]", nilai):
                kurang.append(nama)
        self._dbg("O obyektif", f"{teks!r} kurang={kurang}")
        self._set("obyektif", not kurang, teks, ("kosong: " + ", ".join(kurang)) if kurang else "")

    def _parse_s(self, teks):
        """Pecah isi S berdasar sub-label: Anamnesis..., Keluhan Utama :, Riwayat Penyakit Terdahulu :"""
        pola = {
            "s_anamnesis": r"^\s*Anamnesis\s*/\s*Riwayat\s+Penyakit\s+Sekarang\s*:",
            "s_keluhan": r"^\s*Keluhan\s+Utama\s*:",
            "s_rpd": r"^\s*Riwayat\s+Penyakit\s+Terdahulu\s*:",
        }
        stop = r"^\s*(?:KU|RPS|RPD|RPK)\s*:|^\s*Riwayat\s+(?:MRS|Psikososial)\s*:"
        pos = []
        for k, p in pola.items():
            m = re.search(p, teks, re.I | re.M)
            if m:
                pos.append((m.start(), m.end(), k))
        pos.sort()
        stops = [m.start() for m in re.finditer(stop, teks, re.I | re.M)]
        hasil = {}
        for i, (s, e, k) in enumerate(pos):
            akhir = len(teks)
            if i + 1 < len(pos):
                akhir = pos[i + 1][0]
            sl = [x for x in stops if x > e and x < akhir]
            if sl:
                akhir = min(sl)
            hasil[k] = teks[e:akhir]
        for k in pola:
            if k in hasil:
                v = hasil[k]
                self._dbg("S: " + LABEL[k], v)
                ok = _isi(v) and (k != "s_keluhan" or _keluhan_valid(v))
                ket = "keluhan memuat kata 'kontrol'" if (_isi(v) and not ok) else ""
                if ok and k == "s_keluhan" and _keluhan_ragu(v):
                    ket = "cek manual: keterangan setelah kata kontrol hanya sedikit"
                self._set(k, ok, v, ket)
            else:
                self._set(k, False, "", "sub-label tidak ada di PDF")

    # ============================================================== diagnosis & ICD
    def _icd(self):
        d = self.d
        DU = d.cari([r"Diagnosis", r"Utama"], x_max=110, awal=True)
        DS = d.cari([r"Diagnosis", r"Sekunder"], x_max=110, awal=True)
        PR = d.cari([r"Prosedur"], x_max=110, awal=True, setelah=DU["yc"] if DU else None)
        if not DU:
            for k in ("dx_utama", "icd_utama", "dx_sekunder", "icd_sekunder"):
                self._set(k, False, "", "tabel diagnosis tidak ditemukan")
            return
        icd_h = d.cari([r"ICD"], x_min=120, setelah=DU["yc"] - 3, sebelum=DU["yc"] + 3)
        # batas kolom 1|2 dan 2|3
        xs = sorted({round(x) for x, a, b in d.V if a <= DU["yc"] <= b and 60 < x < 565})
        if len(xs) >= 2:
            b1, b2 = xs[0], xs[1]
        elif icd_h:
            b1, b2 = icd_h["x0"] - 13, icd_h["x1"] + 22
            self.catatan.append("Garis kolom ICD tidak terdeteksi, memakai posisi cadangan")
        else:
            b1, b2 = 167, 217
        self._dbg("ICD: batas kolom", f"b1={b1} b2={b2}")

        # baris-baris tabel = ruas garis vertikal di batas kolom 1|2
        ruas = sorted({(round(a, 1), round(b, 1)) for x, a, b in d.V if abs(x - b1) < 1.5})
        pita = [(a, b) for a, b in ruas]

        def pita_berisi(g):
            for a, b in pita:
                if a - 0.5 <= g <= b + 0.5:
                    return (a, b)
            return None

        def kumpul(hdr, batas_bawah):
            """kembalikan list (nama, kode) per baris data di bawah header hdr hingga batas_bawah"""
            baris = []
            ph = pita_berisi(hdr["yc"])
            if ph:
                data = [p for p in pita if p[0] >= ph[1] - 0.5 and p[1] <= batas_bawah + 0.5]
                for a, b in data:
                    nm = d.teks(d.area(57, b1, a, b)).replace("\n", " ")
                    kd = d.teks(d.area(b1, b2, a, b)).replace("\n", " ")
                    baris.append((nm.strip(), kd.strip()))
            else:
                a, b = hdr["y1"] + 3, batas_bawah
                nm = d.teks(d.area(57, b1, a, b)).replace("\n", " ")
                kd = d.area(b1, b2, a, b)
                kodes = [w["t"] for w in kd]
                baris.append((nm.strip(), " ".join(kodes)))
                self.catatan.append("Baris tabel ICD tidak terdeteksi per-baris (dibaca sebagai satu blok)")
            return baris

        def batas(next_anchor):
            if next_anchor:
                ph = pita_berisi(next_anchor["yc"])
                return ph[0] if ph else next_anchor["y0"] - 3
            return DU["yc"] + 200

        b_utama = batas(DS if DS else PR)
        utama = kumpul(DU, b_utama)
        sekunder = kumpul(DS, batas(PR)) if DS else []
        self._dbg("ICD: baris Utama", repr(utama))
        self._dbg("ICD: baris Sekunder", repr(sekunder))

        def kode_valid(s):
            return [t for t in re.split(r"[\s,;]+", s.strip()) if ICD_RE.match(t)]

        # utama
        nm = [n for n, k in utama if _isi(n)]
        kd = [c for n, k in utama for c in kode_valid(k)]
        self._set("dx_utama", bool(nm), "; ".join(nm), "" if nm else "nama diagnosis utama kosong")
        ket = ""
        if not kd:
            ket = "kode ICD utama kosong/tidak valid"
            if any(_isi(k) for n, k in utama):
                ket += f" (terbaca: {[k for n, k in utama if _isi(k)]})"
        self._set("icd_utama", bool(kd), "; ".join(kd), ket)
        # sekunder
        if DS:
            nm = [n for n, k in sekunder if _isi(n)]
            kd = [c for n, k in sekunder for c in kode_valid(k)]
            self._set("dx_sekunder", bool(nm), "; ".join(nm), "" if nm else "diagnosis sekunder kosong")
            self._set("icd_sekunder", bool(kd), "; ".join(kd), "" if kd else "ICD sekunder kosong/tidak valid")
            for i, (n, k) in enumerate(sekunder, 1):
                if _isi(n) and not kode_valid(k):
                    self.catatan.append(f"Sekunder baris {i}: '{_ringkas(n, 30)}' tanpa ICD")
        else:
            self._set("dx_sekunder", False, "", "tabel diagnosis sekunder tidak ditemukan")
            self._set("icd_sekunder", False, "", "tabel diagnosis sekunder tidak ditemukan")

    # ============================================================== kondisi saat meninggalkan fasilitas
    def _kondisi(self):
        d = self.d
        K = d.cari([r"Kondisi", r"Saat", r"Meninggalkan", r"Fasilitas"], x_max=110)
        if not K:
            self._set("kondisi_pulang", False, "", "label tidak ditemukan")
            return
        NO = d.cari([r"Nama", r"Obat"], x_max=110, awal=True, setelah=K["yc"])
        g1 = (NO["y0"] - 3) if NO else K["yc"] + 40
        ws = d.area(K["x1"], 9999, K["y0"] - 8, g1)
        cek = [w for w in ws if any(c in w["t"] for c in CEKLIS)]
        # nama pilihan terdekat di kanan tiap centang
        pilihan = []
        for c in cek:
            sisa = [w for w in ws if w["x0"] >= c["x0"] - 1 and abs(w["yc"] - c["yc"]) < 6
                    and not any(k in w["t"] for k in CEKLIS)]
            sisa.sort(key=lambda w: w["x0"])
            nama = []
            for w in sisa:
                if nama and w["x0"] - nama[-1]["x1"] > 12:
                    break
                nama.append(w)
            pilihan.append(_gabung(nama))
        self._dbg("19 kondisi", f"{len(cek)} centang: {pilihan}")
        self._set("kondisi_pulang", len(cek) > 0, "; ".join(pilihan),
                  "" if cek else "tidak ada kondisi dicentang")

    # ============================================================== tanda tangan dokter
    def _ttd(self):
        d = self.d
        L = d.cari([r"Nama", r"dan", r"Tanda", r"Tangan", r"Dokter"])
        if not L:
            self._set("ttd_dokter", False, "", "kotak tanda tangan tidak ditemukan")
            self.info["dpjp"] = ""
            return
        x0 = max(0, L["x0"] - 45)
        # dasar kotak = garis horizontal lebar (>=150) yang menutup kotak, di bawah label
        dasar = [y for y, a, b in d.HS if y > L["y1"] + 10 and a <= x0 + 10 and b >= 560 and (b - a) >= 150]
        g1 = min(dasar) if dasar else L["y1"] + 160
        g0 = L["y1"] + 2
        ws = d.area(x0, 600, g0, g1)
        nama = d.teks(ws).replace("\n", " ").strip()
        self.info["dpjp"] = nama
        # gambar
        gambar = [im for im in d.imgs
                  if x0 <= (im[0] + im[2]) / 2 <= 600 and g0 <= (im[1] + im[3]) / 2 <= g1
                  and (im[2] - im[0]) > 8 and (im[3] - im[1]) > 8]
        # goresan vektor (kurva) di dalam kotak
        vek = [v for v in d.vekt
               if x0 + 4 <= (v[0] + v[2]) / 2 <= 566 and g0 <= (v[1] + v[3]) / 2 <= g1]
        ada = bool(gambar or len(vek) >= 3)
        self._dbg("20 tanda tangan", f"nama={nama!r} gambar={len(gambar)} goresan_vektor={len(vek)} "
                                      f"area x>={x0:.0f} y[{g0:.0f}..{g1:.0f}]")
        ket = ""
        if not ada:
            ket = "kotak tanda tangan kosong" + (" (nama dokter tercetak)" if _isi(nama) else "")
        self._set("ttd_dokter", ada, nama, ket)

    # ============================================================== hasil
    def _hasil(self):
        item = {k: self.item.get(k) or dict(label=l, terisi=False, nilai="", keterangan="tidak diperiksa")
                for k, l in ITEM}
        kosong = [v["label"] for v in item.values() if not v["terisi"]]
        return dict(
            info=self.info,
            halaman=len(self.d.doc),
            item={v["label"]: v for v in item.values()},
            terisi=len(item) - len(kosong),
            total=len(item),
            tidak_terisi=kosong,
            catatan=self.catatan,
            debug="\n".join(self.log) if self.debug else "",
        )


def periksa_resume(sumber, debug=False):
    """sumber: path PDF atau bytes. Mengembalikan dict hasil (lihat PemeriksaResume._hasil)."""
    if isinstance(sumber, (bytes, bytearray)):
        sumber = fitz.open(stream=bytes(sumber), filetype="pdf")
    return PemeriksaResume(sumber, debug=debug).jalankan()


def _cli():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        return
    h = periksa_resume(args[0], debug="--debug" in sys.argv)
    if "--json" in sys.argv:
        print(json.dumps(h, ensure_ascii=False, indent=2, default=str))
        return
    i = h["info"]
    print(f"{i.get('nama')} | RM {i.get('no_rm')} | {i.get('tanggal_masuk')} | umur {i.get('umur_teks')} | DPJP {i.get('dpjp')}")
    print(f"Terisi {h['terisi']}/{h['total']}  (halaman PDF: {h['halaman']})")
    for v in h["item"].values():
        print(f"  [{'OK ' if v['terisi'] else '-- '}] {v['label']:<38} {v['nilai'][:55]}"
              + (f"   <{v['keterangan']}>" if v["keterangan"] else ""))
    if h["tidak_terisi"]:
        print("Tidak terisi:", ", ".join(h["tidak_terisi"]))
    for c in h["catatan"]:
        print("  catatan:", c)
    if h["debug"]:
        print("\n===== DEBUG =====\n" + h["debug"])


if __name__ == "__main__":
    _cli()
