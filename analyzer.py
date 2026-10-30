# %%
import os
try:
    import pymupdf as fitz  # PyMuPDF (nama baru)
except ImportError:
    import fitz
import pandas as pd
import rule_manager

# Audit lama berbasis aturan Tab 3 (cari kata di seluruh teks PDF) dimatikan karena hasilnya sering keliru.
# Pemeriksaan resume + bukti penunjang kini ada di rekap_resume.py. Ubah ke True bila ingin menyalakan lagi.
JALANKAN_AUDIT_LAMA = False


class ClaimAnalyzer:
    folder_hasil = None

    def __init__(self, callback_log, callback_progress):
        self.callback_log = callback_log
        self.callback_progress = callback_progress

    def baca_teks_pdf(self, path_pdf):
        teks_lengkap = ""
        try:
            dokumen = fitz.open(path_pdf)
            for hal in dokumen:
                teks_lengkap += hal.get_text("text").lower() + " "
            dokumen.close()
        except Exception as e:
            pass
        return teks_lengkap

    def analisis_semua_berkas(self, folder_utama):
        self.callback_log("Memulai Analisis Dokumen PDF...")

        # Rekap kelengkapan Resume Medis (20 item) -> Rekap_Kelengkapan_Resume.xlsx
        try:
            import rekap_resume
            path_excel = rekap_resume.buat_rekap(folder_utama, self.callback_log, self.callback_progress)
            if path_excel:      # folder sesi terakhir tempat Excel dibuat (dipakai tombol "Buka Folder Hasil Excel")
                self.folder_hasil = os.path.dirname(path_excel)
        except Exception as e:
            self.callback_log(f"⚠️ Rekap Resume gagal: {e}")

        if not JALANKAN_AUDIT_LAMA:
            return   # rekap kelengkapan (resume + bukti) sudah selesai di atas

        aturan_audit = rule_manager.muat_aturan()
        if not aturan_audit:
            self.callback_log("⚠️ Aturan audit kosong! Harap tambahkan aturan di Tab 3.")
            return

        hasil_akhir = []
        
        # Penelusuran folder: boleh pilih folder induk ATAU folder sesi; urut sesuai daftar SIMRS
        import rekap_resume
        semua_pasien_path = rekap_resume.daftar_pasien(folder_utama)
        total_pasien = len(semua_pasien_path)
        if total_pasien == 0:
            self.callback_log("⚠️ Tidak ada folder pasien yang ditemukan.")
            return

        for index, data_p in enumerate(semua_pasien_path):
            self.callback_progress((index + 1) / total_pasien)
            
            # Ekstrak No RM dan Nama
            no_rm = data_p["no_rm"]
            nama_pasien = data_p["nama"] or "Unknown"
            
            file_pdf = [f for f in os.listdir(data_p["path"]) if f.endswith('.pdf')]
            teks_gabungan = ""
            
            for pdf in file_pdf:
                teks_gabungan += self.baca_teks_pdf(os.path.join(data_p["path"], pdf))
            
            data_row = {
                "No Urut": data_p["urutan"],
                "Sesi / Tanggal": data_p["sesi"],
                "No RM": no_rm,
                "Nama Pasien": nama_pasien,
                "Jumlah PDF": len(file_pdf)
            }

            temuan_gagal = []
            
            # Evaluasi Berdasarkan Rule JSON
            for rule in aturan_audit:
                kategori = rule["kategori"]
                trigger_ada = any(t.lower() in teks_gabungan for t in rule["trigger"])
                
                if trigger_ada:
                    syarat_ada = any(s.lower() in teks_gabungan for s in rule["syarat"])
                    if syarat_ada:
                        data_row[kategori] = "Lengkap"
                    else:
                        data_row[kategori] = rule["pesan_gagal"]
                        temuan_gagal.append(rule["pesan_gagal"])
                else:
                    data_row[kategori] = "Tidak Ada Tindakan"

            if len(temuan_gagal) == 0:
                data_row["Kesimpulan Audit"] = "LENGKAP"
            else:
                data_row["Kesimpulan Audit"] = ", ".join(temuan_gagal)

            hasil_akhir.append(data_row)
            self.callback_log(f"✔️ Audit Selesai: {nama_pasien} -> {data_row['Kesimpulan Audit']}")

        # Ekspor ke Excel
        if hasil_akhir:
            peta = {p["sesi"]: p["path_sesi"] for p in semua_pasien_path}
            per_sesi = {}
            for r in hasil_akhir:
                per_sesi.setdefault(r["Sesi / Tanggal"], []).append(r)
            path_excel = folder_utama
            for sesi, rows in per_sesi.items():
                # disimpan satu folder dengan folder-folder pasien pada sesi tersebut
                path_excel = os.path.join(peta.get(sesi, folder_utama), "Rekap_Audit_Klaim.xlsx")
                pd.DataFrame(rows).to_excel(path_excel, index=False)
            self.callback_log("\n==================================")
            self.callback_log(f"✅ EXCEL BERHASIL DIBUAT:\n{path_excel}")
            self.callback_log("==================================")


