# %%
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
import subprocess
import queue
import sys
import traceback
from datetime import datetime
from scraper import SIMRSScraper
from analyzer import ClaimAnalyzer

# Setup UI Theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Smart Claim & Medical Record Auditor - By Aldy NP")
        self.geometry("950x650")
        self.scraper_instance = None

        # Pesan log & progress dari thread kerja ditampung di antrian, lalu ditampilkan oleh thread utama
        # (mengubah widget langsung dari thread lain bisa membuat aplikasi menutup sendiri, apalagi bila log banyak)
        self._antrian = queue.Queue()
        self.MAKS_BARIS_LOG = 3000
        self.report_callback_exception = self._error_tk
        self.after(100, self._kuras_antrian)

        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tab_scraper = self.tabview.add("1. Robot Download (Selenium)")
        self.tab_analyzer = self.tabview.add("2. Analisis & Audit PDF")

        self.setup_tab_scraper()
        self.setup_tab_analyzer()
    
    # ==========================
    # TAB 1: ROBOT DOWNLOAD
    # ==========================
    def setup_tab_scraper(self):
        frame_input = ctk.CTkFrame(self.tab_scraper)
        frame_input.pack(fill="x", padx=10, pady=5)

        # Baris 1: URL & Poli
        ctk.CTkLabel(frame_input, text="URL SIMRS:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.ent_url = ctk.CTkEntry(frame_input, width=300)
        self.ent_url.grid(row=0, column=1, padx=10, pady=5)
        self.ent_url.insert(0, "http://192.168.8.21/apps/SIMpel/#main")

        ctk.CTkLabel(frame_input, text="Poli/Klinik:").grid(row=0, column=2, padx=10, pady=5, sticky="w")
        self.ent_poli = ctk.CTkEntry(frame_input, width=200)
        self.ent_poli.grid(row=0, column=3, padx=10, pady=5)
        self.ent_poli.insert(0, "Poliklinik Penyakit Dalam")

        # Baris 2: Folder & Tanggal
        ctk.CTkLabel(frame_input, text="Folder Utama:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.ent_folder = ctk.CTkEntry(frame_input, width=300)
        self.ent_folder.grid(row=1, column=1, padx=10, pady=5)
        ctk.CTkButton(frame_input, text="Browse", width=80, command=self.pilih_folder).grid(row=1, column=2, padx=10, pady=5, sticky="w")

        # --- Kolom Tanggal Awal ---
        ctk.CTkLabel(frame_input, text="Tgl Awal (DD-MM-YYYY):").grid(row=1, column=3, padx=10, pady=5, sticky="w")
        self.ent_tgl_awal = ctk.CTkEntry(frame_input, width=120)
        self.ent_tgl_awal.grid(row=1, column=4, padx=10, pady=5)
        self.ent_tgl_awal.insert(0, "01-10-2026")

        # --- Kolom Tanggal Akhir ---
        ctk.CTkLabel(frame_input, text="Tgl Akhir (DD-MM-YYYY):").grid(row=2, column=3, padx=10, pady=5, sticky="w")
        self.ent_tgl_akhir = ctk.CTkEntry(frame_input, width=120)
        self.ent_tgl_akhir.grid(row=2, column=4, padx=10, pady=5)
        self.ent_tgl_akhir.insert(0, "31-10-2026")

        # Baris 3: Option & Buttons (SUDAH DIUBAH KE ROW=3)
        self.chk_skip = ctk.CTkCheckBox(frame_input, text="Auto-Skip pasien yang foldernya sudah berisi PDF")
        self.chk_skip.grid(row=3, column=0, columnspan=2, padx=10, pady=15, sticky="w")
        self.chk_skip.select()

        self.btn_mulai = ctk.CTkButton(frame_input, text="Mulai Download Berkas", fg_color="green", hover_color="darkgreen", command=self.mulai_scraping)
        self.btn_mulai.grid(row=3, column=2, columnspan=2, pady=15, sticky="we", padx=5)

        self.btn_stop = ctk.CTkButton(frame_input, text="Hentikan Proses", fg_color="red", hover_color="darkred", command=self.stop_scraping)
        self.btn_stop.grid(row=3, column=4, pady=15, sticky="we", padx=5)

        # Log
        self.log_scraper = ctk.CTkTextbox(self.tab_scraper, height=350)
        self.log_scraper.pack(fill="both", expand=True, padx=10, pady=10)
    def pilih_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.ent_folder.delete(0, 'end')
            self.ent_folder.insert(0, folder)

    # ---------- infrastruktur log yang aman untuk thread
    def _tulis_error(self, teks):
        """Catat error ke log_error.txt (di samping exe / skrip) supaya penyebab penutupan bisa dilacak."""
        try:
            dasar = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
            with open(os.path.join(dasar, "log_error.txt"), "a", encoding="utf-8") as f:
                f.write(f"[{datetime.now():%d-%m-%Y %H:%M:%S}]\n{teks}\n\n")
        except Exception:
            pass

    def _error_tk(self, exc, val, tb):
        teks = "".join(traceback.format_exception(exc, val, tb))
        self._tulis_error(teks)
        self._antrian.put(("analyzer", f"⚠️ Error tak terduga (tercatat di log_error.txt): {val}"))

    def _aman(self, fungsi, kirim_log):
        """Bungkus fungsi thread: error apa pun dicatat dan ditampilkan, bukan mematikan aplikasi."""
        def jalan(*args):
            try:
                fungsi(*args)
            except BaseException as e:
                self._tulis_error(traceback.format_exc())
                kirim_log(f"❌ Proses berhenti karena error: {type(e).__name__}: {e} (detail di log_error.txt)")
        return jalan

    def _kuras_antrian(self):
        try:
            baru = {"scraper": [], "analyzer": []}
            progress = None
            for _ in range(500):            # batasi per putaran agar tampilan tetap responsif
                try:
                    jenis, isi = self._antrian.get_nowait()
                except queue.Empty:
                    break
                if jenis == "progress":
                    progress = isi
                elif jenis == "sesi":
                    self.ent_folder_analisis.delete(0, "end")
                    self.ent_folder_analisis.insert(0, isi)
                elif jenis == "hasil":
                    self.folder_hasil_terakhir = isi
                else:
                    baru[jenis].append(isi)
            for jenis, kotak in (("scraper", self.log_scraper), ("analyzer", self.log_analyzer)):
                if baru[jenis]:
                    kotak.insert("end", "\n".join(baru[jenis]) + "\n")
                    kelebihan = int(kotak.index("end-1c").split(".")[0]) - self.MAKS_BARIS_LOG
                    if kelebihan > 0:       # buang baris paling lama agar kotak log tidak membengkak
                        kotak.delete("1.0", f"{kelebihan + 1}.0")
                    kotak.see("end")
            if progress is not None:
                self.progress_bar.set(progress)
        except Exception:
            self._tulis_error(traceback.format_exc())
        finally:
            self.after(100, self._kuras_antrian)

    def log_ke_scraper(self, pesan):
        self._antrian.put(("scraper", str(pesan)))

    def mulai_scraping(self):
        url = self.ent_url.get()
        folder = self.ent_folder.get()
        tgl_awal = self.ent_tgl_awal.get().strip()
        tgl_akhir = self.ent_tgl_akhir.get().strip()
        poli = self.ent_poli.get()
        skip = self.chk_skip.get() == 1

        if not folder:
            messagebox.showwarning("Peringatan", "Harap pilih folder penyimpanan utama terlebih dahulu!")
            return

        self.scraper_instance = SIMRSScraper(self.log_ke_scraper)
        self.log_scraper.delete("1.0", "end")
        
        # Jalankan di thread agar GUI tidak freeze
        inst = self.scraper_instance

        def jalan_scraping(*args):
            try:
                inst.jalankan_scraping(*args)
            finally:
                if getattr(inst, "path_sesi", None):      # folder sesi yang baru dibuat -> terisi otomatis di Tab 2
                    self._antrian.put(("sesi", inst.path_sesi))
        threading.Thread(target=self._aman(jalan_scraping, self.log_ke_scraper), args=(url, folder, tgl_awal, tgl_akhir, poli, skip), daemon=True).start()

    def stop_scraping(self):
        if self.scraper_instance:
            self.scraper_instance.stop_requested = True
            self.log_ke_scraper("⚠️ Permintaan berhenti dikirim. Menunggu iterasi saat ini selesai...")

    # ==========================
    # TAB 2: ANALIS & AUDIT
    # ==========================
    def setup_tab_analyzer(self):
        self.folder_hasil_terakhir = None
        frame_a = ctk.CTkFrame(self.tab_analyzer)
        frame_a.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkLabel(frame_a, text="Folder yang dianalisis:", font=("Arial", 14, "bold")).grid(row=0, column=0, padx=10, pady=8, sticky="w")
        self.ent_folder_analisis = ctk.CTkEntry(frame_a, width=520)
        self.ent_folder_analisis.grid(row=0, column=1, padx=5, pady=8, sticky="we")
        ctk.CTkButton(frame_a, text="Browse", width=80, command=self.pilih_folder_analisis).grid(row=0, column=2, padx=10, pady=8)
        ctk.CTkLabel(frame_a, text="Pilih folder sesi (mis. [Poli] 01-10-2026 sd 26-10-2026) atau folder induknya. "
                                   "Terisi otomatis setelah download di Tab 1 selesai.",
                     text_color="gray").grid(row=1, column=0, columnspan=3, padx=10, pady=(0, 8), sticky="w")
        frame_a.grid_columnconfigure(1, weight=1)

        self.btn_scan = ctk.CTkButton(self.tab_analyzer, text="Scan Sub-Folder & Jalankan Audit PDF", height=40, font=("Arial", 14, "bold"), command=self.mulai_analisis)
        self.btn_scan.pack(pady=10, fill="x", padx=100)

        self.progress_bar = ctk.CTkProgressBar(self.tab_analyzer)
        self.progress_bar.pack(fill="x", padx=100, pady=10)
        self.progress_bar.set(0)

        self.log_analyzer = ctk.CTkTextbox(self.tab_analyzer)
        self.log_analyzer.pack(fill="both", expand=True, padx=10, pady=10)

        self.btn_open_excel = ctk.CTkButton(self.tab_analyzer, text="Buka Folder Hasil Excel", fg_color="#E67E22", hover_color="#D35400", command=self.buka_folder_hasil)
        self.btn_open_excel.pack(pady=10)

    def log_ke_analyzer(self, pesan):
        self._antrian.put(("analyzer", str(pesan)))

    def update_progress(self, nilai):
        self._antrian.put(("progress", nilai))

    def pilih_folder_analisis(self):
        folder = filedialog.askdirectory()
        if folder:
            self.ent_folder_analisis.delete(0, 'end')
            self.ent_folder_analisis.insert(0, folder)

    def mulai_analisis(self):
        folder = self.ent_folder_analisis.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Peringatan", "Pilih folder yang akan dianalisis terlebih dahulu (kolom di Tab 2)!")
            return
        self.folder_hasil_terakhir = None
            
        self.log_analyzer.delete("1.0", "end")
        self.progress_bar.set(0)
        analyzer = ClaimAnalyzer(self.log_ke_analyzer, self.update_progress)

        def jalan_analisis(f):
            try:
                analyzer.analisis_semua_berkas(f)
            finally:
                self._antrian.put(("hasil", getattr(analyzer, "folder_hasil", None)))
        threading.Thread(target=self._aman(jalan_analisis, self.log_ke_analyzer), args=(folder,), daemon=True).start()

    def buka_folder_hasil(self):
        # folder sesi terakhir tempat Excel hasil dibuat; bila belum ada analisis, pakai folder di Tab 2
        folder = self.folder_hasil_terakhir or self.ent_folder_analisis.get().strip()
        if folder and os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showinfo("Info", "Belum ada folder hasil. Jalankan analisis dulu atau pilih folder di Tab 2.")

if __name__ == "__main__":
    app = App()
    app.mainloop()


