# %%
import os
import time
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import StaleElementReferenceException
import urllib3
urllib3.disable_warnings()
try:
    import pymupdf as fitz
except ImportError:
    import fitz


# ==========================================
# KONFIGURASI XPATH SIMRS (BISA DISESUAIKAN)
# ==========================================
XPATH_CONFIG = {
    # --- ALUR LOGIN & MENU ---
    "tombol_menu": "//span[contains(@class, 'fa-navicon')]/ancestor::span[contains(@class, 'x-btn-wrap')]",
    "menu_monitoring": "//div[contains(@class, 'x-treelist-item-text') and text()='Monitoring']",
    "menu_kelengkapan": "//div[contains(@class, 'x-treelist-item-text') and text()='Monitoring Kelengkapan Berkas Klaim']",
    
    # --- ALUR FILTER PENCARIAN ---
    "input_poli": "//input[@placeholder='[Pilih Ruangan]']",
    "input_tgl_awal": "//input[@name='PERIODE_AWAL']",
    "input_tgl_akhir": "//input[@name='PERIODE_AKHIR']",
    "tombol_search": "//span[contains(@class, 'fa-search')]/ancestor::span[contains(@class, 'x-btn-wrap')]",
    
    # --- ALUR DATA PASIEN (POP-UP) ---
    "tabel_pasien": "//table[@id='tabel-pasien']/tbody/tr", # Nanti kita sesuaikan jika tabel Ext JS berbeda
    "modal_header": "//h4[@class='modal-title']", 
    "accordion_dokumen": "//div[@class='accordion-item']", 
    "tombol_download_pdf": ".//button[contains(@class, 'btn-download') or contains(text(), 'Lihat PDF')]", 
    "tombol_close_modal": "//button[@aria-label='Close' or contains(@class, 'btn-close')]"
}

class SIMRSScraper:
    def __init__(self, callback_log):
        self.callback_log = callback_log
        self.stop_requested = False
        self.driver = None
        self.path_sesi = None   # folder sesi yang dibuat saat berjalan (dibaca UI)
        

    def setup_driver(self):
        import os
        chrome_options = Options()
        
        # Tentukan folder default awal (menggunakan folder_utama dari GUI jika ada)
        folder_awal = os.path.abspath(self.folder_utama) if hasattr(self, 'folder_utama') and self.folder_utama else os.path.abspath("./downloads")
        os.makedirs(folder_awal, exist_ok=True)

        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": folder_awal,     # Lokasi awal
            "download.prompt_for_download": False,          # Matikan pop-up konfirmasi simpan
            "download.directory_upgrade": True,
            "profile.default_content_settings.popups": 0,   # Blokir pop-up browser
            "profile.default_content_setting_values.automatic_downloads": 1,  # WAJIB: Izinkan download beruntun tanpa blokir
            "safebrowsing.enabled": True,
            "safebrowsing.disable_download_protection": True                  # Hindari dialog proteksi unduhan
        })
        
        # Tambahan argumen agar browser stabil
        # chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-popup-blocking")
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.maximize_window()

    def ubah_folder_download(self, target_folder):
        import os
        target_path = os.path.abspath(target_folder)
        os.makedirs(target_path, exist_ok=True)
        
        try:
            # Standar baru untuk Chrome versi modern (Chrome 110 s/d 155+)
            self.driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": target_path,
                "eventsEnabled": True   
            })
        except Exception:
            # Fallback untuk Chrome versi lama
            self.driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": target_path
            })

    def tunggu_download_selesai(self, folder_path, timeout=30):
        waktu_awal = time.time()
        while time.time() - waktu_awal < timeout:
            file_sementara = [f for f in os.listdir(folder_path) if f.endswith('.crdownload') or f.endswith('.tmp')]
            if not file_sementara:
                return True
            time.sleep(1)
        self.callback_log("⚠️ Timeout menunggu download file PDF.")
        return False

    def bersihkan_nama(self, teks):
        return re.sub(r'[\\/*?:"<>|]', "", teks).strip()

    UKURAN_HALAMAN = 25

    def _folder_lama_cocok(self, path, tanggal):
        """Folder lama (tanpa penanda tanggal) dianggap milik kunjungan ini bila tanggal resume-nya sama."""
        try:
            for f in os.listdir(path):
                if f.lower().endswith(".pdf") and f.upper().startswith("11 - PENGKAJIAN HARIAN"):
                    with fitz.open(os.path.join(path, f)) as d:
                        return tanggal in d[0].get_text("text")
        except Exception:
            pass
        return False

    def tentukan_folder_pasien(self, path_sesi, urutan, no_rm, nama, tgl=""):
        akhir = f"{no_rm}_{nama}"
        m = re.match(r"(\d{2}-\d{2}-\d{4})(?: (\d{2}):(\d{2}))?", tgl or "")
        tag = ""
        if m:
            tag = f" [{m.group(1)}" + (f" {m.group(2)}{m.group(3)}" if m.group(2) else "") + "]"
        target = os.path.join(path_sesi, f"{urutan:04d} - {akhir}{tag}")
        if os.path.isdir(target):
            return target
        if os.path.isdir(path_sesi):
            for d in os.listdir(path_sesi):
                p = os.path.join(path_sesi, d)
                if not os.path.isdir(p):
                    continue
                cocok = re.fullmatch(r"(?:\d{4} - )?" + re.escape(akhir + tag), d) is not None
                if not cocok and tag and re.fullmatch(r"(?:\d{4} - )?" + re.escape(akhir), d):
                    cocok = self._folder_lama_cocok(p, m.group(1))
                if cocok:
                    try:
                        os.rename(p, target)
                        return target
                    except OSError:
                        return p
        return target

    
    def normalisasi_tanggal(self, teks):
        bagian = re.split(r'[-/.]', teks.strip())
        if len(bagian) == 3 and all(b.isdigit() for b in bagian):
            d, m, y = bagian
            return f"{int(d):02d}-{int(m):02d}-{y}"
        return teks
        
    

    def pasang_hook(self):
        """Tangkap body request-report dan requestAccess yang dikirim aplikasi."""
        self.driver.execute_script("""
            if (window.__rrHook2) return;
            window.__rrHook2 = true;
            window.__rrBody = window.__rrBody || null;
            window.__raBody = null;
            function simpan(u, b) {
                u = String(u || '');
                if (!b) return;
                if (u.indexOf('request-report') !== -1) window.__rrBody = b;
                if (u.indexOf('document/requestAccess') !== -1) window.__raBody = b;
            }
            var o = XMLHttpRequest.prototype.open, s = XMLHttpRequest.prototype.send;
            XMLHttpRequest.prototype.open = function(m, u) { this.__u = u; return o.apply(this, arguments); };
            XMLHttpRequest.prototype.send = function(b) { simpan(this.__u, b); return s.apply(this, arguments); };
            var f = window.fetch;
            window.fetch = function(u, opt) {
                try { simpan(u, opt && opt.body); } catch (e) {}
                return f.apply(this, arguments);
            };
        """)
    def ambil_pdf_viewer(self, timeout=15):
        # Tunggu aplikasi mengirim request-report untuk dokumen yang barusan diklik
        t0 = time.time()
        while time.time() - t0 < timeout:
            if self.driver.execute_script("return window.__rrBody !== null && window.__rrBody !== undefined;"):
                break
            time.sleep(0.3)
        else:
            self.callback_log("   ℹ️ Aplikasi tidak mengirim request-report")
            return None

        # Kirim ulang body yang sama, ambil hash baru, unduh PDF-nya (ulang sampai 3x)
        import base64
        hasil = None
        for percobaan in range(3):
            hasil = self.driver.execute_async_script("""
                var cb = arguments[arguments.length - 1];
                fetch('/webservice/plugins/request-report?_dc=' + Date.now(), {
                    method: 'POST',
                    headers: {'content-type': 'application/json', 'x-requested-with': 'XMLHttpRequest'},
                    credentials: 'include',
                    body: window.__rrBody
                })
                .then(function(r) { return r.text(); })
                .then(function(t) {
                    var m = t.match(/[a-f0-9]{64}/);
                    if (!m) { cb('NOHASH:' + t.slice(0, 200)); return; }
                    return new Promise(function(res) { setTimeout(res, 800); })
                    .then(function() {
                        return fetch('/webservice/report/' + m[0], {credentials: 'include'});
                    })
                    .then(function(r) {
                        if (!r.ok) return r.text().then(function(tx) { throw new Error('HTTP ' + r.status + ' ' + tx.slice(0, 120)); });
                        return r.blob();
                    })
                    .then(function(b) {
                        var fr = new FileReader();
                        fr.onloadend = function() { cb(fr.result); };
                        fr.readAsDataURL(b);
                    });
                })
                .catch(function(e) { cb('ERR:' + e.message); });
            """)
            if hasil and str(hasil).startswith("data:"):
                data = base64.b64decode(hasil.split(",", 1)[1])
                if data.startswith(b"%PDF"):
                    return data
            self.callback_log(f"   🔄 Percobaan {percobaan+1}/3 belum siap, mengulang...")
            time.sleep(1.5)
        self.callback_log(f"   ℹ️ Gagal setelah 3x percobaan: {str(hasil)[:160]}")
        return None

    JS_PAGING = r"""
        function cariPaging() {
            var row = document.querySelector('div[id^="berkas-klaim-kelengkapan-grid"] table.x-grid-item');
            if (!row) return null;
            var p = row;
            while (p && !p.querySelector('.x-tbar-page-next')) p = p.parentElement;
            return p;
        }
    """

    def ambil_upload(self, timeout=15, percobaan=3):
        for n in range(1, percobaan + 1):
            data = self._ambil_upload_sekali(timeout if n == 1 else 1)
            if data:
                return data
            if n < percobaan:
                self.callback_log(f"   🔁 Ulangi requestAccess ({n + 1}/{percobaan})...")
                time.sleep(2)
        return None

    def _ambil_upload_sekali(self, timeout=15):
        import base64, json

        # Tunggu aplikasi mengirim requestAccess untuk item yang barusan diklik
        t0 = time.time()
        while time.time() - t0 < timeout:
            if self.driver.execute_script("return !!window.__raBody;"):
                break
            time.sleep(0.3)
        else:
            self.callback_log("   ℹ️ Aplikasi tidak mengirim requestAccess")
            self.diagnosa_upload()
            return None

        hasil = self.driver.execute_async_script(r"""
            var cb = arguments[arguments.length - 1];
            var log = {resp: '', tried: [], data: null};
            fetch('/webservice/document-storage/document/requestAccess?_dc=' + Date.now(), {
                method: 'POST',
                headers: {'content-type': 'application/json'},
                credentials: 'include',
                body: window.__raBody
            })
            .then(function(r) { return r.text(); })
            .then(function(t) {
                log.resp = t.slice(0, 400);
                var cands = [];
                (t.match(/[a-f0-9]{60,64}/g) || []).forEach(function(h) {
                    cands.push('/webservice/document-storage/document/download/' + h);
                });
                try {
                    (function walk(o) {
                        for (var k in o) {
                            var v = o[k];
                            if (typeof v === 'string' && (v.indexOf('/') === 0 || v.indexOf('http') === 0)) cands.push(v);
                            else if (v && typeof v === 'object') walk(v);
                        }
                    })(JSON.parse(t));
                } catch (e) {}
                var i = 0;
                function coba() {
                    if (i >= cands.length) { cb(JSON.stringify(log)); return; }
                    var u = cands[i++];
                    return fetch(u, {credentials: 'include'}).then(function(r) {
                        log.tried.push(u.slice(-30) + ' -> ' + r.status + ' ' + (r.headers.get('content-type') || ''));
                        if (!r.ok) return r.text().then(function(tx) { log.tried.push('body: ' + tx.slice(0, 150)); return coba(); });
                        return r.blob().then(function(b) {
                            var fr = new FileReader();
                            fr.onloadend = function() { log.data = fr.result; cb(JSON.stringify(log)); };
                            fr.readAsDataURL(b);
                        });
                    }).catch(function(e) { log.tried.push(u.slice(-30) + ' ERR ' + e.message); return coba(); });
                }
                return coba();
            })
            .catch(function(e) { log.resp = 'ERR:' + e.message; cb(JSON.stringify(log)); });
        """)

        try:
            info = json.loads(hasil)
        except Exception:
            self.callback_log(f"   ℹ️ Respons tidak terbaca: {str(hasil)[:160]}")
            return None

        if not info.get("data"):
            self.callback_log(f"   🔎 requestAccess resp: {info.get('resp')}")
            self.callback_log(f"   🔎 dicoba: {info.get('tried')}")
            return None

        data = base64.b64decode(info["data"].split(",", 1)[1])
        self.callback_log(f"   🔎 Upload: {len(data)//1024} KB, awal={data[:8]!r}")
        if data.startswith(b"%PDF"):
            return data
        try:   # gambar (jpg/png/dll.) -> jadikan PDF
            gambar = fitz.open(stream=data, filetype=None)
            if gambar.is_pdf:
                return data
            return gambar.convert_to_pdf()
        except Exception as e:
            self.callback_log(f"   ⚠️ Format upload tidak dikenali: {e}")
            return None

    def diagnosa_upload(self):
        info = self.driver.execute_script(r"""
            function dump(w, label) {
                var out = [];
                try {
                    w.performance.getEntriesByType('resource').forEach(function(e) {
                        if (e.name.indexOf('/webservice/') !== -1)
                            out.push(e.initiatorType + ' ' + e.name.replace(/^https?:\/\/[^\/]+/, '').slice(0, 110));
                    });
                } catch (x) { out.push('tidak bisa dibaca'); }
                return label + ': ' + JSON.stringify(out.slice(-8));
            }
            var hasil = ['hook=' + !!window.__rrHook2 + ' | raBody=' + (window.__raBody ? 'ada' : 'null')];
            hasil.push(dump(window, 'utama'));
            var fr = document.querySelectorAll('iframe');
            for (var i = 0; i < fr.length; i++) {
                try { hasil.push(dump(fr[i].contentWindow, 'iframe' + i)); } catch (x) { hasil.push('iframe' + i + ': tidak bisa diakses'); }
            }
            return hasil;
        """)
        for baris in info:
            self.callback_log(f"   🔎 {baris}")

    def info_paging(self):
        return self.driver.execute_script(self.JS_PAGING + r"""
            var p = cariPaging();
            if (!p) return null;
            var inp = p.querySelector('input[role="spinbutton"]');
            var hal = inp ? parseInt(inp.getAttribute('aria-valuenow') || inp.value) : null;
            var total = null;
            p.querySelectorAll('.x-toolbar-text').forEach(function(e) {
                var m = (e.innerText || '').match(/dari\s+(\d+)/i);
                if (m) total = parseInt(m[1]);
            });
            var btn = p.querySelector('.x-tbar-page-next').closest('a.x-btn');
            var row = document.querySelector('div[id^="berkas-klaim-kelengkapan-grid"] table.x-grid-item');
            return {
                hal: hal, total: total,
                disabled: btn ? btn.classList.contains('x-btn-disabled') : false,
                awal: row ? row.innerText.slice(0, 80) : ''
            };
        """)

    def halaman_berikut(self):
        btn = self.driver.execute_script(self.JS_PAGING + r"""
            var p = cariPaging();
            if (!p) return null;
            var n = p.querySelector('.x-tbar-page-next');
            return n.closest('a.x-btn') || n;
        """)
        ActionChains(self.driver).move_to_element(btn)\
            .pause(0.1).click_and_hold().pause(0.1).release().perform()

    def catat_gagal(self, folder_pasien, path_sesi, no_rm, nama_pasien, nama_kat, nama_dok):
        """Catat dokumen yang gagal diunduh: (1) di folder pasien itu sendiri (dipakai rekap Excel, tepat per kunjungan),
        (2) di folder sesi (catatan gabungan seperti sebelumnya)."""
        try:
            with open(os.path.join(folder_pasien, "_gagal_unduh.txt"), "a", encoding="utf-8") as f:
                f.write(f"{nama_kat} | {nama_dok}\n")
        except OSError:
            pass
        try:
            with open(os.path.join(path_sesi, "_gagal_unduh.txt"), "a", encoding="utf-8") as f:
                f.write(f"{no_rm}_{nama_pasien} | {nama_kat} | {nama_dok}\n")
        except OSError:
            pass

    def mask_loading_aktif(self):
        """True bila ada loading mask Ext JS ("Pemuatan...") yang sedang tampil."""
        return bool(self.driver.execute_script("""
            var m = document.querySelectorAll('.x-mask-msg');
            for (var i = 0; i < m.length; i++) {
                if (m[i].offsetParent !== null || (m[i].getClientRects && m[i].getClientRects().length)) return true;
            }
            return false;
        """))

    def tunggu_list_siap(self, timeout=300, stabil=4):
        """Tunggu daftar pasien SELESAI dirender (sabar untuk daftar besar / server lambat):
        loading mask hilang DAN jumlah baris + info halaman tidak berubah selama `stabil` pengecekan
        berturut-turut (sekitar 1 detik per pengecekan). True bila stabil, False bila habis waktu."""
        xpath_baris = "//div[starts-with(@id, 'berkas-klaim-kelengkapan-grid')]//table[contains(@class, 'x-grid-item')]"
        t0 = time.time()
        terakhir, sama, kosong = None, 0, 0
        log_terakhir = t0
        while time.time() - t0 < timeout:
            if self.stop_requested:
                return False
            try:
                if self.mask_loading_aktif():
                    terakhir, sama, kosong = None, 0, 0
                else:
                    n = len(self.driver.find_elements(By.XPATH, xpath_baris))
                    info = self.info_paging() or {}
                    sidik = (n, info.get("hal"), info.get("total"), info.get("awal"))
                    sama = sama + 1 if sidik == terakhir else 0
                    terakhir = sidik
                    kosong = kosong + 1 if n == 0 else 0
                    if n > 0 and sama >= stabil:
                        return True
                    if kosong >= stabil * 3:      # memang tidak ada pasien
                        return True
            except Exception:
                terakhir, sama = None, 0
            if time.time() - log_terakhir >= 15:
                self.callback_log(f"⏳ Daftar pasien masih memuat... ({int(time.time() - t0)} dtk)")
                log_terakhir = time.time()
            time.sleep(1)
        return False

    def tunggu_halaman(self, target, awal_lama, timeout=120):
        t0 = time.time()
        while time.time() - t0 < timeout:
            info = self.info_paging()
            if info and info["hal"] == target and info["awal"] != awal_lama:
                time.sleep(1.5)
                self.tunggu_list_siap(timeout=120, stabil=2)   # tunggu isi halaman selesai dirender
                return True
            time.sleep(0.5)
        return False


    def jalankan_scraping(self, url_simrs, folder_utama, tgl_awal, tgl_akhir, poli, skip_downloaded):
        try:
            tgl_awal = self.normalisasi_tanggal(tgl_awal)
            tgl_akhir = self.normalisasi_tanggal(tgl_akhir)
            self.setup_driver()
            
            # 1. Buka URL Login SIMRS
            self.callback_log(f"Membuka SIMRS: {url_simrs}")
            self.driver.get(url_simrs)
            
            # 2. PAUSE & TUNGGU USER LOGIN MANUAL
            self.callback_log("⏳ Silakan ketik Username, Password, dan klik 'Masuk' secara manual di browser.")
            self.callback_log("⏳ Bot standby menunggu Anda masuk ke halaman utama...")
            
            # Selenium akan menunggu maksimal 5 menit (300 detik) sampai URL mengandung "SIMpel/#main"
            WebDriverWait(self.driver, 300).until(
                EC.url_contains("SIMpel/#main")
            )
            
            if self.stop_requested: return
            
            # 3. SETELAH LOGIN SUKSES, SELENIUM JALAN LAGI
            self.callback_log("✅ Login berhasil terdeteksi! Bot mengambil alih kendali.")
            time.sleep(3) # Jeda ekstra agar seluruh elemen halaman utama Ext JS selesai dimuat

            self.pasang_hook()
            
            # 4. Klik tombol untuk memunculkan list menu
            self.callback_log("Membuka list menu...")
            btn_menu = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH, XPATH_CONFIG["tombol_menu"]))
            )
            self.driver.execute_script("arguments[0].click();", btn_menu)
            time.sleep(1)
            
            # 5. Pilih Menu Monitoring
            self.callback_log("Memilih menu Monitoring...")
            elem_monitoring = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATH_CONFIG["menu_monitoring"]))
            )
            self.driver.execute_script("arguments[0].click();", elem_monitoring)
            time.sleep(1)
            
            # 6. Pilih Monitoring Kelengkapan Berkas
            self.callback_log("Memilih Monitoring Kelengkapan Berkas Klaim...")
            elem_kelengkapan = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATH_CONFIG["menu_kelengkapan"]))
            )
            self.driver.execute_script("arguments[0].click();", elem_kelengkapan)
            time.sleep(3) # Jeda sedikit lama agar halaman dalam Ext JS terbuka sempurna
            
            # ============================================================
            # 7. Pilih Poli/Ruangan via Keyboard (Ketik -> Enter)
            # ============================================================
            from selenium.webdriver.common.keys import Keys
            
            self.callback_log(f"Mengetik Ruangan: {poli}")
            
            # Ambil semua input poli, lalu pilih yang sedang tampil (visible) di layar
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//input[@placeholder='[Pilih Ruangan]']"))
            )
            inputs = self.driver.find_elements(By.XPATH, "//input[@placeholder='[Pilih Ruangan]']")
            input_poli = next((el for el in inputs if el.is_displayed()), inputs[0])
            
            # Beri fokus ke kolom input
            self.driver.execute_script("arguments[0].focus();", input_poli)
            time.sleep(0.3)
            
            # Bersihkan kolom teks
            input_poli.send_keys(Keys.CONTROL, "a")
            input_poli.send_keys(Keys.BACKSPACE)
            time.sleep(0.2)
            
            # Ketik nama poli dan tekan Enter
            input_poli.send_keys(poli)
            time.sleep(1.2)  # Jeda sejenak agar autocomplete Ext JS mengenali teksnya
            input_poli.send_keys(Keys.ENTER)
            time.sleep(0.5)
            
            nilai_terpilih = input_poli.get_attribute("value")
            self.callback_log(f"✔️ Ruangan terisi: '{nilai_terpilih}'")
                        
            # 8. Isi Periode Awal & Akhir via Keyboard
            # ============================================================
            self.callback_log(f"Mengisi Tanggal: {tgl_awal} s/d {tgl_akhir}")
            
            input_awal = self.driver.find_element(By.XPATH, XPATH_CONFIG["input_tgl_awal"])
            input_akhir = self.driver.find_element(By.XPATH, XPATH_CONFIG["input_tgl_akhir"])
            
            # Tanggal Awal
            self.driver.execute_script("arguments[0].focus();", input_awal)
            input_awal.send_keys(Keys.CONTROL, "a")
            input_awal.send_keys(Keys.BACKSPACE)
            time.sleep(0.2)
            input_awal.send_keys(tgl_awal)
            input_awal.send_keys(Keys.TAB) # TAB untuk menutup pop-up kalender
            time.sleep(0.5)
            
            # Tanggal Akhir
            self.driver.execute_script("arguments[0].focus();", input_akhir)
            input_akhir.send_keys(Keys.CONTROL, "a")
            input_akhir.send_keys(Keys.BACKSPACE)
            time.sleep(0.2)
            input_akhir.send_keys(tgl_akhir)
            self.driver.execute_script("arguments[0].blur();", input_akhir)
            time.sleep(0.5)
            
            # ============================================================
            # 9. Klik Kotak "Cari No. RM" lalu Tekan Ctrl + Shift + F
            # ============================================================
            self.callback_log("Mengaktifkan kotak 'Cari No. RM'...")
            from selenium.webdriver.common.action_chains import ActionChains
            from selenium.webdriver.common.keys import Keys
            
            # 1. Cari kotak input "Cari No. RM" yang sedang tampil
            input_rm = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Cari No. RM']"))
            )
            
            # Klik dan beri fokus nyata ke dalam kotak tersebut
            self.driver.execute_script("arguments[0].focus(); arguments[0].click();", input_rm)
            input_rm.click()
            time.sleep(0.5)
            
            # 2. Tekan kombinasi Ctrl + Shift + F (diulang 3 kali seperti uji coba manual)
            for i in range(1, 4):
                self.callback_log(f"Menekan shortcut di kotak RM (ke-{i})...")
                ActionChains(self.driver)\
                    .key_down(Keys.CONTROL)\
                    .key_down(Keys.SHIFT)\
                    .send_keys("f")\
                    .key_up(Keys.SHIFT)\
                    .key_up(Keys.CONTROL)\
                    .perform()
                time.sleep(0.6)
            
            self.callback_log("✔️ Shortcut berhasil dipicu dari kotak Cari No. RM!")
            
            # ============================================================
            # 10. Tunggu Animasi Loading ("Pemuatan...") Selesai Berputar
            # ============================================================
            self.callback_log("Menunggu tabel list pasien selesai memuat...")
            
            # Beri jeda 1 detik agar loading mask Ext JS sempat muncul ke layar
            time.sleep(1)

            # Tunggu SABAR sampai daftar benar-benar selesai dirender (mask hilang + jumlah baris stabil)
            if not self.tunggu_list_siap():
                self.callback_log("⚠️ Daftar pasien belum stabil setelah menunggu lama, lanjut dengan data yang tampil.")

            time.sleep(1) # Buffer render tabel Ext JS

            # Ambil semua baris tabel yang sudah selesai dirender
            xpath_semua_baris = "//div[contains(@class, 'x-grid-item-container')]//table[contains(@class, 'x-grid-item')]"
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.XPATH, xpath_semua_baris))
            )
            
            total_pasien = len(self.driver.find_elements(By.XPATH, xpath_semua_baris))
            self.callback_log(f"✅ Ditemukan {total_pasien} pasien. Memulai urut dari baris pertama.")

            # Buat folder sesi
            nama_sesi = f"[{poli}] {tgl_awal} sd {tgl_akhir}".replace("/", "-")
            path_sesi = os.path.join(folder_utama, nama_sesi)
            os.makedirs(path_sesi, exist_ok=True)
            self.path_sesi = path_sesi

            # ============================================================
            # 11. Loop Membuka Detil (LANGSUNG EKSEKUSI & TERKUNCI DI GRID KLAIM)
            # ============================================================
            from selenium.webdriver.common.action_chains import ActionChains
            from selenium.webdriver.common.keys import Keys
            from selenium.common.exceptions import StaleElementReferenceException

            time.sleep(2) # Tunggu tabel stabil
            
            # Kunci XPath HANYA di tabel kelengkapan berkas klaim agar tidak salah ambil tabel lain
            xpath_baris = "//div[starts-with(@id, 'berkas-klaim-kelengkapan-grid')]//table[contains(@class, 'x-grid-item')]"
            info = self.info_paging()
            halaman = (info or {}).get("hal") or 1
            total_halaman = (info or {}).get("total") or 1

            while True:
                total_pasien = len(self.driver.find_elements(By.XPATH, xpath_baris))
                self.callback_log(f"📄 Halaman {halaman}/{total_halaman}: {total_pasien} pasien")

                for i in range(total_pasien):
                    if self.stop_requested:
                        self.callback_log("⛔ Proses dihentikan oleh user.")
                        break

                    folder_pasien = None     # agar saat error awal tidak memakai folder pasien sebelumnya
                    try:
                        # 1. AMBIL BARIS BERDASARKAN URUTAN
                        daftar_baris = []
                        for _ in range(60):
                            daftar_baris = self.driver.find_elements(By.XPATH, xpath_baris)
                            if i < len(daftar_baris):
                                break
                            time.sleep(0.5)

                        if i >= len(daftar_baris):
                            self.callback_log(f"⛔ Baris ke-{i+1} tidak ditemukan (grid hanya {len(daftar_baris)} baris). Berhenti.")
                            break
                        
                        baris = daftar_baris[i]
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", baris)
                        time.sleep(0.5)

                        # 2. BACA TEKS BARIS VIA JS
                        teks_baris = self.driver.execute_script("return arguments[0].innerText || arguments[0].textContent;", baris)
                        
                        match = re.search(r"(\d{5,10})\s*-\s*([^\n\r]+)", teks_baris)
                        if match:
                            no_rm = match.group(1).strip()
                            nama_kotor = match.group(2).strip()
                            nama_bersih = re.split(r'\n|\d+\s*thn', nama_kotor)[0].strip()
                            nama_pasien = self.bersihkan_nama(nama_bersih)
                        else:
                            no_rm = f"RM_Unknown_{i+1}"
                            nama_pasien = f"Pasien_{i+1}"

                        m_tgl = re.search(r"\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2}", teks_baris) or re.search(r"\d{2}-\d{2}-\d{4}", teks_baris)
                        tgl_daftar = m_tgl.group(0) if m_tgl else ""
                        urutan = (halaman - 1) * self.UKURAN_HALAMAN + i + 1
                        folder_pasien = self.tentukan_folder_pasien(path_sesi, urutan, no_rm, nama_pasien, tgl_daftar)

                        if skip_downloaded and os.path.exists(folder_pasien):
                            if any(f.endswith('.pdf') for f in os.listdir(folder_pasien)):
                                self.callback_log(f"⏭️ [{i+1}/{total_pasien}] Skip {nama_pasien} ({no_rm})")
                                continue

                        self.callback_log(f"📥 [{i+1}/{total_pasien}] Memproses: {nama_pasien} ({no_rm})")
                        self.ubah_folder_download(folder_pasien)
                        # Catatan gagal unduh kunjungan ini dimulai dari nol (hasil proses sebelumnya tidak berlaku lagi)
                        try:
                            os.remove(os.path.join(folder_pasien, "_gagal_unduh.txt"))
                        except OSError:
                            pass

                        # ============================================================
                        # 3. KLIK TOMBOL DETIL (LANGSUNG DIEKSEKUSI, BUKAN DEF)
                        # ============================================================
                        berhasil_buka_detil = False
                        for percobaan in range(3):
                            try:
                                # Ambil ulang baris segar setiap percobaan
                                baris_segar = self.driver.find_elements(By.XPATH, xpath_baris)[i]
                                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", baris_segar)
                                time.sleep(0.4)

                                # Ambil elemen <a> tombol di kolom paling kanan (widgetcolumn)
                                tombol = baris_segar.find_element(By.CSS_SELECTOR, "td[class*='widgetcolumn'] a.x-btn")
                                btn_id = tombol.get_attribute("id")
                                self.callback_log(f"Mengeklik Detil ({btn_id}) - percobaan {percobaan+1}...")

                                # Cara A: Klik fisik dengan ActionChains (mousedown + mouseup)
                                try:
                                    ActionChains(self.driver)\
                                        .move_to_element(tombol)\
                                        .pause(0.15)\
                                        .click_and_hold()\
                                        .pause(0.1)\
                                        .release()\
                                        .perform()
                                except Exception:
                                    pass

                                # Cek cepat 2 detik apakah tombol 'Berkas Klaim' sudah muncul
                                try:
                                    WebDriverWait(self.driver, 2.5).until(
                                        EC.presence_of_element_located((
                                            By.XPATH, "//a[contains(@class,'x-btn') and contains(., 'Berkas Klaim')]"
                                        ))
                                    )
                                    berhasil_buka_detil = True
                                    break
                                except Exception:
                                    # Cara B (Fallback): Tembak handler internal Ext JS + MouseEvent JS
                                    self.driver.execute_script("""
                                        var el = arguments[0];
                                        if (window.Ext && el.id) {
                                            var cmp = Ext.getCmp(el.id);
                                            if (cmp) {
                                                if (typeof cmp.fireHandler === 'function') cmp.fireHandler();
                                                else if (typeof cmp.handler === 'function') cmp.handler.call(cmp.scope || cmp, cmp);
                                                cmp.fireEvent('click', cmp);
                                            }
                                        }
                                        ['mousedown', 'mouseup', 'click'].forEach(function(t) {
                                            el.dispatchEvent(new MouseEvent(t, {bubbles: true, cancelable: true, view: window}));
                                        });
                                    """, tombol)

                                    WebDriverWait(self.driver, 4).until(
                                        EC.presence_of_element_located((
                                            By.XPATH, "//a[contains(@class,'x-btn') and contains(., 'Berkas Klaim')]"
                                        ))
                                    )
                                    berhasil_buka_detil = True
                                    break

                            except StaleElementReferenceException:
                                self.callback_log(f"Baris me-render ulang, mencoba lagi ({percobaan+1}/3)...")
                                time.sleep(0.5)
                            except Exception as err_klik:
                                time.sleep(0.5)

                        if not berhasil_buka_detil:
                            raise Exception("Modal Detil tidak muncul setelah 3x percobaan klik.")

                    # =========================================================
                        # 4. BUKA TAB BERKAS KLAIM & DOWNLOAD SEMUA GRUP PDF
                        # =========================================================
                        self.callback_log("Membuka tab 'Berkas Klaim'...")
                        btn_berkas = WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((
                                By.XPATH, "//a[contains(@class, 'x-btn') and contains(., 'Berkas Klaim')]"
                            ))
                        )
                        ActionChains(self.driver).move_to_element(btn_berkas).click().perform()
                        time.sleep(2)

                        # Tunggu daftar grup berkas muncul
                        xpath_semua_grup = "//div[contains(@class, 'x-grid-group-hd') and @data-groupname]"
                        WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, xpath_semua_grup))
                        )

                        # Ambil semua nama grup (String) agar aman dari StaleElement saat expand
                        daftar_nama_grup = self.driver.execute_script("""
                            var els = document.querySelectorAll("div.x-grid-group-hd[data-groupname]");
                            var names = [];
                            els.forEach(function(el) {
                                var gn = el.getAttribute("data-groupname");
                                if (gn && names.indexOf(gn) === -1) names.push(gn);
                            });
                            return names;
                        """)
                        self.callback_log(f"📋 Ditemukan {len(daftar_nama_grup)} grup: {', '.join(daftar_nama_grup)}")

                        sudah_diproses = set()
                        self.driver.execute_script("window.__rrBody = null;")

                        # Loop setiap grup berkas klaim
                        for idx_grup, nama_kat in enumerate(daftar_nama_grup):
                            try:
                                xpath_grup_ini = f"//div[contains(@class, 'x-grid-group-hd') and @data-groupname='{nama_kat}']"
                                header_grup = WebDriverWait(self.driver, 5).until(
                                    EC.presence_of_element_located((By.XPATH, xpath_grup_ini))
                                )
                                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", header_grup)
                                time.sleep(0.3)

                                masih_tertutup = self.driver.execute_script("""
                                    var hd = arguments[0];
                                    var tr = hd.closest('tr');
                                    return tr ? tr.classList.contains('x-grid-group-hd-collapsed') : true;
                                """, header_grup)

                                if masih_tertutup:
                                    self.callback_log(f"📂 Membuka grup: {nama_kat}")
                                    judul_grup = header_grup.find_element(By.CSS_SELECTOR, "div.x-grid-group-title")
                                    ActionChains(self.driver).move_to_element(judul_grup).click().perform()
                                    time.sleep(1.2)

                                # KUNCI TEPAT KE SEL DAFTAR DOKUMEN (td pertama dengan judul 16px bold, BUKAN tombol Lihat!)
                                xpath_sel_dokumen = (
                                    "//td[contains(@class, 'x-grid-cell-templatecolumn') and contains(@class, 'x-grid-cell-first')"
                                    " and .//div[contains(@style, '16px') and contains(@style, 'bold')]]"
                                )
                                daftar_sel = self.driver.find_elements(By.XPATH, xpath_sel_dokumen)

                                for idx_file, sel_dok in enumerate(daftar_sel):
                                    try:
                                        tinggi_sel = self.driver.execute_script(
                                            "return arguments[0].getBoundingClientRect().height;", sel_dok
                                        )
                                        if tinggi_sel == 0:
                                            continue

                                        # Ambil teks dulu, dipakai sebagai kunci unik
                                        teks_sel = self.driver.execute_script("return arguments[0].innerText;", sel_dok).strip()
                                        kunci = teks_sel.replace("\n", "|")
                                        if kunci in sudah_diproses:
                                            continue
                                        sudah_diproses.add(kunci)

                                        baris_teks = [b.strip() for b in teks_sel.split('\n') if b.strip()]
                                        nama_dok = baris_teks[0] if baris_teks else f"Dokumen_{idx_file+1}"

                                         # Upload: nama item ada di baris ke-2
                                        if "UPLOAD" in nama_kat.upper() and len(baris_teks) > 1:
                                            nama_dok = baris_teks[1][:80]

                                        # Radiologi: format "No | Nama Pemeriksaan | Tanggal" -> ambil segmen tengah
                                        elif "RADIOLOGI" in nama_kat.upper():
                                            for b in baris_teks:
                                                bagian = [x.strip() for x in b.split("|")]
                                                if len(bagian) >= 3:
                                                    nama_dok = bagian[1][:80]
                                                    break

                                        nama_kat_aman = re.sub(r'[\\/*?:"<>|]', '-', nama_kat).strip()
                                        nama_dok_aman = re.sub(r'[\\/*?:"<>|]', '-', nama_dok).strip()
                                        nama_simpan = f"{nama_kat_aman} - {nama_dok_aman}_{idx_file+1}.pdf"
                                        path_file_target = os.path.join(folder_pasien, nama_simpan)

                                        self.callback_log(f"⬇️ Memilih daftar: {nama_dok}")

                                        # Catat isi folder pasien sebelum klik sel
                                        file_sebelum = set(os.listdir(folder_pasien)) if os.path.exists(folder_pasien) else set()

                                        self.driver.execute_script("window.__rrBody = null;")
                                        self.driver.execute_script("window.__raBody = null; performance.clearResourceTimings();")

                                        # 1. KLIK FISIK TEPAT PADA SEL DAFTAR DOKUMEN (td.x-grid-cell-first)
                                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", sel_dok)
                                        time.sleep(0.3)
                                        ActionChains(self.driver)\
                                            .move_to_element(sel_dok)\
                                            .pause(0.1)\
                                            .click_and_hold()\
                                            .pause(0.1)\
                                            .release()\
                                            .perform()

                                        # 2. Tunggu apakah Chrome langsung auto-download PDF (maksimal 5 detik)
                                        path_unduhan_baru = None
                                        

                                        # 3. Ambil PDF                          <- GANTI BLOK INI
                                        if not path_unduhan_baru:
                                            if "UPLOAD" in nama_kat.upper():
                                                raw_pdf = self.ambil_upload()
                                            else:
                                                raw_pdf = self.ambil_pdf_viewer()
                                            if raw_pdf and len(raw_pdf) > 2048:
                                                with open(path_file_target, "wb") as f_out:
                                                    f_out.write(raw_pdf)
                                                self.callback_log(f"   ✅ Tersimpan ({len(raw_pdf)//1024} KB): {nama_simpan}")
                                                continue

                                        # 4. Rapikan nama file jika berhasil masuk via auto-download Chrome
                                        if path_unduhan_baru and os.path.exists(path_unduhan_baru):
                                            if os.path.exists(path_file_target):
                                                os.remove(path_file_target)
                                            os.rename(path_unduhan_baru, path_file_target)
                                            ukuran_kb = os.path.getsize(path_file_target) // 1024
                                            self.callback_log(f"   ✅ Berhasil tersimpan ({ukuran_kb} KB): {nama_simpan}")
                                        else:
                                            # Bersihkan sisa file 1 KB jika ada yang nyasar
                                            sekarang = set(os.listdir(folder_pasien)) if os.path.exists(folder_pasien) else set()
                                            for f_sisa in (sekarang - file_sebelum):
                                                p_sisa = os.path.join(folder_pasien, f_sisa)
                                                if os.path.exists(p_sisa) and os.path.getsize(p_sisa) <= 2048:
                                                    try:
                                                        os.remove(p_sisa)
                                                    except Exception:
                                                        pass
                                            self.callback_log(f"   ⚠️ Belum terunduh untuk: {nama_dok}")
                                            self.catat_gagal(folder_pasien, path_sesi, no_rm, nama_pasien, nama_kat, nama_dok)

                                    except Exception as err_file:
                                        self.callback_log(f"   ⚠️ Lewati dokumen karena error: {str(err_file)}")
                                        self.catat_gagal(folder_pasien, path_sesi, no_rm, nama_pasien, nama_kat,
                                                         f"(error: {str(err_file).splitlines()[0][:60] if str(err_file) else 'tidak diketahui'})")

                            except Exception as err_grup:
                                self.callback_log(f"⚠️ Gagal di grup {nama_kat}: {str(err_grup)}")
                                self.catat_gagal(folder_pasien, path_sesi, no_rm, nama_pasien, nama_kat,
                                                 "(grup gagal dibuka)")

                        # Tutup modal pasien setelah seluruh grup selesai
                        time.sleep(1)
                        self.tutup_modal()
                        self.callback_log(f"✔️ Selesai memproses semua berkas: {nama_pasien}")

                    except Exception as e:
                        self.callback_log(f"❌ Gagal memproses baris {i + 1}: {str(e)}")
                        if folder_pasien and os.path.isdir(folder_pasien):
                            self.catat_gagal(folder_pasien, path_sesi, no_rm, nama_pasien, "SELURUH BERKAS",
                                             f"(proses berhenti karena error: {(str(e).splitlines() or ['?'])[0][:60]})")
                        try:
                            self.tutup_modal()
                        except Exception:
                            pass
                
                if self.stop_requested:
                    break
                info = self.info_paging()
                if not info or info["disabled"] or halaman >= total_halaman:
                    break
                awal_lama = info["awal"]
                self.halaman_berikut()
                halaman += 1
                if not self.tunggu_halaman(halaman, awal_lama):
                    self.callback_log(f"⛔ Halaman {halaman} tidak termuat. Berhenti.")
                    break

            self.callback_log("🎉 Seluruh proses scraping selesai!")

        except Exception as e:
            self.callback_log(f"💥 Error Kritis: {str(e)}")
        finally:
            if self.driver:
                self.driver.quit()

        

    

    def tutup_modal(self, jumlah=2):
        """Tutup window detail pasien, tapi JANGAN window yang memuat grid daftar pasien."""
        for ke in range(jumlah):
            try:
                hasil = self.driver.execute_script("""
                    var tools = document.querySelectorAll('div.x-tool-close');
                    for (var i = tools.length - 1; i >= 0; i--) {
                        var r = tools[i].getBoundingClientRect();
                        if (r.width <= 0 || r.height <= 0) continue;
                        var winEl = tools[i].closest('.x-window');
                        if (!winEl) continue;
                        // lewati window yang berisi grid daftar pasien
                        if (winEl.querySelector('[id^="berkas-klaim-kelengkapan-grid"]')) continue;
                        var judul = (winEl.querySelector('.x-title-text') || {}).innerText || '?';
                        var cmp = window.Ext ? Ext.getCmp(winEl.id) : null;
                        if (cmp) {
                            try { cmp.close(); return winEl.id + ' | ' + judul; }
                            catch (e) { try { cmp.destroy(); return winEl.id + ' | ' + judul; } catch (e2) {} }
                        }
                        ['mousedown', 'mouseup', 'click'].forEach(function(t) {
                            tools[i].dispatchEvent(new MouseEvent(t, {bubbles: true, cancelable: true, view: window}));
                        });
                        return 'klik ' + winEl.id + ' | ' + judul;
                    }
                    return null;
                """)
                if not hasil:
                    break
                self.callback_log(f"   🔒 Tutup window ke-{ke+1}: {hasil}")
                time.sleep(1.2)
            except Exception as e:
                self.callback_log(f"   ⚠️ Gagal menutup window ke-{ke+1}: {e}")
                break
        self.driver.execute_script("window.__rrBody = null;")
        self.driver.execute_script("window.__raBody = null;")

        # Grid daftar pasien di-render ulang setelah window ditutup: tunggu sampai barisnya muncul
        xpath_grid = "//div[starts-with(@id, 'berkas-klaim-kelengkapan-grid')]//table[contains(@class, 'x-grid-item')]"
        for _ in range(40):   # maks 20 detik
            try:
                if len(self.driver.find_elements(By.XPATH, xpath_grid)) > 0:
                    break
            except Exception:
                pass
            time.sleep(0.5)
        time.sleep(1.5)   # beri waktu render selesai agar tidak stale




