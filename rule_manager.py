# %%
import json
import os
import sys

def get_base_path():
    """ Mengambil path absolut untuk kompabilitas PyInstaller (.exe) """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

JSON_FILE = os.path.join(get_base_path(), 'aturan_audit.json')

DEFAULT_RULES = [
    {
        "kategori": "Laboratorium",
        "trigger": ["laboratorium", "darah rutin", "kimia klinik", "biaya lab"],
        "syarat": ["hasil pemeriksaan laboratorium", "hemoglobin", "leukosit", "trombosit"],
        "pesan_gagal": "Tidak Ada Hasil Lab"
    },
    {
        "kategori": "Radiologi",
        "trigger": ["radiologi", "rontgen", "thorax", "usg"],
        "syarat": ["ekspertise", "hasil pemeriksaan radiologi", "kesan"],
        "pesan_gagal": "Tidak Ada Ekspertise Radiologi"
    }
]

def muat_aturan():
    """ Membaca file JSON. Jika belum ada, buat otomatis dengan default. """
    if not os.path.exists(JSON_FILE):
        with open(JSON_FILE, 'w') as f:
            json.dump(DEFAULT_RULES, f, indent=4)
        return DEFAULT_RULES
    
    try:
        with open(JSON_FILE, 'r') as f:
            return json.load(f)
    except Exception:
        return []

def simpan_aturan(data):
    with open(JSON_FILE, 'w') as f:
        json.dump(data, f, indent=4)

def tambah_aturan(kategori, list_trigger, list_syarat, pesan_gagal):
    aturan = muat_aturan()
    aturan.append({
        "kategori": kategori,
        "trigger": list_trigger,
        "syarat": list_syarat,
        "pesan_gagal": pesan_gagal
    })
    simpan_aturan(aturan)

def update_aturan(kategori_lama, kategori_baru, list_trigger, list_syarat, pesan_gagal):
    aturan = muat_aturan()
    for item in aturan:
        if item["kategori"].lower() == kategori_lama.lower():
            item["kategori"] = kategori_baru
            item["trigger"] = list_trigger
            item["syarat"] = list_syarat
            item["pesan_gagal"] = pesan_gagal
            break
    simpan_aturan(aturan)

def hapus_aturan(kategori):
    aturan = muat_aturan()
    aturan = [item for item in aturan if item["kategori"].lower() != kategori.lower()]
    simpan_aturan(aturan)


