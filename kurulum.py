# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  SEYMEN RAPORLAMA — Kurulum Sihirbazı                                  ║
# ║  Çalıştır: python kurulum.py  (Python kurulu sistemler için)           ║
# ║  Python'suz kurulum için: GitHub Releases'dan .exe / .dmg indir       ║
# ╚══════════════════════════════════════════════════════════════════════════╝

import tkinter as tk
from tkinter import ttk, messagebox
import subprocess, sys, os, platform, threading, shutil

UYGULAMA_ADI = "Seymen Raporlama"
SURUMU       = "v1.0"
PAKETLER     = ["customtkinter", "pandas", "openpyxl", "python-pptx", "pillow", "xlsxwriter"]
PROG_DIR     = os.path.dirname(os.path.abspath(__file__))
ANA_DOSYA    = os.path.join(PROG_DIR, "seymen_raporlama.py")
PYTHON_EXE   = sys.executable
RELEASE_URL  = "https://github.com/shtseymen-eng/SeymenRaporlama/releases/latest"

# Renkler
BG = "#0A0E17"; CARD = "#141C2B"; BORDER = "#252D3D"
ALTIN = "#D4AF37"; MAVI = "#2E86DE"; YESIL = "#26A65B"
KIRMIZI = "#C0392B"; METIN = "#E8EDF5"; METIN_DIM = "#7F8C9A"
WIN_RENK = "#0078D4"; MAC_RENK = "#636366"

class KurulumApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{UYGULAMA_ADI} — Kurulum Sihirbazı {SURUMU}")
        self.geometry("660x640")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._os_secim = tk.StringVar(value=self._os_algi())
        self._durum    = tk.StringVar(value="Kuruluma başlamak için sistem seçin.")
        self._arayuz_olustur()
        self._merkeze_al()

    @staticmethod
    def _os_algi():
        return "macOS" if platform.system() == "Darwin" else "Windows"

    def _merkeze_al(self):
        self.update_idletasks()
        w, h = 660, 640
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── Arayüz ───────────────────────────────────────────────────────────────
    def _arayuz_olustur(self):
        # Başlık
        ust = tk.Frame(self, bg=CARD, height=88)
        ust.pack(fill="x"); ust.pack_propagate(False)
        tk.Label(ust, text=UYGULAMA_ADI, bg=CARD, fg=ALTIN,
                 font=("Segoe UI", 22, "bold")).pack(pady=(16,0))
        tk.Label(ust, text="Pregate Operasyon Yönetim Sistemi  ·  Kurulum Sihirbazı",
                 bg=CARD, fg=METIN_DIM, font=("Segoe UI", 10)).pack()
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        # ── Python YOK uyarı kutusu ──────────────────────────────────────────
        uyari = tk.Frame(self, bg="#1A1000", highlightbackground="#D4AF37",
                         highlightthickness=1)
        uyari.pack(fill="x", padx=24, pady=(14,0))
        tk.Label(uyari, bg="#1A1000", fg=ALTIN,
                 text="💡  Python kurulu olmayan bilgisayarlar için:",
                 font=("Segoe UI", 10, "bold"), anchor="w").pack(fill="x", padx=12, pady=(8,2))
        tk.Label(uyari, bg="#1A1000", fg=METIN,
                 text="GitHub Releases sayfasından hazır .exe (Windows) veya .dmg (macOS) indirin —\n"
                      "Python gerekmez, çift tıklayarak doğrudan çalışır.",
                 font=("Segoe UI", 9), anchor="w", justify="left").pack(fill="x", padx=12)
        tk.Button(uyari, text="  🔗  Releases sayfasını aç  ", bg="#2A2000", fg=ALTIN,
                  font=("Segoe UI", 9, "bold"), relief="flat", bd=0, cursor="hand2",
                  padx=8, pady=4,
                  command=lambda: self._url_ac(RELEASE_URL)).pack(anchor="w", padx=12, pady=(4,8))

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=24, pady=(10,0))

        # ── OS seçimi (Python varsa) ─────────────────────────────────────────
        tk.Label(self, text="Python kurulu sistemler için — bağımlılık kurulumu:",
                 bg=BG, fg=METIN_DIM, font=("Segoe UI", 9)).pack(pady=(10,6))

        os_f = tk.Frame(self, bg=BG); os_f.pack()
        self._win_kart = self._os_karti(os_f, "🪟  Windows",
            "pip ile paketler kurulur,\nmasaüstüne kısayol eklenir.",
            "Windows", WIN_RENK, "left")
        tk.Frame(os_f, bg=BG, width=14).pack(side="left")
        self._mac_kart = self._os_karti(os_f, "🍎  macOS",
            "pip ile paketler kurulur,\nmasaüstüne .command eklenir.",
            "macOS", MAC_RENK, "left")
        self._os_secim_guncelle()

        # Paket listesi
        tk.Label(self, text="  ·  ".join(PAKETLER),
                 bg=BG, fg=METIN_DIM, font=("Segoe UI", 8)).pack(pady=(6,0))
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=24, pady=(8,0))

        # İlerleme
        self._progress = ttk.Progressbar(self, length=500, mode="determinate")
        self._progress.pack(pady=(14,4))
        tk.Label(self, textvariable=self._durum, bg=BG, fg=METIN_DIM,
                 font=("Segoe UI", 9), wraplength=520).pack()

        # Log
        self._log = tk.Text(self, height=5, bg=CARD, fg="#7ECB9F",
                            font=("Courier New", 8), bd=0, relief="flat",
                            state="disabled", padx=8, pady=6)
        self._log.pack(fill="x", padx=24, pady=(8,0))

        # Butonlar
        alt = tk.Frame(self, bg=BG); alt.pack(pady=14)
        self._kur_btn = tk.Button(alt, text="  ▶  Kurulumu Başlat  ",
            bg=ALTIN, fg="#0A0E17", font=("Segoe UI", 11, "bold"),
            relief="flat", cursor="hand2", bd=0, padx=16, pady=7,
            command=self._baslat)
        self._kur_btn.pack(side="left", padx=(0,10))
        tk.Button(alt, text="  ✕  Çıkış  ", bg=CARD, fg=METIN,
                  font=("Segoe UI", 10), relief="flat", cursor="hand2",
                  bd=0, padx=12, pady=7, command=self.destroy).pack(side="left")

    # ── OS kart ──────────────────────────────────────────────────────────────
    def _os_karti(self, parent, baslik, acik, deger, renk, side):
        f = tk.Frame(parent, bg=CARD, width=200, height=90, cursor="hand2")
        f.pack(side=side); f.pack_propagate(False)
        tk.Label(f, text=baslik, bg=CARD, fg=METIN,
                 font=("Segoe UI", 12, "bold")).pack(pady=(14,2))
        tk.Label(f, text=acik, bg=CARD, fg=METIN_DIM,
                 font=("Segoe UI", 8), justify="center").pack()
        def _sec(e=None):
            self._os_secim.set(deger); self._os_secim_guncelle()
        for w in [f] + f.winfo_children():
            w.bind("<Button-1>", _sec)
        f._renk = renk; f._deger = deger
        return f

    def _os_secim_guncelle(self):
        for kart in [self._win_kart, self._mac_kart]:
            kart.configure(
                highlightbackground=kart._renk if kart._deger == self._os_secim.get() else BORDER,
                highlightthickness=2 if kart._deger == self._os_secim.get() else 1,
                highlightcolor=kart._renk)

    # ── URL aç ───────────────────────────────────────────────────────────────
    @staticmethod
    def _url_ac(url):
        import webbrowser; webbrowser.open(url)

    # ── Log ──────────────────────────────────────────────────────────────────
    def _log_yaz(self, s):
        self._log.configure(state="normal")
        self._log.insert("end", s + "\n"); self._log.see("end")
        self._log.configure(state="disabled")
        self.update_idletasks()

    # ── Kurulum ──────────────────────────────────────────────────────────────
    def _baslat(self):
        self._kur_btn.configure(state="disabled", text="  ⏳  Kuruluyor...  ")
        threading.Thread(target=self._kur_thread, daemon=True).start()

    def _kur_thread(self):
        os_sec = self._os_secim.get()
        adim = 0; toplam = len(PAKETLER) + 2

        def _ilerleme(msg):
            nonlocal adim; adim += 1
            self._progress["value"] = adim / toplam * 100
            self._durum.set(msg); self._log_yaz(f"  ✔  {msg}")
            self.update_idletasks()

        try:
            self._log_yaz(f"[{os_sec}] Python: {PYTHON_EXE}\n")
            for p in PAKETLER:
                self._durum.set(f"Kuruluyor: {p} ...")
                self.update_idletasks()
                r = subprocess.run([PYTHON_EXE, "-m", "pip", "install", p,
                                    "--quiet", "--disable-pip-version-check"],
                                   capture_output=True, text=True)
                if r.returncode != 0:
                    self._log_yaz(f"  ⚠  {p}: {r.stderr.strip()[:80]}")
                else:
                    _ilerleme(f"{p} hazır")

            masaustu = os.path.join(os.path.expanduser("~"), "Desktop")
            self._durum.set("Kısayol oluşturuluyor ...")
            if os_sec == "Windows":
                self._win_kisayol(masaustu)
            else:
                self._mac_launcher(masaustu)
            _ilerleme("Kısayol oluşturuldu")

            self._progress["value"] = 100
            self._durum.set("✅  Kurulum tamamlandı!")
            self._log_yaz("\n✅  Programı masaüstünden başlatabilirsiniz.")
            self._kur_btn.configure(state="normal", bg=YESIL,
                text="  ✔  Tamam — Kapat  ", command=self.destroy)

        except Exception as e:
            self._durum.set(f"❌  Hata: {e}")
            self._log_yaz(f"\n❌  {e}")
            self._kur_btn.configure(state="normal", bg=KIRMIZI,
                text="  ↩  Tekrar Dene  ")

    def _win_kisayol(self, masaustu):
        bat = os.path.join(masaustu, f"{UYGULAMA_ADI}.bat")
        with open(bat, "w", encoding="utf-8") as f:
            f.write(f'@echo off\n"{PYTHON_EXE}" "{ANA_DOSYA}"\n')
        self._log_yaz(f"  → {bat}")
        try:
            lnk = os.path.join(masaustu, f"{UYGULAMA_ADI}.lnk")
            ico = os.path.join(PROG_DIR, "assets", "logo.ico")
            vbs = (f'Set s=WScript.CreateObject("WScript.Shell")\n'
                   f'Set l=s.CreateShortcut("{lnk}")\n'
                   f'l.TargetPath="{PYTHON_EXE}"\n'
                   f'l.Arguments=chr(34)&"{ANA_DOSYA}"&chr(34)\n'
                   f'l.WorkingDirectory="{PROG_DIR}"\n')
            if os.path.exists(ico):
                vbs += f'l.IconLocation="{ico}"\n'
            vbs += 'l.Save\n'
            vbs_f = os.path.join(PROG_DIR, "_ks.vbs")
            with open(vbs_f,"w") as f: f.write(vbs)
            subprocess.run(["cscript","//nologo",vbs_f], capture_output=True)
            os.remove(vbs_f)
            self._log_yaz(f"  → {lnk}")
        except Exception: pass

    def _mac_launcher(self, masaustu):
        cmd = os.path.join(masaustu, f"{UYGULAMA_ADI}.command")
        with open(cmd,"w") as f:
            f.write(f'#!/bin/bash\ncd "{PROG_DIR}"\n"{PYTHON_EXE}" "{ANA_DOSYA}"\n')
        os.chmod(cmd, 0o755)
        self._log_yaz(f"  → {cmd}")
        try:
            app = os.path.join(masaustu, f"{UYGULAMA_ADI}.app")
            mac_dir = os.path.join(app,"Contents","MacOS")
            os.makedirs(mac_dir, exist_ok=True)
            exe_f = os.path.join(mac_dir, UYGULAMA_ADI.replace(" ",""))
            with open(exe_f,"w") as f:
                f.write(f'#!/bin/bash\n"{PYTHON_EXE}" "{ANA_DOSYA}"\n')
            os.chmod(exe_f, 0o755)
            res = os.path.join(app,"Contents","Resources")
            os.makedirs(res, exist_ok=True)
            plist = os.path.join(app,"Contents","Info.plist")
            with open(plist,"w") as f:
                f.write(f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>CFBundleName</key><string>{UYGULAMA_ADI}</string>
  <key>CFBundleExecutable</key><string>{UYGULAMA_ADI.replace(" ","")}</string>
  <key>CFBundleIdentifier</key><string>com.seymen.raporlama</string>
  <key>CFBundleVersion</key><string>1.0</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>NSHighResolutionCapable</key><true/>
</dict></plist>''')
            for png in ["icon_256.png","icon_512.png","logo.png"]:
                src = os.path.join(PROG_DIR,"assets",png)
                if os.path.exists(src):
                    shutil.copy(src, os.path.join(res,png)); break
            self._log_yaz(f"  → {app}")
        except Exception as ex:
            self._log_yaz(f"  ⚠  .app oluşturulamadı: {ex}")

if __name__ == "__main__":
    KurulumApp().mainloop()
