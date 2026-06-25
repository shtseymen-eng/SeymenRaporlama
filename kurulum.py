# ╔══════════════════════════════════════════════════════════════════════════╗
# ║          SEYMEN RAPORLAMA — Kurulum Sihirbazı                          ║
# ║          Çalıştır: python kurulum.py  (ek paket gerekmez)              ║
# ╚══════════════════════════════════════════════════════════════════════════╝

import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import sys
import os
import platform
import threading
import shutil

# ── Sabitler ──────────────────────────────────────────────────────────────────
UYGULAMA_ADI  = "Seymen Raporlama"
SURUMU        = "v1.0"
PAKETLER      = [
    "customtkinter", "pandas", "openpyxl",
    "python-pptx", "pillow", "xlsxwriter"
]
PROG_DIR = os.path.dirname(os.path.abspath(__file__))
ANA_DOSYA    = os.path.join(PROG_DIR, "seymen_raporlama.py")
PYTHON_EXE   = sys.executable

# ── Renkler ───────────────────────────────────────────────────────────────────
BG        = "#0A0E17"
CARD      = "#141C2B"
BORDER    = "#252D3D"
ALTIN     = "#D4AF37"
MAVI      = "#2E86DE"
YESIL     = "#26A65B"
KIRMIZI   = "#C0392B"
METIN     = "#E8EDF5"
METIN_DIM = "#7F8C9A"
WIN_RENK  = "#0078D4"   # Windows mavi
MAC_RENK  = "#555555"   # macOS gri

# ═════════════════════════════════════════════════════════════════════════════
class KurulumApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{UYGULAMA_ADI} — Kurulum Sihirbazı {SURUMU}")
        self.geometry("620x580")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._os_secim = tk.StringVar(value=self._os_algilا())
        self._durum    = tk.StringVar(value="Kuruluma başlamak için sistem seçin.")
        self._kurulum_tamam = False
        self._arayuz_olustur()
        self._merkeze_al()

    # ── OS algılama ──────────────────────────────────────────────────────────
    @staticmethod
    def _os_algilا():
        s = platform.system()
        return "macOS" if s == "Darwin" else "Windows"

    # ── Pencereyi ekran merkezine al ─────────────────────────────────────────
    def _merkeze_al(self):
        self.update_idletasks()
        w, h = 620, 580
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── Arayüz ───────────────────────────────────────────────────────────────
    def _arayuz_olustur(self):
        # ── Başlık ───────────────────────────────────────────────────────────
        ust = tk.Frame(self, bg=CARD, height=90)
        ust.pack(fill="x")
        ust.pack_propagate(False)

        tk.Label(ust, text=UYGULAMA_ADI, bg=CARD,
                 fg=ALTIN, font=("Segoe UI", 22, "bold")).pack(pady=(18, 0))
        tk.Label(ust, text="Pregate Operasyon Yönetim Sistemi  ·  Kurulum Sihirbazı",
                 bg=CARD, fg=METIN_DIM, font=("Segoe UI", 10)).pack()

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        # ── OS Seçim kartları ─────────────────────────────────────────────────
        tk.Label(self, text="Kurulumu yapılacak işletim sistemini seçin:",
                 bg=BG, fg=METIN, font=("Segoe UI", 11)).pack(pady=(22, 10))

        os_cerceve = tk.Frame(self, bg=BG)
        os_cerceve.pack(pady=4)

        self._win_kart = self._os_karti(
            os_cerceve, "🪟  Windows",
            "Python bağımlılıkları kurulur,\nmasaüstüne kısayol eklenir.",
            "Windows", WIN_RENK, side="left"
        )
        tk.Frame(os_cerceve, bg=BG, width=16).pack(side="left")
        self._mac_kart = self._os_karti(
            os_cerceve, "🍎  macOS",
            "Python bağımlılıkları kurulur,\nmasaüstüne başlatıcı eklenir.",
            "macOS", MAC_RENK, side="left"
        )

        # Algılanan OS'u vurgula
        self._os_secim_guncelle()

        # ── Paket listesi ─────────────────────────────────────────────────────
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=30, pady=(16, 0))
        tk.Label(self, text="Kurulacak paketler:",
                 bg=BG, fg=METIN_DIM, font=("Segoe UI", 9)).pack(pady=(8, 2))
        tk.Label(self, text="  ·  ".join(PAKETLER),
                 bg=BG, fg=METIN_DIM, font=("Segoe UI", 9)).pack()
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=30, pady=(8, 0))

        # ── İlerleme ─────────────────────────────────────────────────────────
        self._progress = ttk.Progressbar(self, length=480, mode="determinate")
        self._progress.pack(pady=(18, 6))

        self._durum_lbl = tk.Label(self, textvariable=self._durum,
                                   bg=BG, fg=METIN_DIM,
                                   font=("Segoe UI", 9), wraplength=500)
        self._durum_lbl.pack()

        # Log kutusu
        self._log = tk.Text(self, height=6, bg=CARD, fg="#7ECB9F",
                            font=("Courier New", 8), bd=0, relief="flat",
                            state="disabled", padx=8, pady=6)
        self._log.pack(fill="x", padx=30, pady=(10, 0))

        # ── Alt butonlar ──────────────────────────────────────────────────────
        alt = tk.Frame(self, bg=BG)
        alt.pack(pady=18)

        self._kur_btn = tk.Button(
            alt, text="  ▶  Kurulumu Başlat  ",
            bg=ALTIN, fg="#0A0E17",
            font=("Segoe UI", 11, "bold"),
            relief="flat", cursor="hand2", bd=0,
            padx=18, pady=8,
            command=self._kurulumu_baslat
        )
        self._kur_btn.pack(side="left", padx=(0, 12))

        self._cik_btn = tk.Button(
            alt, text="  ✕  Çıkış  ",
            bg=CARD, fg=METIN,
            font=("Segoe UI", 10),
            relief="flat", cursor="hand2", bd=0,
            padx=14, pady=8,
            command=self.destroy
        )
        self._cik_btn.pack(side="left")

    # ── OS kart widget'ı ─────────────────────────────────────────────────────
    def _os_karti(self, parent, baslik, aciklama, deger, renk, side):
        cerceve = tk.Frame(parent, bg=CARD, width=210, height=110,
                           cursor="hand2", relief="flat", bd=0)
        cerceve.pack(side=side)
        cerceve.pack_propagate(False)

        tk.Label(cerceve, text=baslik, bg=CARD, fg=METIN,
                 font=("Segoe UI", 13, "bold")).pack(pady=(16, 4))
        tk.Label(cerceve, text=aciklama, bg=CARD, fg=METIN_DIM,
                 font=("Segoe UI", 8), justify="center").pack()

        def _sec(e=None):
            self._os_secim.set(deger)
            self._os_secim_guncelle()

        for w in [cerceve] + cerceve.winfo_children():
            w.bind("<Button-1>", _sec)

        cerceve._renk = renk
        cerceve._deger = deger
        return cerceve

    def _os_secim_guncelle(self):
        secili = self._os_secim.get()
        for kart in [self._win_kart, self._mac_kart]:
            if kart._deger == secili:
                kart.configure(highlightbackground=kart._renk,
                               highlightthickness=2,
                               highlightcolor=kart._renk)
            else:
                kart.configure(highlightthickness=0)

    # ── Log yazıcı ───────────────────────────────────────────────────────────
    def _log_yaz(self, metin):
        self._log.configure(state="normal")
        self._log.insert("end", metin + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")
        self.update_idletasks()

    # ── Kurulum akışı ─────────────────────────────────────────────────────────
    def _kurulumu_baslat(self):
        self._kur_btn.configure(state="disabled", text="  ⏳  Kuruluyor...  ")
        t = threading.Thread(target=self._kur_thread, daemon=True)
        t.start()

    def _kur_thread(self):
        os_sec = self._os_secim.get()
        toplam = len(PAKETLER) + 2   # paketler + kısayol + son adım
        adim   = 0

        def _ilerleme(mesaj):
            nonlocal adim
            adim += 1
            self._progress["value"] = (adim / toplam) * 100
            self._durum.set(mesaj)
            self._log_yaz(f"  ✔  {mesaj}")
            self.update_idletasks()

        try:
            # ── 1. Paketleri kur ─────────────────────────────────────────────
            self._log_yaz(f"[{os_sec}] Kurulum başlıyor — Python: {PYTHON_EXE}\n")
            for paket in PAKETLER:
                self._durum.set(f"Kuruluyor: {paket} ...")
                self.update_idletasks()
                sonuc = subprocess.run(
                    [PYTHON_EXE, "-m", "pip", "install", paket,
                     "--quiet", "--disable-pip-version-check"],
                    capture_output=True, text=True
                )
                if sonuc.returncode != 0:
                    self._log_yaz(f"  ⚠  {paket}: {sonuc.stderr.strip()[:80]}")
                else:
                    _ilerleme(f"{paket} hazır")

            # ── 2. Masaüstü kısayolu ─────────────────────────────────────────
            self._durum.set("Masaüstü kısayolu oluşturuluyor ...")
            masaustu = self._masaustu_yol()

            if os_sec == "Windows":
                self._windows_kisayol(masaustu)
            else:
                self._mac_launcher(masaustu)

            _ilerleme("Masaüstü kısayolu oluşturuldu")

            # ── 3. Tamamlandı ─────────────────────────────────────────────────
            self._progress["value"] = 100
            self._durum.set("✅  Kurulum başarıyla tamamlandı!")
            self._log_yaz("\n✅  Her şey hazır. Programı masaüstünden başlatabilirsiniz.")
            self._kur_btn.configure(
                state="normal", text="  ✔  Kurulum Tamam  ",
                bg=YESIL, command=self.destroy
            )
            self._kurulum_tamam = True

        except Exception as e:
            self._durum.set(f"❌  Hata: {e}")
            self._log_yaz(f"\n❌  HATA: {e}")
            self._kur_btn.configure(
                state="normal", text="  ↩  Tekrar Dene  ", bg=KIRMIZI
            )

    # ── Masaüstü yolu ─────────────────────────────────────────────────────────
    @staticmethod
    def _masaustu_yol():
        if platform.system() == "Windows":
            return os.path.join(os.path.expanduser("~"), "Desktop")
        else:
            return os.path.join(os.path.expanduser("~"), "Desktop")

    # ── Windows kısayolu (.bat) ───────────────────────────────────────────────
    def _windows_kisayol(self, masaustu):
        bat = os.path.join(masaustu, f"{UYGULAMA_ADI}.bat")
        ico = os.path.join(PROG_DIR, "assets", "logo.ico")
        with open(bat, "w", encoding="utf-8") as f:
            f.write(f'@echo off\n')
            f.write(f'"{PYTHON_EXE}" "{ANA_DOSYA}"\n')
        self._log_yaz(f"  → Başlatıcı: {bat}")

        # .lnk kısayolu (pywin32 olmadan, VBS ile)
        try:
            lnk = os.path.join(masaustu, f"{UYGULAMA_ADI}.lnk")
            vbs = (
                f'Set oWS = WScript.CreateObject("WScript.Shell")\n'
                f'sLinkFile = "{lnk}"\n'
                f'Set oLink = oWS.CreateShortcut(sLinkFile)\n'
                f'oLink.TargetPath = "{PYTHON_EXE}"\n'
                f'oLink.Arguments = """{ANA_DOSYA}"""\n'
                f'oLink.WorkingDirectory = "{PROG_DIR}"\n'
            )
            if os.path.exists(ico):
                vbs += f'oLink.IconLocation = "{ico}"\n'
            vbs += 'oLink.Save\n'
            vbs_dosya = os.path.join(PROG_DIR, "_temp_kisayol.vbs")
            with open(vbs_dosya, "w") as f:
                f.write(vbs)
            subprocess.run(["cscript", "//nologo", vbs_dosya],
                           capture_output=True)
            os.remove(vbs_dosya)
            self._log_yaz(f"  → Kısayol (.lnk): {lnk}")
        except Exception:
            pass   # .bat yeterli, .lnk isteğe bağlı

    # ── macOS başlatıcı (.command) ────────────────────────────────────────────
    def _mac_launcher(self, masaustu):
        cmd = os.path.join(masaustu, f"{UYGULAMA_ADI}.command")
        with open(cmd, "w") as f:
            f.write("#!/bin/bash\n")
            f.write(f'cd "{PROG_DIR}"\n')
            f.write(f'"{PYTHON_EXE}" "{ANA_DOSYA}"\n')
        os.chmod(cmd, 0o755)
        self._log_yaz(f"  → Başlatıcı: {cmd}")

        # macOS .app paketi — basit Automator alternatifi (AppleScript shell)
        try:
            app_yol = os.path.join(masaustu, f"{UYGULAMA_ADI}.app")
            mac_dir  = os.path.join(app_yol, "Contents", "MacOS")
            os.makedirs(mac_dir, exist_ok=True)
            exec_yol = os.path.join(mac_dir, UYGULAMA_ADI.replace(" ", ""))
            with open(exec_yol, "w") as f:
                f.write("#!/bin/bash\n")
                f.write(f'"{PYTHON_EXE}" "{ANA_DOSYA}"\n')
            os.chmod(exec_yol, 0o755)

            # Info.plist
            plist = os.path.join(app_yol, "Contents", "Info.plist")
            with open(plist, "w") as f:
                f.write(f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>CFBundleName</key><string>{UYGULAMA_ADI}</string>
  <key>CFBundleExecutable</key><string>{UYGULAMA_ADI.replace(' ','')}</string>
  <key>CFBundleIdentifier</key><string>com.seymen.raporlama</string>
  <key>CFBundleVersion</key><string>1.0</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>LSUIElement</key><false/>
</dict></plist>""")

            # İkon kopyala
            res_dir = os.path.join(app_yol, "Contents", "Resources")
            os.makedirs(res_dir, exist_ok=True)
            for png in ["icon_256.png", "icon_512.png", "logo.png"]:
                src = os.path.join(PROG_DIR, "assets", png)
                if os.path.exists(src):
                    shutil.copy(src, os.path.join(res_dir, png))
                    break

            self._log_yaz(f"  → .app paketi: {app_yol}")
        except Exception as ex:
            self._log_yaz(f"  ⚠  .app oluşturulamadı ({ex}), .command yeterli")


# ═════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = KurulumApp()
    app.mainloop()
