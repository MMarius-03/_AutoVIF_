import csv
import io
import json
import os
import re
import sys
import subprocess
import threading
import time
import traceback
import pyperclip
import keyboard
import ctypes
import tkinter as tk
from tkinter import ttk
from datetime import datetime
from pywinauto import Application

try:
    import pyautogui
    SCREENSHOT_OK = True
except:
    SCREENSHOT_OK = False

# ======================================================
# INCARCARE CONFIG DIN config.json
# ======================================================
_CONFIG_FILE = "config.json"
try:
    with open(_CONFIG_FILE, encoding="utf-8") as _f:
        _CFG = json.load(_f)
except FileNotFoundError:
    import ctypes as _ctypes
    _ctypes.windll.user32.MessageBoxW(0,
        f"Nu am gasit fisierul de configurare: {_CONFIG_FILE}\n\n"
        f"Asigura-te ca config.json este in acelasi folder cu scriptul.",
        "Eroare configurare", 0)
    sys.exit()
except json.JSONDecodeError as _e:
    import ctypes as _ctypes
    _ctypes.windll.user32.MessageBoxW(0,
        f"Eroare in config.json:\n{_e}\n\nVerifica sintaxa JSON.",
        "Eroare configurare", 0)
    sys.exit()

# Putty
CALE_PUTTY      = _CFG["putty"]["cale"]
SESIUNE_PUTTY   = _CFG["putty"]["sesiune"]
TITLU_PUTTY     = _CFG["putty"]["titlu_fereastra"]

# Fisiere
CALE_CSV        = _CFG["fisiere"]["csv"]
CALE_FILTRE     = _CFG["fisiere"]["filtre"]
CALE_PROGRESS   = _CFG["fisiere"]["progress"]
CALE_FAILED     = _CFG["fisiere"]["failed"]
LOG_FILE        = _CFG["fisiere"]["log"]
SCREENSHOT_DIR  = _CFG["fisiere"]["screenshots"]

# Viteza
INTERVAL_TASTARE   = _CFG["viteza"]["interval_tastare"]
PAUZA_DUPA_TASTA   = _CFG["viteza"]["pauza_dupa_tasta"]
PAUZA_CITIRE_ECRAN = _CFG["viteza"]["pauza_citire_ecran"]
INTERVAL_SCANARE   = _CFG["viteza"]["interval_scanare"]

# Limite
TIMEOUT_ARTICOL        = _CFG["limite"]["timeout_articol"]
ARTICOLE_PER_SESIUNE   = _CFG["limite"]["articole_per_sesiune"]
# 0 = fara limita, orice alt numar = limita per sesiune
# ======================================================


# ======================================================
# CLASS: Logger
# ======================================================
class Logger:
    def __init__(self, log_file, screenshot_dir):
        self.log_file = log_file
        self.screenshot_dir = screenshot_dir
        os.makedirs(screenshot_dir, exist_ok=True)

    def log(self, mesaj, idx=None):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        linie = f"[{ts}]" + (f" [Articol {idx+1}]" if idx is not None else "") + f" {mesaj}"
        print(linie)
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(linie + "\n")

    def screenshot(self, nume="eroare"):
        if SCREENSHOT_OK:
            try:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                cale = os.path.join(self.screenshot_dir, f"{nume}_{ts}.png")
                pyautogui.screenshot(cale)
                self.log(f"Screenshot: {cale}")
            except:
                pass


# ======================================================
# CLASS: CSVReader
# ======================================================
class CSVReader:
    def __init__(self, cale):
        self.cale = cale
        self.header_idx = None
        self.coloane = None
        self.primul_rand = None

    # Regex pentru curatare format Excel: ="valoare" sau ="123" sau =valoare
    _RE_EXCEL = re.compile(r'^=?"?([^"]*)"?$')

    @staticmethod
    def curata(val):
        if val is None:
            return ""
        val = val.strip()
        # Prinde toate variantele: ="val", ="123", =val, "val"
        m = CSVReader._RE_EXCEL.match(val)
        if m:
            return m.group(1).strip()
        return val

    # Regex pentru extragere numar din cantitate (ex: "6.000 Kg" → "6.000")
    _RE_NUMAR = re.compile(r'[-+]?\d+[.,]?\d*')

    @staticmethod
    def converteste_cantitate(cant):
        if not cant:
            return "0"
        # Extragem doar numarul din string (ignora "Kg", "Buc", etc.)
        m = CSVReader._RE_NUMAR.search(cant)
        if not m:
            return cant
        try:
            val = float(m.group().replace(",", "."))
            if val == int(val):
                return str(int(val))
            else:
                return f"{val:.3f}".rstrip("0").rstrip(".").replace(".", ",")
        except:
            return cant.replace(".", ",")

    def gaseste_header(self):
        with open(self.cale, encoding="utf-8") as f:
            linii = f.readlines()
        for i, linie in enumerate(linii):
            if "Articol" in linie and "Denumire" in linie:
                self.header_idx = i
                continut = "".join(linii[i:])
                reader = csv.DictReader(io.StringIO(continut), delimiter=";")
                self.coloane = reader.fieldnames
                self.primul_rand = {k: self.curata(v) for k, v in next(reader, {}).items()}
                return True
        return False

    # Regex pentru validare cod articol — trebuie sa fie numeric
    _RE_COD_VALID = re.compile(r'^\d+$')

    def citeste(self, col_articol, col_cantitate, filtre):
        with open(self.cale, encoding="utf-8") as f:
            linii = f.readlines()
        continut = "".join(linii[self.header_idx:])
        reader = csv.DictReader(io.StringIO(continut), delimiter=";")
        articole = []
        for row in reader:
            cod  = self.curata(row.get(col_articol, ""))
            cant = self.curata(row.get(col_cantitate, ""))
            if not cod:
                continue
            # Sarim peste randurile cu cod non-numeric (TOTAL GENERAL, subtotaluri etc.)
            if not self._RE_COD_VALID.match(cod):
                continue
            if any(cod.startswith(f) for f in filtre):
                continue
            cant = self.converteste_cantitate(cant)
            articole.append((cod, cant))
        return articole


# ======================================================
# CLASS: FailedWriter
# ======================================================
class FailedWriter:
    def __init__(self, cale):
        self.cale = cale

    def scrie(self, cod, cant, motiv):
        existe = os.path.exists(self.cale)
        with open(self.cale, "a", encoding="utf-8", newline="") as f:
            writer = csv.writer(f, delimiter=";")
            if not existe:
                writer.writerow(["Cod Articol", "Cantitate", "Motiv"])
            writer.writerow([cod, cant, motiv])


# ======================================================
# CLASS: PuttyController
# ======================================================
class PuttyController:
    def __init__(self, cale_putty, sesiune, titlu, logger, stop_flag):
        self.cale_putty = cale_putty
        self.sesiune    = sesiune
        self.titlu      = titlu
        self.logger     = logger
        self.stop_flag  = stop_flag
        self.app        = None

    def deschide(self):
        self.logger.log("Deschid PuTTY...")
        subprocess.Popen([self.cale_putty, "-load", self.sesiune])
        time.sleep(4.5)
        self.logger.log("Caut fereastra...")
        try:
            self.app = Application(backend="win32").connect(title=self.titlu)
            self.logger.log(f"PuTTY gasit: {self.titlu}")
        except Exception as e:
            ctypes.windll.user32.MessageBoxW(0, f"Nu am gasit fereastra PuTTY!\n\nEroare: {e}", "Eroare", 0)
            sys.exit()

    def get(self):
        try:
            w = self.app.window(title=self.titlu)
            w.set_focus()
            return w
        except:
            time.sleep(1)
            w = self.app.window(title=self.titlu)
            w.set_focus()
            return w

    def exista(self):
        try:
            self.app.window(title=self.titlu).exists()
            return True
        except:
            return False

    def apasa(self, tasta, pauza=PAUZA_DUPA_TASTA):
        if self.stop_flag.is_set():
            return
        self.get().type_keys(tasta)
        time.sleep(pauza)

    def scrie(self, text, pauza=PAUZA_DUPA_TASTA):
        if self.stop_flag.is_set():
            return
        self.get().type_keys(text, with_spaces=True, pause=INTERVAL_TASTARE)
        time.sleep(pauza)

    def citeste_ecran(self):
        pyperclip.copy('')
        self.get().type_keys("%{SPACE}o")
        time.sleep(PAUZA_CITIRE_ECRAN)
        return pyperclip.paste()

    # Regex patterns pentru erori cunoscute pe ecran AS400/VIF
    _RE_ERORI = re.compile(
        r'(inexistent|nu exist|invalid|not found|eroare|error|'
        r'incorect|gresit|lipsa|introduct.*eronat)',
        re.IGNORECASE
    )

    def asteapta_text(self, text, timeout=15, idx=None, mesaj_eroare=None):
        t_start = time.time()
        while time.time() - t_start < timeout:
            # Verifica ESC/F1 la fiecare iteratie
            if self.stop_flag.is_set():
                self.logger.log("Stop detectat in asteapta_text!", idx)
                return False
            ecran = self.citeste_ecran()
            if text in ecran:
                return True
            # Detectare erori cunoscute cu regex
            if self._RE_ERORI.search(ecran):
                if idx is not None:
                    self.logger.log(f"Eroare detectata pe ecran!", idx)
                return False
            time.sleep(INTERVAL_SCANARE)
        if mesaj_eroare:
            self.logger.log(f"TIMEOUT: {mesaj_eroare}", idx)
        return False

    def detecteaza_ecran(self):
        """Detecteaza ecranul curent cu regex si returneaza un string descriptiv"""
        ecran = self.citeste_ecran()
        patterns = {
            "SELECTIE_OF":    r'CTRL-O=Nou OF',
            "CREARE_OF":      r'CREAREA UNUI OF',
            "INTROD_CANT":    r'INTROD CANTITATILOR DE INTRARE',
            "CONFIRMARE":     r'Sunteti sigur',
            "LOGIN":          r'login as:',
            "LOCATIE":        r'Locatie:',
            "MENIU":          r'Fabricatie',
            "POST_LUCRU":     r'Post de lucru',
            "EROARE_ECRAN":   r'(inexistent|nu exist|invalid|not found)',
        }
        for nume, pattern in patterns.items():
            if re.search(pattern, ecran, re.IGNORECASE):
                return nume
        return "NECUNOSCUT"


# ======================================================
# CLASS: OFAutomation
# ======================================================
class OFAutomation:
    def __init__(self):
        self.logger      = Logger(LOG_FILE, SCREENSHOT_DIR)
        self.stop_flag   = threading.Event()  # creat primul!
        self.putty       = PuttyController(CALE_PUTTY, SESIUNE_PUTTY, TITLU_PUTTY, self.logger, self.stop_flag)
        self.csv_reader  = CSVReader(CALE_CSV)
        self.failed      = FailedWriter(CALE_FAILED)
        self.articole    = []
        self.start_index = 0

    # -------------------------------------------------------
    # Stop listener
    # -------------------------------------------------------
    def start_listener(self):
        def _listen():
            while not self.stop_flag.is_set():
                if keyboard.is_pressed("esc") or keyboard.is_pressed("f1"):
                    self.stop_flag.set()
                time.sleep(0.1)
        threading.Thread(target=_listen, daemon=True).start()

    def opreste(self, idx, motiv):
        with open(CALE_PROGRESS, "w") as f:
            f.write(str(idx))
        self.logger.screenshot(f"stop_{idx+1}")
        self.logger.log(f"OPRIT — {motiv}", idx)
        ctypes.windll.user32.MessageBoxW(0,
            f"Programul s-a oprit!\n\nMotiv: {motiv}\nArticol: {idx+1}\n\n"
            f"Progres salvat. Log + screenshot in folderul data\\.",
            "Oprit", 0)
        sys.exit()

    def verifica_stop(self, idx):
        if self.stop_flag.is_set():
            self.opreste(idx, "Oprit manual (ESC sau F1)")

    # -------------------------------------------------------
    # Error recovery dialog
    # -------------------------------------------------------
    def dialog_eroare(self, cod, cant, motiv):
        """Arata dialog cu 3 optiuni: Stop / Skip & Log / Retry"""
        rezultat = [None]

        def on_stop():
            rezultat[0] = "stop"
            d.destroy()

        def on_skip():
            rezultat[0] = "skip"
            d.destroy()

        d = tk.Tk()
        d.title("Eroare articol")
        d.geometry("420x220")
        d.resizable(False, False)
        d.attributes("-topmost", True)

        tk.Label(d, text="⚠️ Eroare la introducerea articolului",
                 font=("Arial", 11, "bold"), fg="#e74c3c", pady=10).pack()
        tk.Label(d, text=f"Cod: {cod}    Cantitate: {cant}",
                 font=("Courier", 10)).pack()
        tk.Label(d, text=f"Motiv: {motiv}",
                 font=("Arial", 9), fg="gray", wraplength=380).pack(pady=5)

        btn_f = tk.Frame(d)
        btn_f.pack(pady=15)

        tk.Button(btn_f, text="⛔ Stop Script", command=on_stop,
                  bg="#e74c3c", fg="white", font=("Arial", 10, "bold"),
                  width=14, height=2).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_f, text="⏭ Skip & Log", command=on_skip,
                  bg="#f39c12", fg="white", font=("Arial", 10, "bold"),
                  width=14, height=2).pack(side=tk.LEFT, padx=8)

        d.mainloop()
        return rezultat[0]

    # -------------------------------------------------------
    # Selectie coloane CSV
    # -------------------------------------------------------
    def selectie_coloane(self):
        col_art  = [None]
        col_cant = [None]

        def on_confirma():
            if not col_art[0] or not col_cant[0]:
                return
            if col_art[0] == col_cant[0]:
                return
            root_col.destroy()

        root_col = tk.Tk()
        root_col.title("Selectie coloane CSV")
        root_col.geometry("500x400")
        root_col.resizable(False, False)

        tk.Label(root_col, text="Selecteaza coloanele din CSV:",
                 font=("Arial", 11, "bold"), pady=10).pack()

        frame_prev = tk.LabelFrame(root_col, text="Preview (primul rand)", padx=5, pady=5)
        frame_prev.pack(fill=tk.X, padx=15, pady=5)
        for col in self.csv_reader.coloane:
            val = self.csv_reader.primul_rand.get(col, "")
            tk.Label(frame_prev, text=f"{col}: {val}", anchor="w",
                     font=("Courier", 9)).pack(fill=tk.X)

        frame_sel = tk.LabelFrame(root_col, text="Alege coloanele", padx=10, pady=10)
        frame_sel.pack(fill=tk.X, padx=15, pady=10)

        tk.Label(frame_sel, text="COD ARTICOL:").grid(row=0, column=0, sticky="w", pady=5)
        combo_a = ttk.Combobox(frame_sel, values=self.csv_reader.coloane, state="readonly", width=25)
        combo_a.grid(row=0, column=1, padx=10)

        tk.Label(frame_sel, text="CANTITATE:").grid(row=1, column=0, sticky="w", pady=5)
        combo_c = ttk.Combobox(frame_sel, values=self.csv_reader.coloane, state="readonly", width=25)
        combo_c.grid(row=1, column=1, padx=10)

        for c in self.csv_reader.coloane:
            if "articol" in c.lower() or "cod" in c.lower():
                combo_a.set(c); col_art[0] = c; break
        for c in self.csv_reader.coloane:
            if any(x in c.lower() for x in ["cant", "ctq", "cte", "qty"]):
                combo_c.set(c); col_cant[0] = c; break

        combo_a.bind("<<ComboboxSelected>>", lambda e: col_art.__setitem__(0, combo_a.get()))
        combo_c.bind("<<ComboboxSelected>>", lambda e: col_cant.__setitem__(0, combo_c.get()))

        btn_f = tk.Frame(root_col)
        btn_f.pack(pady=10)
        tk.Button(btn_f, text="✓ Confirma", command=on_confirma,
                  bg="#2ecc71", fg="white", font=("Arial", 10, "bold"),
                  width=12, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_f, text="✗ Anuleaza", command=sys.exit,
                  bg="#e74c3c", fg="white", font=("Arial", 10, "bold"),
                  width=12, height=2).pack(side=tk.LEFT, padx=10)

        root_col.mainloop()
        return col_art[0], col_cant[0]

    # -------------------------------------------------------
    # Fereastra confirmare articole
    # -------------------------------------------------------
    def confirmare_start(self):
        confirmat = [False]

        def on_da():
            confirmat[0] = True
            root.destroy()

        root = tk.Tk()
        root.title(f"Articole de introdus ({len(self.articole) - self.start_index} total)")
        root.geometry("500x500")
        root.resizable(True, True)

        tk.Label(root, text=f"Urmeaza {len(self.articole) - self.start_index} articole:",
                 font=("Arial", 10, "bold"), pady=5).pack()

        frame = tk.Frame(root)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("Nr", "Cod Articol", "Cantitate")
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=18)
        for c in cols:
            tree.heading(c, text=c)
        tree.column("Nr", width=50, anchor="center")
        tree.column("Cod Articol", width=200, anchor="center")
        tree.column("Cantitate", width=150, anchor="center")

        sb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(fill=tk.BOTH, expand=True)

        for idx, (cod, cant) in enumerate(self.articole[self.start_index:], start=1):
            tree.insert("", tk.END, values=(idx, cod, cant))

        btn_f = tk.Frame(root)
        btn_f.pack(pady=10)
        tk.Button(btn_f, text="✓ DA — Incepe", command=on_da,
                  bg="#2ecc71", fg="white", font=("Arial", 11, "bold"),
                  width=15, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_f, text="✗ NU — Anuleaza", command=sys.exit,
                  bg="#e74c3c", fg="white", font=("Arial", 11, "bold"),
                  width=15, height=2).pack(side=tk.LEFT, padx=10)

        root.mainloop()
        return confirmat[0]

    # -------------------------------------------------------
    # Navigare initiala la p04a
    # -------------------------------------------------------
    def navigare_initiala(self):
        self.logger.log("Navigare la p04a...")

        if not self.putty.asteapta_text("login as:", timeout=15):
            self.opreste(self.start_index, "login as: nu a aparut")
        self.putty.scrie("p04")
        self.putty.apasa("{ENTER}")

        if not self.putty.asteapta_text("Locatie:", timeout=15):
            self.opreste(self.start_index, "Ecranul Locatie nu a aparut")
        self.putty.apasa("{TAB}")
        self.putty.apasa("{TAB}")
        self.putty.scrie("2")
        self.putty.apasa("{ENTER}")
        self.putty.apasa("{ENTER}")

        if not self.putty.asteapta_text("Fabricatie", timeout=15):
            self.opreste(self.start_index, "Meniul principal nu a aparut")
        self.putty.apasa("{RIGHT}")
        self.putty.apasa("{RIGHT}")
        self.putty.apasa("{ENTER}")
        self.putty.apasa("{ENTER}")

        if not self.putty.asteapta_text("Post de lucru", timeout=15):
            self.opreste(self.start_index, "Ecranul Post de lucru nu a aparut")
        self.putty.scrie("p04a")
        self.putty.apasa("{ENTER}")

        if not self.putty.asteapta_text("CTRL-O=Nou OF", timeout=15):
            self.opreste(self.start_index, "Ecranul SELECTIE OF nu a aparut")
        self.logger.log("Ecran SELECTIE OF incarcat!")

    # -------------------------------------------------------
    # Introducere un articol
    # -------------------------------------------------------
    def introduce_articol(self, i, cod, cant):
        self.putty.apasa("^o")
        if not self.putty.asteapta_text("CREAREA UNUI OF", timeout=10, idx=i):
            return False, "Ecranul CREAREA UNUI OF nu s-a incarcat"

        self.putty.scrie(cod)
        self.putty.apasa("{ENTER}")

        if not self.putty.asteapta_text(cod, timeout=10, idx=i):
            return False, f"Articolul {cod} nu a fost recunoscut de server"

        self.putty.scrie(cant)
        self.putty.apasa("{ENTER}")
        self.putty.apasa("{F2}")

        if not self.putty.asteapta_text("INTROD CANTITATILOR DE INTRARE", timeout=15, idx=i):
            return False, "Salvarea OF a esuat"

        self.putty.apasa("{F4}")
        if self.putty.asteapta_text("Sunteti sigur", timeout=5):
            self.putty.apasa("{ENTER}")

        if not self.putty.asteapta_text("CTRL-O=Nou OF", timeout=10, idx=i):
            return False, "Nu s-a revenit la SELECTIE OF"

        return True, None

    # -------------------------------------------------------
    # Revenire la ecranul SELECTIE OF (dupa eroare/skip)
    # Citeste ecranul si se adapteaza in functie de unde e
    # -------------------------------------------------------
    def revino_la_selectie(self, i):
        self.logger.log("Revenire la ecranul SELECTIE OF...", i)
        ecran_curent = self.putty.detecteaza_ecran()
        self.logger.log(f"Ecran detectat: {ecran_curent}", i)

        if ecran_curent == "SELECTIE_OF":
            self.logger.log("Deja pe SELECTIE OF.", i)
            return True

        # Pe orice ecran care necesita F4 pentru iesire
        if ecran_curent in ("CREARE_OF", "INTROD_CANT", "NECUNOSCUT", "EROARE_ECRAN"):
            self.putty.apasa("{F4}")
            if self.putty.asteapta_text("Sunteti sigur", timeout=3):
                self.putty.apasa("{ENTER}")

        if self.putty.asteapta_text("CTRL-O=Nou OF", timeout=10):
            self.logger.log("Revenit la SELECTIE OF.", i)
            return True
        else:
            self.logger.log("NU s-a putut reveni la SELECTIE OF!", i)
            return False

    # -------------------------------------------------------
    # Bucla principala
    # -------------------------------------------------------
    def ruleaza(self):
        for i in range(self.start_index, len(self.articole)):
            self.verifica_stop(i)

            if not self.putty.exista():
                self.opreste(i, "Fereastra PuTTY a disparut!")

            cod, cant = self.articole[i]
            self.logger.log(f"Introducere: {cod} - {cant}", i)

            try:
                succes, motiv = self.introduce_articol(i, cod, cant)

                if not succes:
                    self.logger.log(f"EROARE: {motiv}", i)
                    self.logger.screenshot(f"eroare_{i+1}")

                    actiune = self.dialog_eroare(cod, cant, motiv)

                    if actiune == "stop":
                        self.opreste(i, f"Oprit de utilizator dupa eroare: {motiv}")
                    elif actiune == "skip":
                        self.failed.scrie(cod, cant, motiv)
                        self.logger.log(f"SKIP: {cod} salvat in failed.csv", i)
                        if not self.revino_la_selectie(i):
                            self.opreste(i, "Nu s-a putut reveni la SELECTIE OF dupa skip")
                        continue
                else:
                    self.logger.log(f"OK: {cod}", i)

            except Exception as e:
                self.logger.log(f"EROARE: {traceback.format_exc()}", i)
                actiune = self.dialog_eroare(cod, cant, str(e))
                if actiune == "stop":
                    self.opreste(i, str(e))
                elif actiune == "skip":
                    self.failed.scrie(cod, cant, str(e))
                    self.logger.log(f"SKIP: {cod} salvat in failed.csv", i)
                    if not self.revino_la_selectie(i):
                        self.opreste(i, "Nu s-a putut reveni la SELECTIE OF dupa skip")
                    continue

        if os.path.exists(CALE_PROGRESS):
            os.remove(CALE_PROGRESS)

        self.logger.log("Procesare finalizata!")
        ctypes.windll.user32.MessageBoxW(0,
            f"Gata! S-au introdus {len(self.articole) - self.start_index} articole.",
            "Finalizat", 0)

    # -------------------------------------------------------
    # Main
    # -------------------------------------------------------
    def run(self):
        self.start_listener()
        self.putty.deschide()

        # Citire CSV
        if not self.csv_reader.gaseste_header():
            ctypes.windll.user32.MessageBoxW(0, "Nu am gasit headerul in CSV!", "Eroare", 0)
            sys.exit()

        col_art, col_cant = self.selectie_coloane()
        self.logger.log(f"Coloane: Articol='{col_art}', Cantitate='{col_cant}'")

        # Citire filtre
        filtre = []
        try:
            with open(CALE_FILTRE, encoding="utf-8") as f:
                filtre = [l.strip() for l in f if l.strip()]
        except FileNotFoundError:
            pass

        self.articole = self.csv_reader.citeste(col_art, col_cant, filtre)

        if not self.articole:
            ctypes.windll.user32.MessageBoxW(0, "Nu exista articole dupa filtrare!", "Eroare", 0)
            sys.exit()

        # Limita per sesiune din config.json (0 = fara limita)
        if ARTICOLE_PER_SESIUNE > 0:
            self.articole = self.articole[:ARTICOLE_PER_SESIUNE]

        # Reluare progres
        if os.path.exists(CALE_PROGRESS):
            try:
                with open(CALE_PROGRESS) as f:
                    salvat = int(f.read().strip())
                if 0 < salvat < len(self.articole):
                    r = ctypes.windll.user32.MessageBoxW(0,
                        f"Progres salvat la articolul {salvat+1}.\n\nDA = continua\nNU = inceput",
                        "Reluare", 4)
                    if r == 6:
                        self.start_index = salvat
            except:
                pass

        if not self.confirmare_start():
            sys.exit()

        time.sleep(1.0)
        self.navigare_initiala()
        self.ruleaza()


# ======================================================
# ENTRY POINT
# ======================================================
if __name__ == "__main__":
    OFAutomation().run()
