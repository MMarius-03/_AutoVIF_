import csv
import time
import ctypes
import sys
import threading
import keyboard
import os
import subprocess
import traceback
import pyperclip
from datetime import datetime
from pywinauto import Application

try:
    import pyautogui

    SCREENSHOT_OK = True
except:
    SCREENSHOT_OK = False

# ======================================================
# SETARI
# ======================================================
CALE_CSV = "data\\of_articole.csv"
CALE_FILTRE = "data\\filtre.txt"
CALE_PROGRESS = "data\\progress_putty.txt"
TITLU_PUTTY = "VIF5_7"
CALE_PUTTY = r"C:\vifout\Putty\putty.exe"
SESIUNE_PUTTY = "srvvif57"
PAUZA_DUPA_COD = 2.5  # fallback daca citirea ecranului esueaza
PAUZA_DUPA_F2 = 4.5  # fallback
TIMEOUT_ARTICOL = 60
LOG_FILE = "data\\log.txt"
SCREENSHOT_DIR = "data\\screenshots"
# ======================================================

# VITEZA — modifica doar aici
INTERVAL_TASTARE = 0.13  # viteza intre taste
PAUZA_DUPA_TASTA = 0.9  # pauza dupa fiecare tasta
PAUZA_CITIRE_ECRAN = 0.8  # pauza citire ecran
INTERVAL_SCANARE = 0.4  # cat de des verifica ecranul

os.makedirs(SCREENSHOT_DIR, exist_ok=True)


# -------------------------------------------------------
# LOG + SCREENSHOT
# -------------------------------------------------------
def log(mesaj, idx=None):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linie = f"[{ts}]" + (f" [Articol {idx + 1}]" if idx is not None else "") + f" {mesaj}"
    print(linie)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(linie + "\n")


def screenshot(nume="eroare"):
    if SCREENSHOT_OK:
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            cale = os.path.join(SCREENSHOT_DIR, f"{nume}_{ts}.png")
            pyautogui.screenshot(cale)
            log(f"Screenshot: {cale}")
        except:
            pass


# -------------------------------------------------------
# STOP — manual sau automat
# -------------------------------------------------------
stop_flag = threading.Event()


def ascultaTaste():
    while not stop_flag.is_set():
        if keyboard.is_pressed("esc") or keyboard.is_pressed("f1"):
            stop_flag.set()
        time.sleep(0.1)


threading.Thread(target=ascultaTaste, daemon=True).start()


def opreste(idx, motiv):
    with open(CALE_PROGRESS, "w") as f:
        f.write(str(idx))
    screenshot(f"stop_{idx + 1}")
    log(f"OPRIT — {motiv}", idx)
    ctypes.windll.user32.MessageBoxW(0,
                                     f"Programul s-a oprit!\n\n"
                                     f"Motiv: {motiv}\n"
                                     f"Articol: {idx + 1}\n\n"
                                     f"Progres salvat. Log + screenshot in folderul data\\.",
                                     "Oprit", 0)
    sys.exit()


def verificaStop(idx):
    if stop_flag.is_set():
        opreste(idx, "Oprit manual (ESC sau F1)")


def verificaPutty(idx):
    try:
        app.window(title=TITLU_PUTTY).exists()
    except:
        opreste(idx, "Fereastra PuTTY a disparut!")


# -------------------------------------------------------
# Deschidere PuTTY
# -------------------------------------------------------
log("Deschid PuTTY...")
subprocess.Popen([CALE_PUTTY, "-load", SESIUNE_PUTTY])
time.sleep(4.5)

log("Caut fereastra...")
try:
    app = Application(backend="win32").connect(title=TITLU_PUTTY)
    vif = app.window(title=TITLU_PUTTY)
    log(f"PuTTY gasit: {vif.window_text()}")
except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f"Nu am gasit fereastra PuTTY!\n\nEroare: {e}", "Eroare", 0)
    sys.exit()


def get_putty():
    try:
        w = app.window(title=TITLU_PUTTY)
        w.set_focus()
        return w
    except:
        time.sleep(1)
        w = app.window(title=TITLU_PUTTY)
        w.set_focus()
        return w


def apasa(tasta, pauza=PAUZA_DUPA_TASTA):
    get_putty().type_keys(tasta)
    time.sleep(pauza)


def scrie(text, pauza=PAUZA_DUPA_TASTA):
    get_putty().type_keys(text, with_spaces=True, pause=INTERVAL_TASTARE)
    time.sleep(pauza)


# -------------------------------------------------------
# CITIRE ECRAN + ASTEPTARE TEXT DINAMIC
# -------------------------------------------------------
def citeste_ecran():
    pyperclip.copy('')
    get_putty().type_keys("%{SPACE}o")
    time.sleep(PAUZA_CITIRE_ECRAN)
    return pyperclip.paste()


def asteapta_text(text_cautat, timeout=15, idx=None, mesaj_eroare=None):
    """Asteapta pana apare textul pe ecran. Returneaza True/False."""
    t_start = time.time()
    while time.time() - t_start < timeout:
        if stop_flag.is_set():
            return False
        ecran = citeste_ecran()
        if text_cautat in ecran:
            return True
        # Verificare erori cunoscute pe ecran
        if "inexistent" in ecran.lower() or "eroare" in ecran.lower():
            log(f"Eroare detectata pe ecran!", idx)
            screenshot(f"eroare_ecran_{idx + 1 if idx is not None else 0}")
            return False
        time.sleep(INTERVAL_SCANARE)
    if mesaj_eroare:
        log(f"TIMEOUT: {mesaj_eroare}", idx)
    return False


# -------------------------------------------------------
# Detectare coloane din CSV
# (fisierul poate avea randuri de metadata inainte de header)
# -------------------------------------------------------
def curata_valoare(val):
    """Curata formatul Excel =\"...\" si spatii"""
    if val is None:
        return ""
    val = val.strip()
    if val.startswith("=\"") and val.endswith("\""):
        val = val[2:-1]
    return val


def gaseste_header_si_date(cale):
    """Cauta randul cu Articol ca header si returneaza coloanele + primul rand de date"""
    with open(cale, encoding="utf-8") as f:
        linii = f.readlines()

    header_idx = None
    for i, linie in enumerate(linii):
        if "Articol" in linie and "Denumire" in linie:
            header_idx = i
            break

    if header_idx is None:
        return None, None, None

    # Citim cu DictReader incepand de la randul header
    import io
    continut = "".join(linii[header_idx:])
    reader = csv.DictReader(io.StringIO(continut), delimiter=";")
    coloane = reader.fieldnames
    primul = next(reader, None)
    return coloane, primul, header_idx


try:
    coloane_csv, primul_rand, header_idx = gaseste_header_si_date(CALE_CSV)
    if coloane_csv is None:
        ctypes.windll.user32.MessageBoxW(0, "Nu am gasit randul cu \'Articol\' in CSV!", "Eroare", 0)
        sys.exit()
    # Curatam valorile primului rand pentru preview
    primul_rand = {k: curata_valoare(v) for k, v in primul_rand.items()}
except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f"Eroare la citirea CSV:\n{e}", "Eroare", 0)
    sys.exit()

if not coloane_csv:
    ctypes.windll.user32.MessageBoxW(0, "CSV-ul nu are coloane!", "Eroare", 0)
    sys.exit()

# Fereastra selectie coloane
import tkinter as tk
from tkinter import ttk

col_articol_var = [None]
col_cantitate_var = [None]


def on_confirma_coloane():
    if not col_articol_var[0] or not col_cantitate_var[0]:
        tk.messagebox.showerror("Eroare", "Selecteaza ambele coloane!")
        return
    if col_articol_var[0] == col_cantitate_var[0]:
        tk.messagebox.showerror("Eroare", "Coloanele trebuie sa fie diferite!")
        return
    root_col.destroy()


def on_anuleaza_coloane():
    sys.exit()


root_col = tk.Tk()
root_col.title("Selectie coloane CSV")
root_col.geometry("500x400")
root_col.resizable(False, False)

tk.Label(root_col, text="Selecteaza coloanele din fisierul CSV:",
         font=("Arial", 11, "bold"), pady=10).pack()

# Preview CSV
frame_prev = tk.LabelFrame(root_col, text="Preview CSV (primul rand)", padx=5, pady=5)
frame_prev.pack(fill=tk.X, padx=15, pady=5)

for col in coloane_csv:
    val = primul_rand.get(col, "")
    tk.Label(frame_prev, text=f"{col}: {val}", anchor="w",
             font=("Courier", 9)).pack(fill=tk.X)

# Selectie coloane
frame_sel = tk.LabelFrame(root_col, text="Alege coloanele", padx=10, pady=10)
frame_sel.pack(fill=tk.X, padx=15, pady=10)

tk.Label(frame_sel, text="Coloana COD ARTICOL:").grid(row=0, column=0, sticky="w", pady=5)
combo_articol = ttk.Combobox(frame_sel, values=coloane_csv, state="readonly", width=25)
combo_articol.grid(row=0, column=1, padx=10)
# Auto-selectare daca gasim coloana potrivita
for c in coloane_csv:
    if "articol" in c.lower() or "cod" in c.lower():
        combo_articol.set(c)
        col_articol_var[0] = c
        break

tk.Label(frame_sel, text="Coloana CANTITATE:").grid(row=1, column=0, sticky="w", pady=5)
combo_cant = ttk.Combobox(frame_sel, values=coloane_csv, state="readonly", width=25)
combo_cant.grid(row=1, column=1, padx=10)
# Auto-selectare daca gasim coloana potrivita
for c in coloane_csv:
    if any(x in c.lower() for x in ["cant", "ctq", "ct,", "cantitate", "qty", "quantity"]):
        combo_cant.set(c)
        col_cantitate_var[0] = c
        break


def on_select_articol(e):
    col_articol_var[0] = combo_articol.get()


def on_select_cant(e):
    col_cantitate_var[0] = combo_cant.get()


combo_articol.bind("<<ComboboxSelected>>", on_select_articol)
combo_cant.bind("<<ComboboxSelected>>", on_select_cant)

# Butoane
btn_frame_col = tk.Frame(root_col)
btn_frame_col.pack(pady=10)
tk.Button(btn_frame_col, text="✓ Confirma", command=on_confirma_coloane,
          bg="#2ecc71", fg="white", font=("Arial", 10, "bold"),
          width=12, height=2).pack(side=tk.LEFT, padx=10)
tk.Button(btn_frame_col, text="✗ Anuleaza", command=on_anuleaza_coloane,
          bg="#e74c3c", fg="white", font=("Arial", 10, "bold"),
          width=12, height=2).pack(side=tk.LEFT, padx=10)

root_col.mainloop()

COL_ARTICOL = col_articol_var[0]
COL_CANTITATE = col_cantitate_var[0]
log(f"Coloane selectate: Articol='{COL_ARTICOL}', Cantitate='{COL_CANTITATE}'")

# -------------------------------------------------------
# Citire filtre
# -------------------------------------------------------
filtre = []
try:
    with open(CALE_FILTRE, encoding="utf-8") as f:
        filtre = [line.strip() for line in f if line.strip()]
except FileNotFoundError:
    ctypes.windll.user32.MessageBoxW(0, f"Nu am gasit fisierul de filtre:\n{CALE_FILTRE}", "Eroare", 0)
    sys.exit()

# -------------------------------------------------------
# Citire CSV (sarind peste randurile de metadata)
# -------------------------------------------------------
articole = []
try:
    import io

    with open(CALE_CSV, encoding="utf-8") as f:
        linii = f.readlines()
    continut = "".join(linii[header_idx:])
    reader = csv.DictReader(io.StringIO(continut), delimiter=";")
    for row in reader:
        cod = curata_valoare(row.get(COL_ARTICOL, ""))
        cant = curata_valoare(row.get(COL_CANTITATE, ""))
        if not cod:
            continue
        # Convertim corect cantitatea indiferent de nr de zecimale (6.000, 4.500, 18.00 etc.)
        try:
            val = float(cant.replace(",", "."))
            if val == int(val):
                cant = str(int(val))  # 6.000 → "6", 56.000 → "56"
            else:
                cant = f"{val:.3f}".rstrip("0").rstrip(".").replace(".", ",")  # 4.500 → "4,5"
        except:
            cant = cant.replace(".", ",")
        articole.append((cod, cant))
except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f"Eroare la citirea CSV:\n{e}", "Eroare", 0)
    sys.exit()

if not articole:
    ctypes.windll.user32.MessageBoxW(0, "Nu exista date in fisier!", "Eroare", 0)
    sys.exit()

# -------------------------------------------------------
# FILTRARE
# -------------------------------------------------------
articole = [(cod, cant) for cod, cant in articole if not cod.startswith(tuple(filtre))]

if not articole:
    ctypes.windll.user32.MessageBoxW(0, "Nu au ramas articole dupa filtrare!", "Eroare", 0)
    sys.exit()

# LIMITARE TEST — comenteaza linia pt run complet
articole = articole[:110]

# -------------------------------------------------------
# RELUARE DE UNDE S-A OPRIT
# -------------------------------------------------------
start_index = 0
if os.path.exists(CALE_PROGRESS):
    try:
        with open(CALE_PROGRESS, "r") as f:
            salvat = int(f.read().strip())
        if 0 < salvat < len(articole):
            raspuns = ctypes.windll.user32.MessageBoxW(0,
                                                       f"Gasit progres salvat la articolul {salvat + 1} din {len(articole)}.\n\n"
                                                       f"DA = continua de la articolul {salvat + 1}\n"
                                                       f"NU = incepe de la inceput",
                                                       "Reluare progres", 4)
            if raspuns == 6:
                start_index = salvat
    except:
        pass

# -------------------------------------------------------
# Fereastra cu lista articolelor — confirmare start
# -------------------------------------------------------

confirmat = [False]


def on_da():
    confirmat[0] = True
    root.destroy()


def on_nu():
    root.destroy()


root = tk.Tk()
root.title(f"Articole de introdus ({len(articole) - start_index} total)")
root.geometry("500x500")
root.resizable(True, True)

tk.Label(root, text=f"Urmeaza sa fie introduse {len(articole) - start_index} articole:",
         font=("Arial", 10, "bold"), pady=5).pack()

frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

cols = ("Nr", "Cod Articol", "Cantitate")
tree = ttk.Treeview(frame, columns=cols, show="headings", height=18)
tree.heading("Nr", text="Nr")
tree.heading("Cod Articol", text="Cod Articol")
tree.heading("Cantitate", text="Cantitate")
tree.column("Nr", width=50, anchor="center")
tree.column("Cod Articol", width=200, anchor="center")
tree.column("Cantitate", width=150, anchor="center")

scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
tree.pack(fill=tk.BOTH, expand=True)

for idx, (cod, cant) in enumerate(articole[start_index:], start=1):
    tree.insert("", tk.END, values=(idx, cod, cant))

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)
tk.Button(btn_frame, text="✓ DA — Incepe", command=on_da,
          bg="#2ecc71", fg="white", font=("Arial", 11, "bold"),
          width=15, height=2).pack(side=tk.LEFT, padx=10)
tk.Button(btn_frame, text="✗ NU — Anuleaza", command=on_nu,
          bg="#e74c3c", fg="white", font=("Arial", 11, "bold"),
          width=15, height=2).pack(side=tk.LEFT, padx=10)

root.mainloop()

if not confirmat[0]:
    log("Anulat de utilizator.")
    sys.exit()

time.sleep(1.0)

# =======================================================
# FLUX INITIAL — navigare la p04a
# =======================================================
log("Navigare la p04a...")

# Pas 1 — login SSH: asteptam "login as:" si scriem p04
if not asteapta_text("login as:", timeout=15, mesaj_eroare="Nu a aparut promptul login as:"):
    opreste(start_index, "Nu s-a putut conecta la server (login as: nu a aparut)")
scrie("p04")
apasa("{ENTER}")

# Pas 2 — ecranul cu Locatie: 1 → schimbam in 2
if not asteapta_text("Locatie:", timeout=15, mesaj_eroare="Nu a aparut ecranul cu Locatie"):
    opreste(start_index, "Nu s-a incarcat ecranul de login VIF")
apasa("{TAB}")
apasa("{TAB}")
scrie("2")
apasa("{ENTER}")
apasa("{ENTER}")

# Pas 3 — meniu principal: asteptam "Fabricatie" si dam RIGHT RIGHT
if not asteapta_text("Fabricatie", timeout=15, mesaj_eroare="Nu a aparut meniul principal"):
    opreste(start_index, "Nu s-a incarcat meniul principal VIF")
apasa("{RIGHT}")
apasa("{RIGHT}")
apasa("{ENTER}")
apasa("{ENTER}")

# Pas 4 — ecranul "Post de lucru": scriem p04a
if not asteapta_text("Post de lucru", timeout=15, mesaj_eroare="Nu a aparut ecranul Post de lucru"):
    opreste(start_index, "Nu s-a incarcat ecranul Post de lucru")
scrie("p04a")
apasa("{ENTER}")

# Pas 5 — asteptam SELECTIE OF cu CTRL-O=Nou OF
if not asteapta_text("CTRL-O=Nou OF", timeout=15, mesaj_eroare="Nu s-a incarcat ecranul SELECTIE OF"):
    opreste(start_index, "Nu s-a putut naviga la ecranul SELECTIE OF")
log("Ecran SELECTIE OF incarcat — gata de introdus articole!")

# =======================================================
# BUCLA PRINCIPALA
# =======================================================
for i in range(start_index, len(articole)):
    verificaStop(i)
    verificaPutty(i)

    cod, cant = articole[i]
    log(f"Introducere: {cod} - {cant}", i)

    try:
        # Ctrl+O — deschide CREAREA UNUI OF
        apasa("^o")
        if not asteapta_text("CREAREA UNUI OF", timeout=10, idx=i,
                             mesaj_eroare="Nu s-a deschis ecranul CREAREA UNUI OF"):
            opreste(i, f"Ecranul CREAREA UNUI OF nu s-a incarcat pentru articolul {cod}")

        # Cod articol
        scrie(cod)
        apasa("{ENTER}")

        # Asteptam ca articolul sa fie recunoscut — codul apare pe ecran
        if not asteapta_text(cod, timeout=10, idx=i,
                             mesaj_eroare=f"Articolul {cod} nu a fost recunoscut de server"):
            opreste(i, f"Articolul {cod} nu a fost gasit in sistem")

        # Cantitate
        scrie(cant)
        apasa("{ENTER}")

        # Salvare F2 — asteptam INTROD CANTITATILOR
        apasa("{F2}")
        if not asteapta_text("INTROD CANTITATILOR DE INTRARE", timeout=15, idx=i,
                             mesaj_eroare="Nu s-a salvat OF-ul"):
            opreste(i, f"Salvarea OF pentru articolul {cod} a esuat")

        # F4 — inchide OF, apare dialog confirmare
        apasa("{F4}")
        if asteapta_text("Sunteti sigur", timeout=5):
            apasa("{ENTER}")  # Confirmam Da

        # Asteptam sa revenim la SELECTIE OF pentru urmatorul articol
        if not asteapta_text("CTRL-O=Nou OF", timeout=10, idx=i,
                             mesaj_eroare="Nu s-a revenit la ecranul SELECTIE OF"):
            opreste(i, f"Nu s-a putut reveni la SELECTIE OF dupa articolul {cod}")

        log(f"OK: {cod}", i)

    except Exception as e:
        log(f"EROARE: {traceback.format_exc()}", i)
        opreste(i, f"Eroare neasteptata: {str(e)}")

# -------------------------------------------------------
# Final
# -------------------------------------------------------
if os.path.exists(CALE_PROGRESS):
    os.remove(CALE_PROGRESS)

log("Procesare finalizata cu succes!")
ctypes.windll.user32.MessageBoxW(0,
                                 f"Gata! S-au introdus {len(articole) - start_index} articole.",
                                 "Finalizat", 0)