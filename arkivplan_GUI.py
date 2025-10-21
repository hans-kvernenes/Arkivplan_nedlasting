import tkinter as tk
import subprocess
import sys
import os
import tempfile
import shutil
import runpy
import atexit

# --- KONFIG ---
external_scripts = [
    "arkivplan_login.py",
    "arkivplan_listenedlasting.py",
    "arkivplan_sidelistehenting.py",
    "arkivplan_nedlasting.py",
]

# --- BASE PATHS ---
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    base_path = sys._MEIPASS  # PyInstaller-ekstraksjonsmappe
else:
    base_path = os.path.dirname(os.path.abspath(__file__))  # mappen hvor .py ligger

# --- TEMP MAPPE FOR KOPIER (behøves hvis du vil jobbe på midlertidige filer) ---
temp_dir = tempfile.mkdtemp()
atexit.register(lambda: shutil.rmtree(temp_dir, ignore_errors=True))

for script in external_scripts:
    src = os.path.join(base_path, script)
    dst = os.path.join(temp_dir, script)
    if os.path.exists(src):
        shutil.copy(src, dst)

def _run_script_inproc(script_path: str):
    """Kjør et Python-skript i samme prosess (blokkerende)."""
    runpy.run_path(script_path, run_name="__main__")

def kjør_skript(skript_navn: str):
    script_path = os.path.join(temp_dir, skript_navn)
    if getattr(sys, "frozen", False):
        # Start samme exe med en "dispatcher"-switch
        subprocess.Popen([sys.executable, "--run-script", skript_navn])
    else:
        # Lokalt utviklingsmiljø: kjør via python-tolkeren
        subprocess.Popen([sys.executable, script_path])

def generer_pdfs_med_forrutgående_henting():
    kjør_skript("arkivplan_sidelistehenting.py")
    kjør_skript("arkivplan_nedlasting.py")

def oppdater_knappetilstand():
    finnes_auth = os.path.exists("auth.json")
    knapp_lastned.config(state=tk.NORMAL if finnes_auth else tk.DISABLED)
    knapp_generer.config(state=tk.NORMAL if finnes_auth else tk.DISABLED)
    vindu.after(5000, oppdater_knappetilstand)

# --- DISPATCHER-MODUS: Kjør et delskript og avslutt ---
# Dette gjør at når du starter exe-en med "--run-script X.py", så kjører den X.py
# uten å åpne GUI. Slik unngår du "ny GUI-instans"-problemet.
if "--run-script" in sys.argv:
    try:
        idx = sys.argv.index("--run-script")
        target = sys.argv[idx + 1]
    except (ValueError, IndexError):
        sys.exit("Feil: --run-script krever et skriptnavn")
    target_path = os.path.join(temp_dir, target)
    _run_script_inproc(target_path)
    sys.exit(0)

# --- GUI ---
vindu = tk.Tk()
vindu.title("Arkivplan Verktøy")

knapp_login = tk.Button(vindu, text="Logg inn på SharePoint", width=40,
                        command=lambda: kjør_skript("arkivplan_login.py"))
knapp_login.pack(pady=5)

knapp_lastned = tk.Button(vindu, text="Last ned lister til Excel", width=40,
                          command=lambda: kjør_skript("arkivplan_listenedlasting.py"))
knapp_lastned.pack(pady=5)

knapp_generer = tk.Button(vindu, text="Last ned sider til PDF", width=40,
                          command=generer_pdfs_med_forrutgående_henting)
knapp_generer.pack(pady=5)

oppdater_knappetilstand()
vindu.mainloop()