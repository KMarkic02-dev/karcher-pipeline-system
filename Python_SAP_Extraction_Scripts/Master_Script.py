import os
import sys
import time
import subprocess
from datetime import datetime
import gspread

# --- SANITIZED: Local Root Folder & Script Directory ---
korensk_mapa = r"C:\Users\YOUR_USERNAME\Path\To\Project"
mapa_skript = os.path.join(korensk_mapa, "GoogleSheetsScrypt")
mapa_logov = os.path.join(korensk_mapa, "Logs")

if not os.path.exists(mapa_logov):
    os.makedirs(mapa_logov)

datum_danes = time.strftime("%d-%m-%Y")
log_datoteka = os.path.join(mapa_logov, f"Log_{datum_danes}.txt")

skripte = [
    "Izvoz_Zaloge.py",
    "Izvoz_Kupcev.py",
    "Izvoz_Artiklov.py",
    "Izvoz_Delivery_Daily.py"
]

print("==================================================")
print("ZAGON MASTER SKRIPTE ZA OSVEŽEVANJE SAP PODATKOV")
print("==================================================\n")

zacetni_cas_master = time.time()
zacetni_cas_str = time.strftime('%d-%m-%Y %H:%M:%S')

log_rezultati = []
uspesne_skripte = 0  # Števec uspešno izvedenih skript

for i, skripta in enumerate(skripte, start=1):
    pot_do_skriptne_datoteke = os.path.join(mapa_skript, skripta)
    
    print(f"[{i}/{len(skripte)}] Zaženem skripto: {skripta}...")
    
    if not os.path.exists(pot_do_skriptne_datoteke):
        sporocilo_napake = f"{skripta} - Error: Script file missing"
        print(f" Datoteka ne obstaja v {mapa_skript}!\n")
        log_rezultati.append(sporocilo_napake)
        continue

    try:
        rezultat = subprocess.run(
            [sys.executable, pot_do_skriptne_datoteke],
            check=True,
            capture_output=True,
            text=True
        )
        status_ok = f"{skripta} uspešno zaključena."
        print(f" {status_ok}")
        log_rezultati.append(status_ok)
        uspesne_skripte += 1

    except subprocess.CalledProcessError as e:
        razlog = "Unknown system error"
        if e.stderr:
            for vrstica in e.stderr.splitlines():
                if "REASON:" in vrstica:
                    razlog = vrstica.replace("REASON:", "").strip()
                    break

        status_err = f"{skripta} - Error: {razlog}"
        
        print(f"\n NAPAKA PRI IZVEDBI SKRIPTE: {skripta}")
        print(f"Razlog: {razlog}")
        print("--> Prehajam na naslednjo skripto...\n")
        
        log_rezultati.append(status_err)

    if i < len(skripte):
        print(" Čakam 10 sekund pred naslednjim izvozom...\n")
        time.sleep(10)

koncni_cas_sec = time.time()
koncni_cas_str = time.strftime('%d-%m-%Y %H:%M:%S')
skupno_trajanje = koncni_cas_sec - zacetni_cas_master

print("\n==================================================")
print(f" VSI PROCESI ZAKLJUČENI! Skupno trajanje: {skupno_trajanje:.2f} s")
print("==================================================")
ima_napako = any("Error:" in status for status in log_rezultati)

try:
    print("\n Posodabljam status v Google Sheet...")

    pot_do_json = os.path.join(korensk_mapa, "Credentials", "credentials.json")
    gc = gspread.service_account(filename=pot_do_json)

    # --- SANITIZED: Google Sheet URL ---
    url_tabele = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit"
    sh = gc.open_by_url(url_tabele)
    status_sheet = sh.worksheet("Status")

    trenutni_cas_prikaz = datetime.now().strftime("%d.%m.%Y %H:%M")

    if ima_napako or uspesne_skripte == 0:
        # Ob napaki zapiše ERROR in datum/čas
        vrednost_za_zapis = f"ERROR: {trenutni_cas_prikaz}"
        print(f" ⚠️ Zapisujem napako v Google Sheet: {vrednost_za_zapis}")
    else:
        # Ob uspehu zapiše samo datum/čas
        vrednost_za_zapis = trenutni_cas_prikaz
        print(f" Datum uspešno zapisan v Google Sheet: {vrednost_za_zapis}")

    # Zapis v celico A1
    status_sheet.update_acell("A1", vrednost_za_zapis)

except Exception as e:
    print(f" Napaka pri zapisu v Google Sheet: {e}")