import os
import sys
import time
import glob
import xml.etree.ElementTree as ET
import gspread

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

print("1. Fetching XML data via Serv-U WebClient...")
zacetni_cas = time.time()

# --- 0. NASTAVITEV MAP ZA PRENOS IN POTI ---
# --- SANITIZED: Local Root Folder ---
korensk_mapa = r"C:\Users\YOUR_USERNAME\Path\To\Project"
download_dir = os.path.join(korensk_mapa, "Downloads")
os.makedirs(download_dir, exist_ok=True)

# Počiščenje starih XML datotek z diska
for f in glob.glob(os.path.join(download_dir, "*.xml")):
    try:
        os.remove(f)
    except Exception:
        pass

# --- 1. PODATKI ZA PRIJAVO ---
# --- SANITIZED: Credentials & URL ---
login_url = "https://your-serv-u-webclient.com/login"
username_val = "YOUR_USERNAME"
password_val = "YOUR_PASSWORD"

podatki_za_prenos = []

# --- NASTAVITEV CHROME BRSKALNIKA ---
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "safebrowsing.disable_download_protection": True,
    "profile.default_content_setting_values.automatic_downloads": 1
}
chrome_options.add_experimental_option("prefs", prefs)

try:
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)
    
    # 1. Prijava na portal
    driver.get(login_url)

    user_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']")))
    user_input.clear()
    user_input.send_keys(username_val)

    pass_input = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
    pass_input.clear()
    pass_input.send_keys(password_val)

    login_btn = driver.find_element(By.CSS_SELECTOR, "button, input[type='submit']")
    login_btn.click()

    # 2. Počakaj na nalaganje datoteke NEWFILE.xml
    wait.until(EC.presence_of_element_located((
        By.XPATH, "//*[contains(text(), 'NEWFILE.xml')]"
    )))
    time.sleep(2)

    # 3. Klikni gumb s tremi pikicami (menu icon) v vrstici z NEWFILE.xml
    three_dots = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//span[contains(text(), 'NEWFILE.xml')]/ancestor::tr//button[contains(@class, 'menu-button') or @aria-label='menu icon']"
    )))
    driver.execute_script("arguments[0].click();", three_dots)
    time.sleep(1)

    # 4. Klikni gumb 'Download' v meniju
    download_btn = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//*[contains(text(), 'Download') or contains(text(), 'Prenesi')]"
    )))
    driver.execute_script("arguments[0].click();", download_btn)

    print("Čakanje na prenos datoteke NEWFILE.xml...")
    time.sleep(10)
    driver.quit()

    # Poišči preneseno XML datoteko
    xml_files = glob.glob(os.path.join(download_dir, "*.xml"))
    if not xml_files:
        raise Exception("XML datoteka se ni uspešno prenesla v mapo Downloads.")

    prenesena_datoteka = xml_files[0]

    # --- 2. PARSE XML DATA IZ DATOTEKE ---
    print("2. Parsing XML structure...")
    
    tree = ET.parse(prenesena_datoteka)
    root = tree.getroot()
    
    # Iskanje vseh <Materials> vrstic v XML datoteki
    for record in root.findall(".//Materials"):
        # Preberi Material in odstrani vodilne ničle (npr. 000000000063711540 -> 63711540)
        material_raw = record.findtext("Material", default="").strip()
        material = material_raw.lstrip("0")
        
        # Preberi Status (npr. Available / Un-Available)
        status = record.findtext("Status", default="").strip()
        
        if material:
            podatki_za_prenos.append([material, status])
            
    print(f" Extracted {len(podatki_za_prenos)} items from XML.")

    # Pobriši lokalno datoteko po obdelavi
    os.remove(prenesena_datoteka)

except Exception as e:
    print(f"REASON: Could not fetch or parse XML ({e})", file=sys.stderr)
    sys.exit(1)


# --- 3. UPLOAD TO GOOGLE SHEETS ---
print("3. Connecting to Google Sheets...")

pot_do_json = os.path.join(korensk_mapa, "Credentials", "credentials.json")

try:
    gc = gspread.service_account(filename=pot_do_json)
    
    # --- SANITIZED: Google Sheet URL ---
    url_tabele = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit"
    sh = gc.open_by_url(url_tabele)
    
    worksheet = sh.worksheet("PGARTIKLI")

    # Pobriši obstoječe podatke od vrstice A2 dalje
    worksheet.batch_clear(["A2:Z10000"])
    
    # Zapiši Material v Stolpec A in Status v Stolpec B
    if podatki_za_prenos:
        worksheet.update("A2", podatki_za_prenos)

except Exception as e:
    print(f"REASON: Google Sheets error ({e})", file=sys.stderr)
    sys.exit(1)

trajanje_sekund = time.time() - zacetni_cas
print(f"[OK] XML Data uploaded to PGARTIKLI tab [Timestamp: {time.strftime('%d-%m-%Y %H:%M:%S')}] (Duration: {trajanje_sekund:.2f} s)")