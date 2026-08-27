import os
import sys
import time
import datetime
import win32com.client
import pandas as pd
import gspread
import pyautogui
import pyperclip


def pridobi_sap_sejo(zahtevan_sistem="PG1"):
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine   
        for conn in application.Children:
            for sess in conn.Children:
                sys_name = str(sess.Info.SystemName).upper()
                app_server = str(sess.Info.ApplicationServer).lower()

                if zahtevan_sistem.upper() in sys_name or zahtevan_sistem.lower() in app_server:
                    print(f"[OK] SAP session found for system: {sys_name} ({sess.Info.ApplicationServer})")
                    return sess

        print("REASON: SAP session not open", file=sys.stderr)
        sys.exit(1)

    except Exception:
        print("REASON: SAP GUI not running", file=sys.stderr)
        sys.exit(1)


# --- GLAVNI PROGRAM ---
try:
    print("1. Starting script for 2-hour Delivery export (ZDE1SD131)...")
    zacetni_cas = time.time()

    session = pridobi_sap_sejo("PG1")

    # 1. Navigacija in izvoz iz SAP
    try:
        print("--> Navigating to transaction ZDE1SD131...")
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZDE1SD131"
        session.findById("wnd[0]").sendVKey(0)

        time.sleep(3)

        print("--> Opening variant selection window...")

        try:
            session.findById("wnd[0]/tbar[1]/btn[17]").press()
        except Exception:
            session.findById("wnd[0]").sendVKey(17)

        time.sleep(2)

        # --- SANITIZED: SAP User / Variant Name ---
        session.findById("wnd[1]/usr/txtENAME-LOW").text = "YOUR_SAP_USER"
        session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
        session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        time.sleep(2)

        danes = datetime.date.today()
        zacetni_datum = danes - datetime.timedelta(days=60)

        danes_str = danes.strftime("%d.%m.%Y")
        zacetni_datum_str = zacetni_datum.strftime("%d.%m.%Y")

        print(f"--> Nastavljam obdobje izvoza: {zacetni_datum_str} do {danes_str}")
        session.findById("wnd[0]/usr/ctxtSO_ERDAT-LOW").text = zacetni_datum_str
        session.findById("wnd[0]/usr/ctxtSO_ERDAT-HIGH").text = danes_str

        print("--> Nastavljam statusne parametre (a - c)...")
        session.findById("wnd[0]/usr/ctxtSO_LFSTK-LOW").text = "a"
        session.findById("wnd[0]/usr/ctxtSO_LFSTK-HIGH").text = "c"
        session.findById("wnd[0]/usr/ctxtSO_LFSTA-LOW").text = "a"
        session.findById("wnd[0]/usr/ctxtSO_LFSTA-HIGH").text = "c"
        session.findById("wnd[0]/usr/ctxtSO_LFSTA-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtSO_LFSTA-LOW").caretPosition = 1

        time.sleep(1)

        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(4)

        print("3. Loading variant/layout and selecting columns...")
        shell = session.findById("wnd[0]/usr/cntlALV/shellcont/shell")

        shell.pressToolbarContextButton("&MB_VARIANT")
        shell.selectContextMenuItem("&LOAD")
        time.sleep(1.5)

        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").setCurrentCell(65, "TEXT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "65"
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        time.sleep(1.5)

        shell.currentCellColumn = "ARKTX"
        shell.selectColumn("BSTDK")
        shell.selectColumn("ZZ_WW_DEBITOR")
        shell.selectColumn("NAME1_WE")
        shell.selectColumn("ZZ_AUFTRAGNR")
        shell.selectColumn("BSTNK")
        shell.selectColumn("MATNR")
        shell.selectColumn("ARKTX")
        shell.selectColumn("BMENG")
        shell.selectColumn("WMENG")
        shell.selectColumn("VDATU")
        shell.selectColumn("GBSTA")

        shell.contextMenu()
        shell.selectContextMenuItem("&XXL")
        time.sleep(1.5)

        session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # --- SANITIZED: Local File Paths ---
        korensk_mapa = r"C:\Users\YOUR_USERNAME\Path\To\Project"
        ciljna_mapa = os.path.join(korensk_mapa, "Export")
        ciljno_ime = "Delivery_Export.xlsx"
        polna_pot = os.path.join(ciljna_mapa, ciljno_ime)

        if os.path.exists(polna_pot):
            try:
                os.remove(polna_pot)
            except Exception:
                pass

        time.sleep(2)

        try:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = ciljna_mapa
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ciljno_ime
            
            try:
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
            except Exception:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                
            print("--> Saved via standard SAP GUI window.")

        except Exception:
            print("--> Standard SAP window not detected, using Windows paste...")
            
            try:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except Exception:
                pass

            time.sleep(1.5)
            pyperclip.copy(polna_pot)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            pyautogui.press('enter')

    except Exception:
        print("REASON: SAP GUI step failed", file=sys.stderr)
        sys.exit(1)

    # 2. Preverjanje in branje datoteke z diska
    print("5. Waiting for the file to be written to disk...")
    max_cakanje = 30
    zacetek_cakanja = time.time()

    while not os.path.exists(polna_pot):
        time.sleep(1)
        if time.time() - zacetek_cakanja > max_cakanje:
            print("REASON: Excel file missing", file=sys.stderr)
            sys.exit(1)

    time.sleep(2)
    print(f"--> Export successfully saved to: {polna_pot}")

    try:
        shell.setCurrentCell(13, "ZZ_AUFTRAGNR")
        shell.clearSelection()
    except Exception:
        pass

    # 3. Obdelava podatkov s Pandas
    print("6. Processing data with Pandas...")
    excel_pot = os.path.join(ciljna_mapa, ciljno_ime)

    try:
        df = pd.read_excel(excel_pot, dtype=str)
        df = df.fillna("")

        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()

        for i in range(len(df)):
            mat = str(df.iat[i, 5]).strip()
            if "-" in mat and not mat.endswith(".0") and mat not in ["", "nan", "None"]:
                df.iat[i, 5] = mat + ".0"

        # --- SANITIZED: Specific internal location mapping ---
        manjkajoce_stevilke = {
            "Location Name 1": "1000000001",
            "Location Name 2": "1000000002",
            "Location Name 3": "1000000003",
        }

        for i in range(len(df)):
            stevilka = str(df.iat[i, 1]).strip()
            ime_centra = str(df.iat[i, 2]).strip()
            
            if stevilka in ["", "nan", "None"]:
                if ime_centra in manjkajoce_stevilke:
                    df.iat[i, 1] = manjkajoce_stevilke[ime_centra]

        podatki_za_prenos = df.values.tolist()

    except Exception:
        print("REASON: Excel read error", file=sys.stderr)
        sys.exit(1)

    # 4. Google Sheets osveževanje
    print("7. Connecting to Google Sheets...")
    pot_do_json = os.path.join(korensk_mapa, "Credentials", "credentials.json")
    if not os.path.exists(pot_do_json):
        print("REASON: Google credentials missing", file=sys.stderr)
        sys.exit(1)

    try:
        gc = gspread.service_account(filename=pot_do_json)

        # --- SANITIZED: Google Sheet URL & Worksheet GID ---
        url_tabele = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit"
        sh = gc.open_by_url(url_tabele)

        ciljni_gid = "0000000000"  # Placeholder Worksheet GID
        worksheet = None
        for ws in sh.worksheets():
            if str(ws.id) == ciljni_gid:
                worksheet = ws
                break

        if not worksheet:
            worksheet = sh.get_worksheet(0)

        koncni_podatki_za_zapis = []

        if podatki_za_prenos:
            print("--> Reading existing data from Google Sheets...")
            obstojeci_podatki = worksheet.get_all_values()

            headers = obstojeci_podatki[0] if len(obstojeci_podatki) > 0 else []
            stari_podatki_vrstice = obstojeci_podatki[1:] if len(obstojeci_podatki) > 1 else []

            def normaliziraj_kljuc(sls_doc, mat):
                d = str(sls_doc).strip()
                m = str(mat).strip()
                if m.endswith(".0"):
                    m = m[:-2]
                return (d, m)

            mapa_obstojecih = {}
            for idx, row in enumerate(stari_podatki_vrstice):
                if len(row) >= 6:
                    kljuc = normaliziraj_kljuc(row[3], row[5])
                else:
                    kljuc = tuple(str(x).strip() for x in row[:4])
                mapa_obstojecih[kljuc] = idx

            posodobljeni_seznam = list(stari_podatki_vrstice)
            st_posodobljenih = 0
            st_novih = 0

            for nova_vrstica in podatki_za_prenos:
                nova_vrstica_str = [str(x) for x in nova_vrstica]
                
                if len(nova_vrstica_str) >= 6:
                    kljuc_nove = normaliziraj_kljuc(nova_vrstica_str[3], nova_vrstica_str[5])
                else:
                    kljuc_nove = tuple(nova_vrstica_str[:4])

                if kljuc_nove in mapa_obstojecih:
                    idx = mapa_obstojecih[kljuc_nove]
                    posodobljeni_seznam[idx] = nova_vrstica_str
                    st_posodobljenih += 1
                else:
                    posodobljeni_seznam.append(nova_vrstica_str)
                    mapa_obstojecih[kljuc_nove] = len(posodobljeni_seznam) - 1
                    st_novih += 1

            print(f"--> Updated: {st_posodobljenih} existing rows, added: {st_novih} new rows.")
            
            print("--> Sorting data by customer and date...")
            df_koncni = pd.DataFrame(posodobljeni_seznam)
            
            if not df_koncni.empty and df_koncni.shape[1] >= 2:
                df_koncni.sort_values(by=[1, 0], ascending=[True, False], inplace=True)
            
            koncni_podatki_za_zapis = df_koncni.values.tolist()

            print("--> Saving refreshed and sorted data to Google Sheets...")
            
            worksheet.clear()
            if headers:
                worksheet.update("A1", [headers] + koncni_podatki_za_zapis, value_input_option='USER_ENTERED')
            else:
                worksheet.update("A2", koncni_podatki_za_zapis, value_input_option='USER_ENTERED')

    except Exception:
        print("REASON: Google Sheets error", file=sys.stderr)
        sys.exit(1)

    trajanje_sekund = time.time() - zacetni_cas
    print(f"[OK] Daily Delivery updated & sorted in Google Sheets [Timestamp: {time.strftime('%d-%m-%Y %H:%M:%S')}] (Duration: {trajanje_sekund:.2f} s)")

except Exception:
    print("REASON: Unknown system error", file=sys.stderr)
    sys.exit(1)