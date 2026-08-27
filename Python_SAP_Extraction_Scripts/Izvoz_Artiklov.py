import os
import sys
import time
import win32com.client
import pandas as pd
import gspread


def pridobi_sap_sejo(zahtevan_sistem="PE1"):
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        
        for conn in application.Children:
            for sess in conn.Children:
                sys_name = str(sess.Info.SystemName).upper()
                app_server = str(sess.Info.ApplicationServer).lower()
                
                if zahtevan_sistem.upper() in sys_name or zahtevan_sistem.lower() in app_server:
                    print(f" SAP session found for system: {sys_name} ({sess.Info.ApplicationServer})")
                    return sess
                    
        print("REASON: SAP session not open", file=sys.stderr)
        sys.exit(1)
        
    except Exception:
        print("REASON: SAP GUI not running", file=sys.stderr)
        sys.exit(1)


# --- GLAVNI PROGRAM ---
try:
    print("1. Starting script for automatic article refresh from SAP...")
    zacetni_cas = time.time()

    session = pridobi_sap_sejo("PE1")

    # 1. Navigacija in izvoz iz SAP
    try:
        print("--> Navigating to transaction MM60...")
        session.sendCommand("/nMM60")
        time.sleep(3)

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/ctxtMS_WERKS-LOW").text = "SI10"
        session.findById("wnd[0]/usr/ctxtMS_WERKS-LOW").caretPosition = 4

        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(4)

        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        grid.currentCellRow = -1

        grid.setCurrentCell(1, "MTART")

        grid.selectColumn("MATNR")
        grid.selectColumn("WERKS")
        grid.selectColumn("BWTAR")
        grid.selectColumn("KTEXT")
        grid.selectColumn("LAEDA")
        grid.selectColumn("MTART")
        grid.selectColumn("MATKL")
        grid.selectColumn("MEINS")
        grid.selectColumn("EKGRP")
        grid.selectColumn("MAABC")
        grid.selectColumn("DISMM")
        grid.selectColumn("BKLAS")
        grid.selectColumn("VPRSV")
        grid.selectColumn("PREIS")
        grid.selectColumn("WAERS")
        grid.selectedRows = "0"

        grid.contextMenu()
        grid.selectContextMenuItem("&XXL")
        time.sleep(1.5)

        session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
        session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "31"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        time.sleep(2)
        
        # --- CENSORSHIIP / SANITIZATION ---
        korensk_mapa = r"C:\Users\YOUR_USERNAME\Path\To\Project"
        ciljna_mapa = os.path.join(korensk_mapa, "Export")
        ciljno_ime = "Artikli_SI10.xlsx"

        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ciljna_mapa
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ciljno_ime

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1.5)

        try:
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
        except Exception:
            pass

    except Exception:
        print("REASON: SAP GUI step failed", file=sys.stderr)
        sys.exit(1)

    print("5. Export from SAP successful! Reading and preparing data from Excel...")

    # 2. Branje Excel datoteke
    excel_pot = os.path.join(ciljna_mapa, ciljno_ime)
    if not os.path.exists(excel_pot):
        print("REASON: Excel file missing", file=sys.stderr)
        sys.exit(1)

    try:
        df = pd.read_excel(excel_pot)
        df = df.fillna("").astype(str)
        podatki_za_prenos = df.values.tolist()
    except Exception:
        print("REASON: Excel read error", file=sys.stderr)
        sys.exit(1)

    # 3. Google Sheets posodobitev
    print("6. Connecting to Google Sheets...")
    pot_do_json = os.path.join(korensk_mapa, "Credentials", "credentials.json")
    if not os.path.exists(pot_do_json):
        print("REASON: Google credentials missing", file=sys.stderr)
        sys.exit(1)

    try:
        gc = gspread.service_account(filename=pot_do_json)

        # --- CENSORSHIIP / SANITIZATION ---
        url_tabele = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit"
        sh = gc.open_by_url(url_tabele)

        ciljni_gid = 0000000000  # Placeholder GID
        worksheet = None
        for ws in sh.worksheets():
            if ws.id == ciljni_gid:
                worksheet = ws
                break

        if not worksheet:
            worksheet = sh.get_worksheet(0)

        worksheet.batch_clear(["A2:Z10000"])

        if podatki_za_prenos:
            worksheet.update("A2", podatki_za_prenos)

    except Exception:
        print("REASON: Google Sheets error", file=sys.stderr)
        sys.exit(1)

    trajanje_sekund = time.time() - zacetni_cas
    print(f" [OK] Google Sheets updated successfully! [Timestamp: {time.strftime('%d-%m-%Y %H:%M:%S')}] (Duration: {trajanje_sekund:.2f} s)")

except Exception:
    print("REASON: Unknown system error", file=sys.stderr)
    sys.exit(1)