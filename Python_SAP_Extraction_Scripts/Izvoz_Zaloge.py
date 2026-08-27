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
    print("1. Starting script for MB52 - Local Storage ...")
    zacetni_cas = time.time()

    session = pridobi_sap_sejo("PE1")

    # 1. Navigacija in izvoz iz SAP
    try:
        session.sendCommand("/nMB52")
        time.sleep(3)

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtMATNR-LOW").caretPosition = 3
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "SI10"
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").caretPosition = 4
        session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtLGORT-LOW").caretPosition = 3
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        time.sleep(4)

        print("2. Transaction opened successfully...")

        session.findById("wnd[0]").maximize()
        
        # Odpre okno za izbiro postavitve (Layout)
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        time.sleep(1.5)

        # Poišče in izbere layout /SI10 EXPORT ali SI10 stocks
        layout_grid = session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
        
        target_row = -1
        
        # Pregledamo vse vrstice v tabeli postavitve
        for row in range(layout_grid.rowCount):
            variant_val = layout_grid.getCellValue(row, "VARIANT")
            text_val = layout_grid.getCellValue(row, "TEXT")
            
            if "/SI10 EXPORT" in variant_val or "SI10 stocks" in text_val:
                target_row = row
                break

        if target_row >= 0:
            layout_grid.currentCellRow = target_row
            layout_grid.selectedRows = str(target_row)
            layout_grid.clickCurrentCell()
        else:
            print("OPOZORILO: Layout /SI10 EXPORT ni bil najden, uporablja se trenutni.")

        time.sleep(1.5)

        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")

        # Odpre meni za izvoz v Excel
        try:
            grid.pressToolbarButton("&XXL")
        except Exception:
            grid.currentCellRow = 0
            grid.currentCellColumn = "MATNR"
            grid.contextMenu()
            grid.selectContextMenuItem("&XXL")

        time.sleep(1.5)

        session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
        session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "31"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        time.sleep(2)
        
        # --- SANITIZED: Local File Paths ---
        korensk_mapa = r"C:\Users\YOUR_USERNAME\Path\To\Project"
        ciljna_mapa = os.path.join(korensk_mapa, "Export")
        ciljno_ime = "Zaloga_SI10.xlsx"

        os.makedirs(ciljna_mapa, exist_ok=True)

        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ciljna_mapa
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ciljno_ime

        try:
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
        except Exception:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

        print("3. Excel export successful...")

    except Exception as e:
        print(f"REASON: SAP GUI step failed -> {e}", file=sys.stderr)
        sys.exit(1)

    excel_pot = os.path.join(ciljna_mapa, ciljno_ime)
    if not os.path.exists(excel_pot):
        print("REASON: Excel file missing", file=sys.stderr)
        sys.exit(1)

    try:
        df = pd.read_excel(excel_pot)

        # Pretvorba stolpca Unrestricted v cela števila (brez .0)
        if "Unrestricted" in df.columns:
            # Numerična pretvorba, prazne vrednosti (NaN) postanejo 0 ali pa se ohranijo prazne
            df["Unrestricted"] = pd.to_numeric(df["Unrestricted"], errors="coerce").fillna(0).astype(int)

        df = df.fillna("").astype(str)
        podatki_za_prenos = df.values.tolist()
    except Exception as e:
        print(f"REASON: Excel read error -> {e}", file=sys.stderr)
        sys.exit(1)

    # 3. Google Sheets posodobitev
    print("4. Connecting to Google Sheets...")
    pot_do_json = os.path.join(korensk_mapa, "Credentials", "credentials.json")
    if not os.path.exists(pot_do_json):
        print("REASON: Google credentials missing", file=sys.stderr)
        sys.exit(1)

    try:
        gc = gspread.service_account(filename=pot_do_json)
        
        # --- SANITIZED: Google Sheet URL ---
        url_tabele = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit"
        sh = gc.open_by_url(url_tabele)
        worksheet = sh.get_worksheet(0)

        worksheet.batch_clear(["A2:Z10000"])
        if podatki_za_prenos:
            worksheet.update("A2", podatki_za_prenos)

    except Exception:
        print("REASON: Google Sheets error", file=sys.stderr)
        sys.exit(1)

    trajanje_sekund = time.time() - zacetni_cas
    print(f"[OK] Zaloga osvežena v Google Sheets [Timestamp: {time.strftime('%d-%m-%Y %H:%M:%S')}] (Trajanje: {trajanje_sekund:.2f} s)")

except Exception:
    print("REASON: Unknown system error", file=sys.stderr)
    sys.exit(1)