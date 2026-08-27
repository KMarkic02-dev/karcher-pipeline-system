# SAP to Google Sheets Automated System

An end-to-end automation and data visualization solution that extracts real-time business data (Inventory, Customers, Delivery Status, and Price Lists) from **SAP GUI via Python**, syncs it to a **Google Sheets backend with the help of Google API**, and presents it through an interactive, responsive **Web Dashboard built with Google Apps Script & Tailwind CSS**.

Developed independently to streamline business operations, improve data transparency, and reduce manual reporting efforts.


## Key Features

### Backend Automation (Python & SAP GUI)
* **Automated SAP Extraction:** Uses `win32com.client` to interface directly with SAP GUI scripting engines (`MM60`, `MB52`, `ZDE1SD131`, `ZWWXSD138`).
* **Web Scraping Integration:** Automated XML data retrieval from Serv-U WebClient using `Selenium`.
* **Data Processing & Normalization:** Employs `Pandas` to clean, format, merge, and structure raw SAP/XML exports (e.g., handling missing identifiers, zero-padding material codes, and deduplication).
* **Cloud Synchronization:** Uses `gspread` and Google Service Accounts to push clean datasets to Google Sheets in real-time.
* **Orchestration & Error Handling:** Includes a **Master Runner script** that sequentially executes data pipelines, logs errors, and updates the system status flag (`ERROR: DD.MM.YYYY HH:MM`).

### Frontend Interface PC & Mobile (Google Apps Script & Web Tech)
* **Custom Web & Mobile Dashboard:** Built with HTML, JavaScript, and styled with **Tailwind CSS**.
* **Multi-Module Navigation:** 
  * **Home:** Quick launcher with consistent, clean card layout.
  * **Customers:** Searchable business database with detailed order breakdowns.
  * **Stock:** Live stock levels categorized by location with integrated price lookup (VPC, MPC, MPC+DDV).
  * **Price List:** Interactive article search by Name, Code, or EAN, optimized with a 30-item render cap for instant response times.
* **Order & Delivery Tracker:** Real-time tracking of orders grouped by Purchase Order (`PO`), filtered by delivery status (**Pending**, **In Processing**, **Shipped**).
* **Responsive Layouts:** Supports dynamic view switching between **Grid Cards** and an **Excel-like Table View**, fully optimized for desktop and mobile devices.
## 🏗 System Architecture
┌─────────────────┐       ┌─────────────────┐
│   SAP System    │       │  Serv-U Portal  │
└────────┬────────┘       └────────┬────────┘
│ (SAP Scripting)         │ (Selenium WebScraper)
▼                         ▼
│                    [.xml Exports]
│                         │
└───────────┬─────────────┘
▼
┌───────────────────────────┐
│   Python Master Pipeline  │
│  (Pandas Data Processing) │
└─────────────┬─────────────┘
│ (gspread API)
▼
┌───────────────────────────┐
│   Google Sheets Backend   │
└─────────────┬─────────────┘
│ (Google Apps Script API)
▼
┌───────────────────────────┐
│  Interactive Web Board    │
│   (HTML5 + Tailwind CSS)  │
└───────────────────────────┘


## 🛠 Tech Stack

* **Automation & Scripting:** Python 3.x, `win32com.client` (SAP GUI Scripting), `Selenium`
* **Data Manipulation:** `pandas`, `openpyxl`, `xml.etree.ElementTree`
* **Cloud & API Integration:** Google Sheets API, `gspread`, Google Apps Script (GAS)
* **Frontend Web App:** HTML5, Modern JavaScript (ES6+), Tailwind CSS, Responsive Design


## 📁 Repository Structure
│ 
├── BackEndConnection/
│   └── Code.gs                  # Google Apps Script Server Engine & API Handlers
│ 
├── FrontEndPage/
│   └── index.html               # Responsive Frontend Web Interface
│ 
├── Python_SAP_Extraction_Scripts/
│   ├── Izvoz_Artiklov.py        # SAP MM60 Material Exporter
│   ├── Izvoz_Delivery_Daily.py  # SAP Orders Exporter (Today - 60 days)
│   ├── Izvoz_Delivery_Monthly.py# SAP Orders Exporter (Today - Start of Year)
│   ├── Izvoz_Kupcev.py          # SAP ZWWXSD138 Customer Exporter
│   ├── Izvoz_XML_Data.py        # Serv-U WebClient XML Exporter
│   ├── Izvoz_Zaloge.py          # SAP MB52 Local Stock Exporter
│   │ 
│    └── Master_Script.py         # Main Orchestration & Google Sheets Sync
│ 
└── README.md                    # Project Documentation 

## 🔒 Security & Privacy

All sensitive company data, API credentials, Google Sheet IDs, internal server paths, and employee identifiers have been completely scrubbed and replaced with generic environment placeholders (`YOUR_SPREADSHEET_ID`, `YOUR_USERNAME`, etc.) to protect sensitive data.


## 🚀 Independent Project Statement

This project was conceived, architected, and fully implemented by me as an independent automation solution to replace manual reporting workflows with an automated, cloud-based data dashboard.