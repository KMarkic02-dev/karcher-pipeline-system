# SAP → Google Sheets Automated InfoBoard System

An end-to-end automation and data visualization solution that extracts real-time business data — inventory, customers, delivery status, and price lists — from SAP GUI via Python, syncs it to a Google Sheets backend, and presents it through an interactive, responsive web dashboard built with Google Apps Script and Tailwind CSS.

Developed to streamline warehouse operations, improve data transparency, and eliminate manual reporting.

> **Note:** This was an internal tool deployed behind the company's Google Workspace. No public demo or screenshots are available — the repository contains the full code and architecture documentation.

## What the dashboard did

- **Home** — launcher with cards for each module
- **Customers** — searchable database of business customers; click one to see their orders grouped by Purchase Order
- **Orders** — filter by status (Pending / In Processing / Shipped) or search by PO number
- **Stock** — live stock levels by warehouse location, with price lookup per article
- **Price List** — searchable article catalogue with a 30-item render cap for speed
- **Layout toggle** — switch between card grid and Excel-like table view
  
## Impact

- Replaced ~20 recurring IT tickets per day of manual SAP data pulls with a self-serve dashboard
- Eliminated repetitive export work for warehouse staff
- Surfaced pipeline failures end-to-end: SAP error → email alert → visible warning on the dashboard

## Key Features

### Backend Automation (Python + SAP GUI)

- **Automated SAP Extraction** — Uses `win32com.client` to interface directly with SAP GUI scripting engines (transactions MM60, MB52, ZDE1SD131, ZWWXSD138).
- **Web Scraping Integration** — Automated XML data retrieval from Serv-U WebClient using Selenium.
- **Data Processing & Normalization** — Employs Pandas to clean, format, merge, and structure raw SAP and XML exports.
- **Cloud Synchronization** — Pushes cleaned datasets to Google Sheets in real time via `gspread` and Google Service Accounts.
- **Orchestration & Error Handling** — A Master Runner script sequentially executes pipelines, logs per-script status, and updates system flags. One failing export does not stop the others.

### Frontend Interface (Google Apps Script + Web)

- **Custom Web Dashboard** — Built with HTML, JavaScript, and Tailwind CSS.
- **Multi-Module Navigation** — Home launcher with card layout; Customers searchable business database with order breakdowns; Stock with live levels by location and integrated price lookup; Price List with a 30-item render cap for instant performance.
- **Order & Delivery Tracker** — Real-time tracking of orders grouped by Purchase Order, filtered by delivery status (Pending, In Processing, Shipped).
- **Responsive Layouts** — Dynamic view switching between Grid Cards and Excel-like Table View.

## Tech Stack

- **Automation & Scripting** — Python 3, `win32com.client` (SAP GUI Scripting), Selenium
- **Data Manipulation** — pandas, openpyxl, `xml.etree.ElementTree`
- **Cloud & API Integration** — Google Sheets API, gspread, Google Apps Script
- **Frontend** — HTML5, JavaScript (ES6+), Tailwind CSS

## Repository Structure

```
BackEndConnection/
  Code.gs                  Google Apps Script server engine & API handlers

FrontEndPage/
  index.html               Responsive frontend web interface

Python SAP Extraction Scripts/
  Izvoz_Artiklov.py        SAP MM60 exporter
  Izvoz_Delivery_Daily.py  SAP orders exporter (today − 60 days)
  Izvoz_Delivery_Monthly.py SAP orders exporter (today − start of month)
  Izvoz_Kupcev.py          SAP ZWWXSD138 customer exporter
  Izvoz_XML_Data.py        Serv-U WebClient XML exporter
  Izvoz_Zaloge.py          SAP MB52 local stock exporter
  Master_Script.py         Main orchestration & Google Sheets sync

README.md                  Project documentation
```

## Known Limitations

- SAP GUI COM scripting requires an active GUI session — the pipeline runs on a dedicated Windows workstation.
- `time.sleep()` waits are used between SAP steps because the COM interface has no reliable "wait for element" primitive.
- A yearly full-rebuild script (to catch deletions outside the daily window) was designed and partially implemented but not fully rolled out before the project was cut.

## Security & Privacy

All sensitive company data, API credentials, Google Sheet IDs, internal server paths, and employee identifiers have been scrubbed and replaced with generic environment placeholders (e.g. `YOUR_SPREADSHEET_ID`, `YOUR_USERNAME`).

Credentials are loaded from environment variables — see `.env.example`. Never commit real credentials.

## Independent Project Statement

This project was conceived and architected by me, and implemented independently as an automation solution to replace manual reporting workflows with a cloud-based data dashboard. AI assistance was used for implementation details; the system design, architecture, and debugging are my own.

The project was cut short by company budget, so some planned improvements remain partially wired.
