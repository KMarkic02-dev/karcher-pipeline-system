SAP to Google Sheets Automated InfoBoard System
An end to end automation and data visualization solution that extracts real time business data like Inventory, Customers, Delivery Status, and Price Lists from SAP GUI via Python, syncs it to a Google Sheets backend, and presents it through an interactive responsive Web Dashboard built with Google Apps Script and Tailwind CSS.

Developed independently to streamline business operations, improve data transparency, and reduce manual reporting efforts.

Key Features
Backend Automation Python and SAP GUI
Automated SAP Extraction: Uses win32com client to interface directly with SAP GUI scripting engines MM60, MB52, ZDE1SD131, and ZWWXSD138.

Web Scraping Integration: Automated XML data retrieval from Serv U WebClient using Selenium.

Data Processing and Normalization: Employs Pandas to clean, format, merge, and structure raw SAP and XML exports.

Cloud Synchronization: Uses gspread and Google Service Accounts to push clean datasets to Google Sheets in real time.

Orchestration and Error Handling: Includes a Master Runner script that sequentially executes data pipelines, logs errors, and updates system flags.

Frontend Interface Google Apps Script and Web Tech
Custom Web Dashboard: Built with HTML, JavaScript, and styled with Tailwind CSS.

Multi Module Navigation: Home offers a quick launcher with a clean card layout. Customers provides a searchable business database with detailed order breakdowns. Stock offers live stock levels categorized by location with integrated price lookup. Price List allows interactive article search optimized with a 30 item render cap for instant performance.

Order and Delivery Tracker: Real time tracking of orders grouped by Purchase Order, filtered by delivery status like Pending, In Processing, and Shipped.

Responsive Layouts: Supports dynamic view switching between Grid Cards and an Excel like Table View.

Tech Stack
Automation and Scripting: Python 3, win32com client for SAP GUI Scripting, and Selenium

Data Manipulation: pandas, openpyxl, and xml etree ElementTree

Cloud and API Integration: Google Sheets API, gspread, and Google Apps Script

Frontend Web App: HTML5, JavaScript ES6 plus, and Tailwind CSS

Repository Structure
BackEndConnection
Code gs: Google Apps Script Server Engine and API Handlers

FrontEndPage
index html: Responsive Frontend Web Interface

Python SAP Extraction Scripts
Izvoz Artiklov py: SAP MM60 Exporter
Izvoz Delivery Daily py: SAP Orders Exporter for Today minus 60 days
Izvoz Delivery Monthly py: SAP Orders Exporter for Today minus Start of Month
Izvoz Kupcev py: SAP ZWWXSD138 Customer Exporter
Izvoz XML Data py: Serv U WebClient XML Exporter
Izvoz Zaloge py: SAP MB52 Local Stock Exporter
Master Script py: Main Orchestration and Google Sheets Sync

README md: Project Documentation

Security and Privacy
All sensitive company data, API credentials, Google Sheet IDs, internal server paths, and employee identifiers have been completely scrubbed and replaced with generic environment placeholders like YOUR SPREADSHEET ID or YOUR USERNAME to comply with security standards.

Independent Project Statement
This project was conceived, architected, and fully implemented by me as an independent automation solution to replace manual reporting workflows with an automated cloud based data dashboard.
