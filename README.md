# End-to-End Weekly Ticket Reporting Automation (SJC)

This project is an end-to-end automation pipeline that collects ticket data from HPSM, processes it through Excel and SharePoint, and generates a weekly operations report and Power BI dashboard.

## 🧩 Overview

**Technologies:**  
Power Automate Desktop, Power Automate Cloud, Excel (Office Scripts), SharePoint Online, Power BI, Virtual Machine (planned), HPSM Ticketing System.

The system:
- Extracts tickets from HPSM using Power Automate Desktop.
- Cleans and filters data in Excel using scripts.
- Sends the Excel file by email for review and uploads it to a SharePoint document library.
- Uses cloud flows to import the Excel data into a SharePoint List (acts as a database).
- Allows the team to categorize and complete ticket data.
- Runs a scheduled flow every 2 hours to check if the latest batch is fully filled.
- When complete, exports the data into a reporting Excel file with pre-defined pivot tables.
- Updates the report weekly with extra data from Power BI and CR/ECR flows.
- Uses Power BI dashboard to visualize insights for management.

## 🛠 Components

- `power-automate-desktop/` – PAD flow for HPSM extraction and Excel pre-processing.
- `power-automate-cloud/` – Cloud flows for import, validation, and report generation.
- `power-bi/` – Power BI report (`.pbix`) for weekly ticket analytics.
- `report-excel-template/` – The master Excel with pivot tables.
- `docs/` – Architecture diagram, screenshots, and documentation.

## 🚀 My Role

I designed and implemented the entire workflow:
- Built the PAD flow to automate HPSM export and Excel filtering.
- Created SharePoint List structure and cloud flows for ingestion and validation.
- Developed the weekly Thursday flows to pull extra data (Power BI + CR/ECR).
- Designed the final reporting Excel and the Power BI dashboard.
- Preparing to host the PAD flow on a virtual machine for full automation.
