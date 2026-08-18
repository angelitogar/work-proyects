# Work Projects
## Using Python (win32com) for Excel Automation

In this repository, you'll find work projects designed to automate time-consuming, repetitive Excel processes. These scripts handle data copying, external link updates, monthly roll-forwards, and general periodic tasks that Python executes efficiently in the background.

---

### 📂 Included Processes

* **Process 1: Link Refresh (`refresh`)**
  * Opens heavy Excel workbooks in the background, updates external links/Power Query data sources, and performs a clean save and close.

* **Process 2: Monthly Billing (`facturacion`)**
  * Updates billing databases and transfers records between reports. It uses a **local temporary files technique** (isolating files from OneDrive/SharePoint during processing) to prevent save crashes and sync conflicts.

* **Process 3: Add New Account (`nuevacuenta`)**
  * Uses an interactive prompt to locate account folders, copy formatting, and automatically generate dynamic Excel link formulas for each month into a central consolidated report.

* **Process 5: Monthly Roll-Forward (`actualizarmes`)**
  * Clones the previous month's project block (columns A to CG), increments period IDs, advances the month/year tags, and prepares the worksheet for new data entry.

* **Process 13-08-26**
  * Extracts unstructured data from 4 separate Excel files and standardizes them into a main file using a safe local temporary directory.
