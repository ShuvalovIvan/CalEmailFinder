# Data Mapper & Web Scraper Application

## 1. Introduction
This application is a powerful, desktop-based tool designed to help you extract public contact information (Principals, Emails, Phone Numbers, and Job Titles) from websites and map them directly into your Excel or CSV files.

It replaces manual searching with an automated, threaded process that includes:
* **Multi-Field Extraction:** Automatically finds First Name, Last Name, Job Title, Email, and Phone.
* **Crash Recovery:** Auto-saves your progress every 10 rows so you never lose data.
* **Merge Back:** A specialized tool to fix failed rows in a separate file and merge them back into your master list seamlessly.
* **Headless Browser:** Runs silently in the background without opening visible windows.

---

## 2. Installation & Setup

**No Python installation is required** if you are using the pre-built executables.

### **Windows Users**
1.  Download the **`DataMapper.exe`** file frome Release.
2.  Double-click **`DataMapper.exe`**.
3.  *Note:* On the very first launch, a command window may appear briefly. This is the app ensuring the necessary browser engine is installed.

### **Mac Users**
1.  Download the **`DataMapper`** file frome Release.
2.  Double-click the **`DataMapper`** file.
3.  *Troubleshooting:* If you see an "Unidentified Developer" warning:
    * Right-click (or Control+Click) the file.
    * Select **Open**.
    * Click **Open** in the confirmation dialog.

---

## 3. How to Use

### **Step 1: Load Your Data**
* Click **"Open File"** and select your `.csv` or `.xlsx` file.
* The application will load your data into the grid view.

### **Step 2: Manage Columns (Optional)**
* Use the sidebar tools to clean up your file before processing.
* **Move Up/Down:** Reorder columns.
* **Delete:** Remove unnecessary columns.
* **Merge:** Select multiple columns (Ctrl+Click) to combine them into one (e.g., combining "School" and "City" to create a better search term).

### **Step 3: Run Extraction**
1.  Select the **column containing the search terms** (e.g., "School Name").
2.  Click **"Extract Principal Info"**.
3.  A mapping window will appear. Choose where to save the data:
    * **First Name** $\rightarrow$ *Create New Column*
    * **Job Title** $\rightarrow$ *Create New Column*
    * **Email** $\rightarrow$ *Overwrite Existing Column* (or Create New)
4.  Click **Start Extraction**.

### **Step 4: Save**
* Click **"Save / Export"** to save your work. The app automatically formats Excel files with text wrapping for better readability.

---

## 4. Advanced Features

### **Fixing Failed Rows (The "Merge Back" Workflow)**
Sometimes websites time out or data is missing. Instead of re-running the whole file, use this workflow:

1.  **Export Failures:** Click **"Export Failed Rows"**. This creates a small CSV containing only the rows that failed (e.g., "Error" or empty cells).
2.  **Fix & Retry:** Open that small CSV, fix the search terms manually, or re-run the scraper on just those rows.
3.  **Merge:**
    * Load your **MASTER** (original) file.
    * Click **"Merge Fixed Data"**.
    * Select the small **FIXED** file.
    * Select the **Unique ID** column (e.g., "School Name") to match rows.
    * The app will automatically update *only* the rows that have new, valid data.

### **Network Recovery**
If the scraper hits a timeout or network error:
1.  The app will **Pause** and show a popup.
2.  The popup will show the **Last Visited URL** so you can check if the link is broken.
3.  You can **Edit the Search Term** directly in the popup and click **Retry** to try again immediately.

---

## 5. For Developers (Running from Source)

If you want to modify the code or run it without the executable:

**Requirements:**
* Python 3.10+
* Dependencies: `pandas`, `openpyxl`, `xlsxwriter`, `tkinterdnd2`, `playwright`

**Setup:**
```bash
pip install pandas openpyxl xlsxwriter tkinterdnd2 playwright
playwright install
```

**Run:**
```bash
python menu.py
```

**Build Executables:**
To create the .exe or Mac app yourself, use PyInstaller:
# Windows
```bash
pyinstaller --noconsole --onefile --name "DataMapper" --collect-all tkinterdnd2 menu.py
```
# Mac (Universal Binary)
```bash
pyinstaller --noconsole --onefile --name "DataMapper" --collect-all tkinterdnd2 --target-architecture universal2 menu.py
```

### Folder Structure for Distribution

When you zip this up for others, your folder should look like this:

```text
DataMapper_v1.0/
│
├── README.md               <-- The file above
│
├── Windows/
│   └── DataMapper.exe      <-- The Windows Executable
│
└── Mac/
    └── DataMapper          <-- The Mac Executable
```
