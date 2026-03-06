
### 1. Required Libraries

Since you are using Python 3.13, you must install the specific libraries that handle API requests, data manipulation, and Excel writing. Run this command in your PowerShell:

```powershell
& C:/Users/adm28207/AppData/Local/Microsoft/WindowsApps/python3.13.exe -m pip install pandas requests openpyxl

```

* 
**pandas**: Used for data structuring and cleaning.


* **requests**: Required to fetch data from the Weather.com API.
* **openpyxl**: The engine that allows pandas to save `.xlsx` files.

---

### 2. Prepare the Script

Ensure your script (`c:/UGA_Weather/weatherAthens.py`) uses the dynamic date logic to pull the last 31 days and saves to a fixed master file .

> **Note:** Ensure the folder `C:\UGA_Weather` exists, or the script will fail when trying to save the master file.

---

### 3. Automation Instructions (Windows Task Scheduler)

This will make the script run every day at a specific time (e.g., 8:00 AM) without you doing anything.

1. **Open Task Scheduler**: Press the `Win` key, type "Task Scheduler," and hit Enter.
2. **Create Basic Task**: Click **Create Basic Task** in the right-hand "Actions" pane.
* **Name**: `UGA_Weather_Update`
* **Trigger**: Select **Daily**.
* **Time**: Set it to a time when your computer is usually on (e.g., `08:00:00 AM`).


3. **Action**: Select **Start a Program**.
4. **Settings**:
* **Program/script**: Paste the path to your Python executable:
`C:\Users\adm28207\AppData\Local\Microsoft\WindowsApps\python3.13.exe`
* **Add arguments**: Paste the path to your script:
`C:\UGA_Weather\weatherAthens.py`
* **Start in**: Paste the folder path:
`C:\UGA_Weather`


5. **Finish**: Click **Finish**.

---

### 4. Verification Checklist

To confirm everything is working:

* 
**Manual Run**: Run the script once manually in PowerShell to ensure `KAHN_Weather_Master.xlsx` is created .


* 
**Check Output**: Open the Excel file and verify that columns like **Temperature (°F)** and **Dew Point (°F)** are rounded and populated .


* **Test Task**: In Task Scheduler, right-click your new task and select **Run**. If the "Last Run Result" says `(0x0)`, the automation is successful.

