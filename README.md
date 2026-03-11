# Teacher Calendar Export

## Goal

Turn the building school Excel timeplan into **ICS calendar files** for individual teachers. The script reads the "06 Timeplan" sheet, finds all classes for a given teacher (by name/code), infers dates, times, subjects, groups, and rooms, then writes a standard `.ics` file you can import into Google Calendar, Outlook, Apple Calendar, etc.

## How to Use `export_teacher_calendar`

**Function signature:**

```python
export_teacher_calendar(excel_path, teacher_name, output_path=None) -> str
```

**Parameters:**

| Argument       | Description |
|----------------|-------------|
| `excel_path`   | Path to the Excel timeplan file (e.g. `"01 Timeplan BYGG 2025-2026.xlsx"`). |
| `teacher_name` | Teacher code/name as it appears in the sheet (e.g. `"RS7"`, `"RS4"`). |
| `output_path`  | Optional. Where to save the `.ics` file. If omitted, the file is saved next to the Excel file as `{teacher_name}_calendar.ics` (e.g. `rs7_calendar.ics`). |

**Returns:** The path to the created ICS file.

**Example:**

```python
from excel2ics import export_teacher_calendar

# Export RS7's calendar (output: rs7_calendar.ics in same folder as Excel)
path = export_teacher_calendar(
    "/path/to/01 Timeplan BYGG 2025-2026.xlsx",
    "RS7"
)

# Export RS4 to a specific file
path = export_teacher_calendar(
    "/path/to/01 Timeplan BYGG 2025-2026.xlsx",
    "RS4",
    output_path="/path/to/rs4_schedule.ics"
)
```

**From the command line:** Run the script; it exports calendars for RS7 and RS4 using the path set in `excel_file` at the bottom of the script.

**Requirements:** Python 3.x and `openpyxl` (`pip install openpyxl`).

## Web app (upload Excel → download ICS)

A simple web UI lets you upload the Excel timeplan, enter one or more teacher codes, and download the resulting `.ics` file(s).

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the app:
   ```bash
   python app.py
   ```
3. Open [http://localhost:5050](http://localhost:5050), upload your Excel file, enter teacher code(s) (e.g. `RS7, RS4`), then click **Generate & download .ics**. One teacher → one `.ics` file; multiple teachers → a ZIP of `.ics` files.

### Deploy on Render (free tier)

Use a **Web Service**, not a Static Site (the app runs Python/Flask).

1. In [Render](https://render.com): **Dashboard → New → Web Service**.
2. Connect your GitHub repo (e.g. `excel2ics_translation`).
3. Settings:
   - **Name:** e.g. `excel2ics`
   - **Branch:** `main`
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn --bind 0.0.0.0:$PORT app:app`
4. Click **Create Web Service**. Render will build and deploy; your app will be at `https://<name>.onrender.com`.  
   Free tier may spin down after inactivity (first load after that can be slow).

### Deploy on PythonAnywhere (free tier)

1. **Sign up** at [pythonanywhere.com](https://www.pythonanywhere.com) and open the **Dashboard**.

2. **Upload your code**  
   - **Files** tab → go to your user directory (e.g. `/home/yourusername`).  
   - Upload your project (or clone from Git): you need `app.py`, `excel2ics.py`, `wsgi.py`, `requirements.txt`, and the `static/` folder (with `index.html` inside).  
   - Example layout: `/home/yourusername/excel/` containing `app.py`, `excel2ics.py`, `wsgi.py`, `requirements.txt`, and `static/index.html`.

3. **Create a virtualenv and install dependencies**  
   - **Consoles** → **$ Bash**. Then:
   ```bash
   cd ~/excel
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```

4. **Add a Web app**  
   - **Web** tab → **Add a new web app** → **Next** → choose **Manual configuration** (not Django) → **Next** → pick **Python 3.10** (or the version you used).  
   - Set **Source code** and **Working directory**:  
     - **Directory:** `/home/yourusername/excel`  
   - In **Code** section, set **WSGI configuration file** to your project’s WSGI file, e.g.:  
     `/home/yourusername/excel/wsgi.py`

5. **WSGI file**  
   - In **Web** → **Code** section, set **WSGI configuration file** to your project’s file: `/home/yourusername/excel/wsgi.py`.  
   - The repo’s `wsgi.py` already loads the Flask app, so you don’t need to edit it.  
   - If you prefer to use the default file (e.g. `/var/www/yourusername_pythonanywhere_com_wsgi.py`), open it and replace its contents with:
   ```python
   import sys
   path = '/home/yourusername/excel'
   if path not in sys.path:
       sys.path.insert(0, path)
   from app import app as application
   ```
   (Replace `yourusername` with your PythonAnywhere username.)

6. **Set virtualenv**  
   - **Web** → **Virtualenv** → enter: `/home/yourusername/excel/venv`  
   - Click the green check to use it.

7. **Static files (optional but recommended)**  
   - **Web** → **Static files**:  
     - **URL:** `/static/`  
     - **Directory:** `/home/yourusername/excel/static`  
   This lets the server serve `index.html` and assets faster.

8. **Reload the app**  
   - **Web** → click the green **Reload** button.  
   - Your app will be at `https://yourusername.pythonanywhere.com`.

**Free tier notes:**  
- Your app stays on; no spin-down like some other hosts.  
- Free accounts use `yourusername.pythonanywhere.com` and have limits on outbound traffic and CPU; fine for light use and file uploads.

## How to Import an ICS File into Outlook

### Outlook on the web (outlook.com / Microsoft 365)

1. Go to [outlook.com](https://outlook.com) and sign in.
2. Open **Calendar** (calendar icon in the left rail).
3. Click **Add calendar** → **Create Blank Calender** and a create a new calender.
4.  Use **Upload from file** and choose your `.ics` (created by the python codes export_teacher_calendar) file and add it to the calender you created in Step 3.  


The events will appear in the calendar you selected. You can delete or move the imported calendar later from the calendar list if you only wanted the events in your main calendar.
