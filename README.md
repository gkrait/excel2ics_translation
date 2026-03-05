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

## How to Import an ICS File into Outlook

### Outlook on the web (outlook.com / Microsoft 365)

1. Go to [outlook.com](https://outlook.com) and sign in.
2. Open **Calendar** (calendar icon in the left rail).
3. Click **Add calendar** → **Create Blank Calender** and a create a new calender.
4.  Use **Upload from file** and choose your `.ics` (created by the python codes export_teacher_calendar) file and add it to the calender you created in Step 3.  


The events will appear in the calendar you selected. You can delete or move the imported calendar later from the calendar list if you only wanted the events in your main calendar.
