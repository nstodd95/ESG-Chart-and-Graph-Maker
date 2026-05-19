# LCA Chart Tool — SimaPro LCIA Visualiser for Excel

Parses SimaPro LCIA CSV exports and generates formatted grouped bar charts
embedded directly into an Excel workbook. A button inside Excel triggers the
Python script via xlwings. No coding required to operate.

---

## Contents

1. [Prerequisites](#prerequisites)
2. [Install Python dependencies](#install-python-dependencies)
3. [Build the Excel workbook](#build-the-excel-workbook)
   - [Step 1 — Create sheets](#step-1--create-sheets)
   - [Step 2 — Settings sheet layout](#step-2--settings-sheet-layout)
   - [Step 3 — Create named ranges](#step-3--create-named-ranges)
   - [Step 4 — Add data validation dropdowns](#step-4--add-data-validation-dropdowns)
   - [Step 5 — Pre-fill default values](#step-5--pre-fill-default-values)
   - [Step 6 — Set up the Config sheet](#step-6--set-up-the-config-sheet)
   - [Step 7 — Install the xlwings add-in](#step-7--install-the-xlwings-add-in)
   - [Step 8 — Add VBA macros](#step-8--add-vba-macros)
   - [Step 9 — Add the buttons](#step-9--add-the-buttons)
   - [Step 10 — Save and close](#step-10--save-and-close)
4. [First run](#first-run)
5. [Windows vs Mac differences](#windows-vs-mac-differences)
6. [Adding more color slots](#adding-more-color-slots)
7. [SimaPro CSV format — assumptions and known limitations](#simapro-csv-format--assumptions-and-known-limitations)
8. [Troubleshooting](#troubleshooting)

---

## Prerequisites

| Requirement | Minimum version | Where to get it |
|---|---|---|
| Python | 3.9 | https://www.python.org/downloads/ |
| Microsoft Excel | 2016 or Microsoft 365 | Your organisation |
| pip | bundled with Python | — |

**Mac users:** Do not use the system Python (`/usr/bin/python3`). Install Python
from python.org or via Homebrew: `brew install python`.

---

## Install Python dependencies

1. Open **Terminal** (Mac) or **Command Prompt** (Windows).

2. Navigate to the project folder:
   ```
   cd /path/to/lca_chart_tool
   ```

3. Install all dependencies:
   ```
   pip install -r requirements.txt
   ```
   This installs pandas, matplotlib, xlwings, and numpy.

4. Verify the install:
   ```
   python run.py
   ```
   You should see: `run.py loaded. Functions load_csv() and generate_chart() are ready.`

---

## Build the Excel workbook

You need to build this once. The workbook must be saved in the **same folder as
`run.py`** so xlwings can find the Python script.

### Step 1 — Create sheets

1. Open Excel. Create a new blank workbook.
2. Save it immediately as **`LCA_Chart_Tool.xlsm`** (macro-enabled workbook)
   inside the `lca_chart_tool/` folder.
   - File > Save As > change "Save as type" to **Excel Macro-Enabled Workbook (\*.xlsm)**
3. Rename the existing sheet tabs to create four sheets in this order:
   - **Settings** — user controls and buttons
   - **Data** — parsed CSV data (written by the script)
   - **Charts** — embedded chart image (written by the script)
   - **Config** — helper data for dropdowns (will be hidden)

   To rename a tab: right-click the tab name > Rename.
   To add a new tab: click the **+** button at the bottom.

---

### Step 2 — Settings sheet layout

Click the **Settings** tab. Enter the following text labels and leave the
adjacent B-column cells empty for now (you'll name them in Step 3).

| Cell | Content | Notes |
|------|---------|-------|
| A1 | `LCA Chart Tool — Settings` | Make this bold, size 14 |
| A3 | `CSV File Path` | Bold |
| C3 | `← Paste the full file path here` | Light grey italic hint text |
| A4 | `Chart Title` | Bold |
| A6 | `VALUE MODE` | Bold, uppercase section header |
| A7 | `Value Mode` | |
| A8 | `Reference Scenario` | |
| A10 | `DISPLAY OPTIONS` | Bold, uppercase section header |
| A11 | `Show Data Labels` | |
| A12 | `X-Axis Label Rotation (degrees)` | |
| A13 | `Output DPI` | |
| A15 | `BAR COLORS — one hex code per scenario` | Bold, uppercase section header |
| A16 | `Color 1` | |
| A17 | `Color 2` | |
| A18 | `Color 3` | |
| A19 | `Color 4` | |
| A20 | `Color 5` | |
| A21 | `Color 6` | |
| A23 | `Status` | Bold |
| A25 | *(leave blank — buttons go here)* | |

Widen column A to about 35 characters and column B to about 40 characters so
content is readable. Columns C onward can stay narrow.

---

### Step 3 — Create named ranges

Named ranges are how the Python script reads and writes values without
referring to specific cell addresses.

Go to **Formulas** tab > **Name Manager** > **New** for each row below.

| Name | Refers to (exact text to enter) | Sheet |
|------|----------------------------------|-------|
| `CSV_FILE_PATH` | `=Settings!$B$3` | Settings |
| `CHART_TITLE` | `=Settings!$B$4` | Settings |
| `VALUE_MODE` | `=Settings!$B$7` | Settings |
| `REFERENCE_SCENARIO` | `=Settings!$B$8` | Settings |
| `SHOW_DATA_LABELS` | `=Settings!$B$11` | Settings |
| `X_LABEL_ROTATION` | `=Settings!$B$12` | Settings |
| `OUTPUT_DPI` | `=Settings!$B$13` | Settings |
| `Color_1` | `=Settings!$B$16` | Settings |
| `Color_2` | `=Settings!$B$17` | Settings |
| `Color_3` | `=Settings!$B$18` | Settings |
| `Color_4` | `=Settings!$B$19` | Settings |
| `Color_5` | `=Settings!$B$20` | Settings |
| `Color_6` | `=Settings!$B$21` | Settings |
| `STATUS_MESSAGE` | `=Settings!$B$23` | Settings |
| `SCENARIO_LIST` | `=Config!$A$1:$A$20` | Config |

> **Important:** The names must be entered exactly as shown above, including
> capitalisation and underscores. The Python script references them by name.

---

### Step 4 — Add data validation dropdowns

Data validation creates the dropdown menus in column B. For each cell below:
select the cell > **Data** tab > **Data Validation** > **Settings** tab.

| Cell | Allow | Source |
|------|-------|--------|
| B7 (`VALUE_MODE`) | List | `Raw values,Normalized (% of max)` |
| B8 (`REFERENCE_SCENARIO`) | List | `=Config!$A$1:$A$20` |
| B11 (`SHOW_DATA_LABELS`) | List | `Yes,No` |
| B12 (`X_LABEL_ROTATION`) | List | `0,30,45,90` |
| B13 (`OUTPUT_DPI`) | List | `150,300` |

For cells B3 (CSV file path) and B4 (chart title): leave as plain text cells,
no validation needed.

For cells B16–B21 (colors): leave as plain text. Users type hex codes directly.

---

### Step 5 — Pre-fill default values

Enter these starting values in the B-column cells so the form works on first use:

| Cell | Default value |
|------|---------------|
| B4 | `LCIA Results` |
| B7 | `Normalized (% of max)` |
| B11 | `No` |
| B12 | `45` |
| B13 | `150` |
| B16 | `#4472C4` |
| B17 | `#ED7D31` |
| B18 | `#A9D18E` |
| B19 | `#FF0000` |
| B20 | `#7030A0` |
| B21 | `#00B0F0` |

Leave B3 (CSV file path), B8 (reference scenario), and B23 (status) blank.

---

### Step 6 — Set up the Config sheet

The Config sheet holds the scenario names that populate the Reference Scenario
dropdown. It will be hidden so users don't accidentally edit it.

1. Click the **Config** tab.
2. Click cell **A1**.
3. You do not need to type anything — the Python script will fill A1:A20
   automatically when "Load CSV" is clicked.
4. Right-click the **Config** tab > **Hide**.

---

### Step 7 — Install the xlwings add-in

The xlwings add-in provides the `RunPython` command that connects the Excel
buttons to your Python script.

**On Mac:**
```
xlwings addin install
```
Open Terminal, run this command, then restart Excel.

**On Windows:**
```
xlwings addin install
```
Open Command Prompt as Administrator, run this command, then restart Excel.

**Verify the install:** After restarting Excel, you should see an **xlwings**
tab in the Excel ribbon. If you do not, see [Troubleshooting](#troubleshooting).

**Alternative — standalone mode (no add-in):**
If you cannot install the add-in (e.g. no admin rights), you can embed the
xlwings VBA module directly in the workbook:

1. In Terminal or Command Prompt, run:
   ```
   python -c "import xlwings, os; print(os.path.join(os.path.dirname(xlwings.__file__), 'xlwings.bas'))"
   ```
   This prints the path to `xlwings.bas`.

2. Open the VBA Editor in Excel: **Alt+F11** (Windows) or **Fn+Option+F11** (Mac).

3. In the VBA Editor: **File** > **Import File** > select the `xlwings.bas`
   file printed in step 1.

4. Close the VBA Editor. `RunPython` is now available in this workbook without
   needing the add-in.

---

### Step 8 — Add VBA macros

1. Open the VBA Editor: **Alt+F11** (Windows) or **Fn+Option+F11** (Mac).

2. In the left panel, find your workbook: **VBAProject (LCA_Chart_Tool.xlsm)**.

3. Double-click **Module1** if it exists, or right-click the workbook name >
   **Insert** > **Module** to create one.

4. Paste the following code into the module:

```vb
Sub LoadCSV()
    RunPython ("import run; run.load_csv()")
End Sub

Sub GenerateChart()
    RunPython ("import run; run.generate_chart()")
End Sub
```

5. Close the VBA Editor. Save the workbook (Ctrl+S / Cmd+S).

---

### Step 9 — Add the buttons

Buttons require the **Developer** tab to be visible.

**Enable the Developer tab** (one-time setup):
- **Windows:** File > Options > Customize Ribbon > check **Developer** > OK
- **Mac:** Excel > Preferences > Ribbon & Toolbar > check **Developer** > Save

**Add the "Load CSV" button:**

1. Click the **Settings** tab.
2. Go to **Developer** tab > **Insert** > under **Form Controls**, click the
   **Button (Form Control)** icon (rectangle with thin border).
3. Draw a button in the area around **row 25, column B**.
4. When the "Assign Macro" dialog appears, select **LoadCSV** > click **OK**.
5. Right-click the button > **Edit Text** > type `Load CSV`.

**Add the "Generate Chart" button:**

1. Repeat steps 2–3, drawing a second button to the right of the first
   (around row 25, column D).
2. Assign macro **GenerateChart**.
3. Label it `Generate Chart`.

**Tip:** Right-click > Format Control > Colors and Lines to style the buttons.

---

### Step 10 — Save and close

1. Save the workbook: **Ctrl+S** (Windows) / **Cmd+S** (Mac).
2. Confirm it is saving as **.xlsm** (macro-enabled). If Excel prompts you to
   "Keep in current format" vs "Use Macro-Free Format", choose **Keep Current**.
3. Close and re-open the workbook to confirm macros load without errors.

When you re-open the file, Excel will ask whether to enable macros.
Click **Enable Content** / **Enable Macros**.

---

## First run

1. Export your LCIA results from SimaPro as a CSV file.
   - In SimaPro: go to your calculation results > Export > CSV
   - Note the full file path where it is saved

2. Open `LCA_Chart_Tool.xlsm`.

3. On the **Settings** sheet:
   - Paste the full path to your SimaPro CSV into **B3** (CSV File Path)
     - **Mac example:** `/Users/yourname/Desktop/my_export.csv`
     - **Windows example:** `C:\Users\yourname\Desktop\my_export.csv`
   - Confirm the Chart Title in B4 is what you want

4. Click **Load CSV**.
   - Watch the **Status** cell (B23). It will show a loading message, then:
     - **Success:** `CSV loaded successfully: 12 impact categories, 2 scenario(s) — Scenario A, Scenario B`
     - **Failure:** an error message describing what went wrong

5. The **Data** sheet now contains the parsed table. Review it to confirm the
   data looks correct before generating the chart.

6. The **Reference Scenario** dropdown (B8) is now populated. It is only used
   when Value Mode = "Normalized (% of max)" — you can leave it as-is for now.

7. Adjust any settings (colors, rotation, data labels, DPI).

8. Click **Generate Chart**.
   - The Status cell will update. On success, switch to the **Charts** sheet
     to see the embedded bar chart.

---

## Windows vs Mac differences

| Topic | Windows | Mac |
|-------|---------|-----|
| xlwings backend | COM automation (built into Windows) | xlwings server or appscript |
| Add-in install | `xlwings addin install` in Command Prompt | `xlwings addin install` in Terminal |
| File paths in B3 | Use backslash: `C:\Users\name\file.csv` | Use forward slash: `/Users/name/file.csv` |
| VBA Editor shortcut | Alt+F11 | Fn+Option+F11 |
| Enable macros on open | "Enable Content" banner | Security dialog, click "Enable Macros" |
| Macro execution speed | Generally fast | May be slightly slower; xlwings shows a progress icon in the dock |
| xlwings server (Mac, newer versions) | Not needed | If `RunPython` hangs, run `xlwings runpython-server` in Terminal first |

**Mac additional step — first run only:**
If buttons do nothing or you see an accessibility permission error, go to:
System Preferences > Security & Privacy > Privacy > Accessibility
and ensure Python (or Excel) is listed and checked.

---

## Adding more color slots

The tool supports up to 6 color slots out of the box. If your CSV has more
than 6 scenarios, the colors cycle (scenario 7 reuses Color 1, etc.).

To add a permanent 7th (or more) slot:

1. **In Excel:** Add a new row below the Color 6 row on the Settings sheet
   (e.g. row 22). Label A22 `Color 7`. Leave B22 empty.

2. **Create the named range:** Formulas > Name Manager > New:
   - Name: `Color_7`
   - Refers to: `=Settings!$B$22`

3. **In `lca_tool/excel_writer.py`:** Find the line:
   ```python
   NR_BAR_COLORS: List[str] = [f"Color_{i}" for i in range(1, 7)]
   ```
   Change `7` to `8` (or however many slots you want).

4. Save the Python file. No other changes are needed.

---

## SimaPro CSV format — assumptions and known limitations

The parser makes the following assumptions. If your export deviates from these,
the relevant adjustment is noted in the code at the `# ASSUMPTION` comment.

| Assumption | Where to adjust if wrong |
|-----------|--------------------------|
| The first column header of the data table is **"Impact category"** (case-insensitive) | Change `_IMPACT_CATEGORY_HEADER` in `parser.py` |
| The second column is **"Unit"** | The parser uses position [1]; if SimaPro adds extra columns before Unit, shift the index in `parse_simapro_csv()` |
| All columns after "Unit" are scenario/product system result columns | The parser treats columns [2:] as scenarios; if SimaPro adds extra metadata columns, exclude them by name |
| The data block ends at the **first blank line** after the header, or at a row starting with **"End"** | Adjust the stop condition in `parse_simapro_csv()` |
| **Delimiter** is comma or semicolon, auto-detected by frequency | If the file uses another delimiter (tab, pipe), change `_detect_delimiter()` in `parser.py` |
| **Encoding** is UTF-8 or Latin-1 | The parser tries utf-8-sig, utf-8, latin-1 in order; add another encoding to the list if needed |
| Scientific notation values (e.g. `1.23E+04`) are handled by Python's `float()` natively | No adjustment needed |
| European-locale files (semicolon delimiter, comma decimal) have commas in numeric cells replaced with periods | This replacement only occurs when the delimiter is `;` |
| SimaPro may insert **subtotal or category group header rows** between data rows | These are skipped automatically because their scenario cells are non-numeric |

**The single-chart limitation:** Because LCIA results span many orders of magnitude,
raw-mode charts are visually misleading — a global warming bar at 10⁴ kg CO₂-eq
will dwarf a freshwater eutrophication bar at 10⁻³ kg P-eq, making most bars
invisible. The chart includes a red warning box when raw mode is selected.
**Normalized mode is strongly recommended** for multi-category comparison.

---

## Troubleshooting

**Status cell shows nothing / button does nothing**
- Confirm the xlwings add-in is installed: look for the xlwings tab in the
  Excel ribbon. If absent, re-run `xlwings addin install` and restart Excel.
- Confirm macros are enabled: close and re-open the workbook, click "Enable Content".

**`Named range 'X' was not found`**
- Open Formulas > Name Manager and verify every name in [Step 3](#step-3--create-named-ranges)
  is listed with the correct spelling and cell reference.

**`File not found` error**
- On Mac: make sure to use forward slashes and the full path starting with `/Users/`.
- On Windows: use the full path. Copy it from File Explorer's address bar.
- Avoid wrapping the path in quotes inside the cell.

**`Could not locate the LCIA data table`**
- Open the CSV in a text editor (not Excel) and look for the row containing
  "Impact category" (or your region's equivalent). If it uses a different label,
  update `_IMPACT_CATEGORY_HEADER` in `parser.py`.
- Check that the file was exported as LCIA / Characterisation results, not
  inventory results (which have a different column structure).

**`RunPython` is not defined (VBA error)**
- The xlwings VBA module is not in this workbook. Either reinstall the add-in
  or import `xlwings.bas` manually (see [Step 7](#step-7--install-the-xlwings-add-in)).

**Chart not appearing on the Charts sheet**
- Check the Status cell for an error message.
- Make sure "Load CSV" was run successfully before "Generate Chart".
- If the chart sheet is blank, try scrolling — the picture may have been placed
  off-screen. Delete the sheet contents and click "Generate Chart" again.

**Mac: button hangs or Excel freezes**
- Open Terminal and run: `xlwings runpython-server`
  Keep Terminal open and try the button again.
- Grant accessibility permissions: System Preferences > Security & Privacy >
  Privacy > Accessibility > add Python or Excel.

**`ModuleNotFoundError: No module named 'lca_tool'`**
- The Excel workbook is not saved in the same folder as `run.py`.
  Move `LCA_Chart_Tool.xlsm` into the `lca_chart_tool/` folder and try again.
- Or the virtual environment with the installed dependencies is not the one
  xlwings is using. Check with: `python -c "import xlwings; print(xlwings.__file__)"`.
