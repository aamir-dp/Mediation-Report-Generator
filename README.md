# CSV/TSV Import & Mediation-Analysis Add-On

This Google Apps Script lets you import a CSV or TSV file of your mediation data, computes an **Earning Share** column, and builds two (or three) pivot-style analysis sheets with sparklines and variant flags.

---

## Features

- **One-click import** of CSV/TSV via a custom menu  
- Automatic creation of a `Data_<rangeName>` sheet with your raw data and an extra **Earning Share** column  
- Two analysis sheets per import:
  1. `Analysis_<rangeName>_EarningShare` — pivot of Earning Share by “Ad source instance” vs. date  
  2. `Analysis_<rangeName>_MatchRate` — pivot of Match Rate (column 7) by “Ad source instance” vs. date  
  3. (Optional) `Analysis_<rangeName>_eCPM` — pivot of eCPM (column 5) by “Ad source instance” vs. date  
- Adds **Variant** column (A vs B) and a **Chart** column with a sparkline trend over all date-columns  
- Applies a filter on “Ad source instance” so you can type your mediation-group name (e.g. `T1 Vidma Rect Banner OPMC`) to focus your view  
- Automatically installed via `onOpen()`

---

## Installation

1. Open your target Google Sheet.  
2. Go to **Extensions → Apps Script**.  
3. Replace the default `Code.gs` with the contents of **Code.js**, and add a new HTML file named **ImportDialog.html** with the provided markup.  
4. Save and **Deploy → Test deployments** (no special permissions needed beyond “Spreadsheet”).  
5. Reload your sheet.

---

## Usage

1. After reload, you’ll see a new **Custom** menu.  
2. Choose **Custom → Import CSV/TSV + Setup Analysis**.  
3. In the dialog:
   - **Select** your `.csv` or `.tsv` file  
   - **Enter** a valid named-range identifier (e.g. `data2025`)  
   - **Enter** your mediation-group filter text (e.g. `T1 Vidma Rect Banner OPMC`)  
4. Click **Import & Build**.  
5. Three new sheets will appear:
   - **Data_<rangeName>**  
     - Contains your imported rows plus an **Earning Share** column  
   - **Analysis_<rangeName>_EarningShare**  
     - Pivot of Earning Share by “Ad source instance” vs. date  
     - Adds **Variant** and **Chart** columns with sparkline trends  
     - Filter applied on “Ad source instance” with your filter text  
   - **Analysis_<rangeName>_MatchRate**  
     - Same as above but pivots Match Rate  
   - **Analysis_<rangeName>_eCPM**  
     - Same as above but pivots eCPM  

---

## Code Overview

### `onOpen()`
Adds a “Custom” menu to trigger the import dialog.

### `showImportDialog()`
Displays an HTML form for file upload, named-range, and filter text.

### `importCsv(fileText, rangeName, filterText)`
1. **Parses** CSV or TSV text into a 2D array  
2. **Writes** it to `Data_<rangeName>` sheet  
3. **Appends** an **Earning Share** column with:
   ```js
   =IFERROR(
     D2 /
     SUMIFS($D$2:$D$<last>, $A$2:$A$<last>, A2, $C$2:$C$<last>, "*␟"&RIGHT(C2,1)),
     ""
   )
4. **Defines** the named range over the full table

5. **buildAnalysis(suffix, valueIndex)**
   - Creates `Analysis_<rangeName>_<suffix>`
   - Inserts a `QUERY({ … }, "select … pivot …")` pivot formula
   - Inserts **Variant** (`=RIGHT(A2)`) and **Chart** sparkline columns
   - Applies a filter on column A matching your `filterText`

---

### Customization

- **Change the pivot field** by adjusting the `valueIndex` when calling `buildAnalysis`.
- **Add/remove analysis sheets** simply by adding/removing calls to `buildAnalysis`.

---

### Troubleshooting

- **Sparkline only shows two points** if your script flushes and measures `lastCol` before all formulas recalculate; the provided code calculates offsets dynamically so it spans **all** date-columns.
- If your named range or filter text is invalid, you’ll see a dialog alert.
