# Pivot Slicer Connect Macros

This repository contains two VBA modules for Excel that automate the creation and connection of PivotTables and Slicers.

## Macros

### `CreatePivotTablesAndSlicers`
File: `Slicer & Pivot v4.bas`

Creates a PivotTable and a Slicer for every column of the **Tidied Data** sheet. Slicers are positioned in three columns, sorted alphabetically and grouped by prefix (`M -`, `Q -`, `SQ -`). Each group uses its own fill color so you can easily differentiate them. At the end a message box confirms completion.

### `ConnectSlicers_StatusBar_Final`
File: `Connections v2.bas`

Links all slicers in the workbook to every PivotTable on the **PivotTable** sheet. The status bar spinner now reports how many PivotTable/slicer pairs have been processed, including a breakdown of new versus already linked connections. A summary is displayed when finished.

Both macros automatically create the **PivotTable** worksheet if it does not already exist.

## Importing the Modules

1. Open your workbook in Excel (tested with Excel 365 on Windows).
2. Press `Alt`+`F11` to open the VBA editor.
3. Choose **File &gt; Import File...**.
4. Select `Slicer & Pivot v4.bas` and click **Open**.
5. Repeat the import for `Connections v2.bas`.
6. Save the workbook as a macro-enabled file (`.xlsm`).

After importing, run the macros from the **Macros** dialog (`Alt`+`F8`).

## License

This project is licensed under the [GNU General Public License version 3](LICENSE).
