ExcelTool — WinIBW4
=====================

This README documents the `excelTool` dialog and scripts shipped with WinIBW4. It is intended as a concise developer and user supplement to the existing wiki article:

- Primary reference (existing): https://wiki.k10plus.de/spaces/K10PLUS/pages/64225386/Excel-Tabelle+erstellen

Purpose
-------
The excelTool generates a CSV/TSV/text export from the current selection in WinIBW4 using a configuration table that describes which PICA fields/subfields to extract and how to format them.

Contents of this folder
-----------------------
- Files/ — UI assets used by the dialog (HTML, client JS, CSS). Key files:
	- Files/excelTool/dialogExcelTabelle.html — dialog UI shown in WinIBW4
	- Files/excelTool/excelTabelle.css — styles
	- Files/excelTool/excelTabelle.js — client-side dialog logic (runs in the WebView/dialog)
- Scripts/ — JScript backend handlers called by the dialog
	- Scripts/dialogExcelTabelle.js — main backend: read/write config files, parse definitions, build CSV output
- README.md — this file

Quick User Guide
----------------
1. Open the dialog: run the `excelTabelle()` script in WinIBW4 (menu or script runner).
2. Choose the table type: either the built-in `Standardtabelle` or a user configuration saved under ProfD\\user.
3. Edit the configuration in the "Auswahl" textarea to control which fields are exported. Each non-comment line is a mapping of column name to definition, e.g.:
	 Erscheinungsjahre: 011@
	 - Lines beginning with `//` are comments.
	 - If a definition contains a colon (`:`) the left side is the header and the right side the mask/definition.
4. Use "Speichern als" to save a personal configuration file under ProfD\\user. Use "Öffnen" to load a file from ProfD\\user.
5. Click the export button to generate the CSV/TSV. The generated file is written under ProfD\\listen and opened automatically.

Behavior and defaults
---------------------
- Last-used user config filename is stored in profile key `Exceltool.LastUserFile`.
- Default user config name: `csvDefinitionUser.txt`.
- Separator/format choices are controlled by profile keys (`Separator`, `Trennzeichen`).

Configuration syntax (summary)
------------------------------
The configuration file is parsed line-by-line. Basic rules:
- Blank lines and lines starting with `//` are ignored.
- A line may be either a simple tag-only entry or a `Header: definition` pair.
	- Example simple entry: `011@` — the parser accepts tag-only masks and will extract the whole field.
	- Example headered entry: `Erscheinungsjahre: 011@` — produces a column titled "Erscheinungsjahre".
- Definitions can include subfield selectors, quoted prefixes/suffixes, and OR/AND partitions. See the wiki for the full grammar.

Notes about whole-field support
------------------------------
- Recent fixes: the parser now accepts tag-only masks (e.g. `011@`) and will return the raw field contents (including the MARC subfield marker character used internally, e.g. `ƒa1934`) when configured to extract the whole field.
- If you want only a specific subfield, use the `$` notation (e.g. `$a`) or the quoted form `\"prefix$a\"`.

Developer notes
---------------
- Client API:
	- `Files/excelTool/excelTabelle.js` uses helper functions `runScript()`, `getFileContent()`, `getProfileString()` — these are provided by the WinIBW dialog host.
	- The client sends textarea contents and file names via hidden form inputs (e.g. `hid_idAuswahlZeilen`, `idSaveAsFileName`) before calling backend handlers.
- Backend API (JScript `Scripts/dialogExcelTabelle.js`): key functions
	- `__excelWriteAuswahlAs(o)` — writes textarea content to `ProfD\\user\\<fname>` and stores last filename in profile
	- `__excelLoadFilesInUser()` — lists files in `ProfD\\user`
	- `__excelReadUserFile(o)` — reads a specific user file
	- `__readControl(inp, must)` — parses the configuration text into header/definition arrays
	- `__replaceDefinitionsWithLookup(content)` — resolves definition lookups against the internal csvDefinitions
	- `__createCtrlArray(content)` + helper parsing functions (`__getSpecial`, `__getTagInfos`, `__orPartitions`, `__sbfPart`, etc.) — construct control objects consumed by the exporter
	- `__excelWriteCSV(o)` — main exporter that iterates records and writes the CSV

- Key profile keys used:
	- `Exceltool.LastUserFile` — last personal config filename
	- `Exceltool.Trennzeichen`, `Exceltool.Separator` — output formatting

Troubleshooting
---------------
- "Speichern" button does nothing: ensure the dialog's hidden inputs exist before calling `runScript` (client-side must sync textarea to `hid_idAuswahlZeilen`, and `idSaveAsFileName` must be present). The client script now creates/syncs these elements automatically on load and before save.
- Parsing errors like "Die Definitionsdatei ist leer" mean the file contained no usable lines — check for stray whitespace or unexpected characters.
- If the export yields unexpected content (e.g. subfields split into many fragments), inspect the original MARC record block for the exact line representing the field (the backend extracts based on the field tag and internal subfield marker). Provide a sample record and configuration line when filing a bug.

Further reading
---------------
- Full grammar, examples and background: https://wiki.k10plus.de/spaces/K10PLUS/pages/64225386/Excel-Tabelle+erstellen
- WinIBW4 dialog API and scripting: refer to the WinIBW developer docs (internal)

Contributing changes
--------------------
- Make client-side changes in `Files/excelTool/excelTabelle.js` and UI in `Files/excelTool/dialogExcelTabelle.html`.
- Make backend changes in `Scripts/dialogExcelTabelle.js`. Keep ES3/JScript compatibility in mind.
- Run iterative tests inside the WinIBW4 environment — many helpers (`runScript`, `getFileContent`, `getProfileString`) are provided by the host and cannot be executed outside WinIBW.

Contact / Reporting bugs
------------------------
Open an issue in the repository or copy a minimal reproducer (one sample record block and the exact configuration lines) into the bug report so the parser logic can be adjusted.


