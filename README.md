# Standalone OCR Picking Ticket Tool

This folder is fully isolated for OCR + Picking Ticket testing and custom development.

## Contents

- `ocr_picking_ticket_standalone.html` - standalone app UI
- `ocr_picking_ticket_standalone.js` - standalone app logic
- `pdf.min.js` - PDF.js runtime
- `pdf.worker.min.js` - PDF.js worker
- `tesseract.min.js` - Tesseract.js runtime
- `html2canvas.min.js` - HTML render helper for PDF export
- `jspdf.umd.min.js` - PDF generation runtime
- `jspdf.plugin.autotable.min.js` - jsPDF table plugin
- `item_code_descriptions.json` - local item code -> description catalog
- `run_standalone.sh` - one-command local launcher
- `run_standalone.bat` - one-click Windows launcher

## Run (Linux/Mac)

```bash
cd standalone/ocr_picking_ticket
./run_standalone.sh
```

Default URL:

- `http://127.0.0.1:8091/` (auto-redirects to app)

Use a custom port:

```bash
./run_standalone.sh 9000
```

## Run (Windows)

- Double-click [run_standalone.bat](run_standalone.bat)
- Or from Command Prompt:

```bat
run_standalone.bat
```

Custom port:

```bat
run_standalone.bat 9000
```

If Python is not installed, `run_standalone.bat` automatically falls back to a built-in PowerShell static server (`run_standalone.ps1`).

## Notes

- Keep all files in this folder together (the HTML references local JS libraries).
- This standalone copy is independent from the main dashboard files.
- Ticket counters are stored in browser `localStorage`.
- The app is HTML/PDF-only and does not require Excel.
- All runtime dependencies are local files in this folder (`pdf.min.js`, `pdf.worker.min.js`, `tesseract.min.js`, `html2canvas.min.js`, `jspdf.umd.min.js`, `jspdf.plugin.autotable.min.js`) so no internet access is required at runtime.
- On **Generate Picking Tickets**, it now:
	- validates item codes against the local item catalog
	- reserves ticket numbers and builds export payload for craft sheets
	- populates material descriptions from local `item_code_descriptions.json`
	- prompts to add code + description when an item code is unrecognized
	- renders craft ticket previews and automatically exports per-craft PDFs
- PDF output can now be generated directly from the standalone app (no Excel required):
	- click `Generate Picking Tickets`
	- when exporting manually, select a main output folder once
	- PDFs are saved into a `Picking Tickets` subfolder
	- one PDF is downloaded per craft with filename:
		- `[picking ticket number]-[Drawing Number]-[sheet number]-R[Revision Number].pdf`
- Item code descriptions are loaded from:
	- `item_code_descriptions.json` (generated from `Picking Tickets Assistant.xlsx` and stored locally in this folder)
	- plus user-added codes saved in browser `localStorage`
	- update this JSON file directly if you want the new codes to be shared across machines/browsers

- OCR capture is column-based: select and drag each required column (`Point Number`, `Item Code`, `Size`, `Quantity`) and set `Expected Item Count`.
- OCR loads the expected number of editable rows; if a column returns fewer/more values, the app pads/truncates and flags the mismatch for manual correction.
- OCR results now include a visible `Material Description` column that auto-fills from the local item code catalog and can be manually edited.
- OCR correction memory is local: when you correct OCR values and generate tickets, those corrections are remembered in browser `localStorage` and auto-applied on future OCR runs.
- If an `Item Code` is blank or dashes in editable OCR results, the app prompts for a usable code. If left blank, it assigns `PLACEHOLDER-CODE-##`.
- Placeholder/prompted item-code values are intentionally not learned by OCR correction memory.
- Drawing details are entered as a single `Drawing Number` field.
- `Starting Ticket No` is editable by the user and auto-increments as tickets are generated.
