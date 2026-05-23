# School ERP Data Formatter

Lightweight desktop tool to clean, normalise and export student data into a 19-column ERP-friendly format.

## Key Features
- Auto column detection and manual column mapping UI
- Format Customizer: per-file DOB parsing (source → desired) and address formatting
	- Address options: Keep spaces or Remove spaces (radio)
- Data cleaning rules: duplicate removal, name normalization, phone validation (invalid → empty), DOB defaulting (empty → `0`)
- Per-class grouping and export to XLSX/CSV; `roll_number` values are concrete and reset starting at 1 per exported file
- Presets: save/load presets (max 7, FIFO). Inline presets panel supports Load / Rename / Delete
- Removed features: theme (light/dark) toggle and undo functionality (intentionally removed)

## Requirements
- Python 3.9+ (uses virtualenv in repository: `Env1`)
- Packages: `customtkinter`, `pandas`, `openpyxl`

Install dependencies (from project root):

```powershell
Env1\Scripts\activate.ps1
pip install customtkinter pandas openpyxl
```

## Run
From the project root:

```powershell
Env1\Scripts\activate.bat   # or Activate.ps1
python apps\emis_app.py
```

The app opens a GUI. Typical workflow:
1. Import a source file (XLSX/CSV)
2. Step 2 — Map columns and use the *Format Customizer* to set DOB formats and address rules
3. Step 3 — Data cleaning preview
4. Step 4 — School settings, save/load presets, generate ERP format
5. Export per-class files (XLSX/CSV)

## Presets
- Presets are stored at `~/.school_erp_formatter/presets.json` and limited to 7 entries. When adding the 8th preset, the oldest is removed (FIFO). Use the inline panel in Step 4 to load, rename or delete presets.

## Notes & Troubleshooting
- If the window does not stay maximized on start, the app attempts a delayed maximize; you can tweak the delay in `apps/emis_app.py` (`after()` call).
- Preset/setting persistence uses the home directory defined by `Path.home()`.

## License
This repository contains custom code; adjust licensing as needed.

---
File: [apps/emis_app.py](apps/emis_app.py#L1)
