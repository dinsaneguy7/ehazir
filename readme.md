# Ehazir Shambles

<p align="center">
  <img src="apps/data_cleaner_exe_app/EhazirShambles_logo.png" alt="Ehazir Shambles logo" width="180" />
</p>

<p align="center">
  A modern desktop data cleaner for student records, ERP mapping, DOB formatting, and class-wise exports.
</p>

<p align="center">
  <a href="apps/data_cleaner_exe_app/EhazirShambles.exe"><img src="https://img.shields.io/badge/Download_EXE-0F766E?style=for-the-badge&logo=windows&logoColor=white" alt="Download EXE" /></a>
  <a href="apps/data_cleaner_formatter.py"><img src="https://img.shields.io/badge/View_Source-2563EB?style=for-the-badge&logo=python&logoColor=white" alt="View Source" /></a>
  <a href="#quick-start"><img src="https://img.shields.io/badge/Quick_Start-F59E0B?style=for-the-badge&logo=rocket&logoColor=white" alt="Quick Start" /></a>
  <a href="#folder-guide"><img src="https://img.shields.io/badge/Folder_Guide-7C3AED?style=for-the-badge&logo=files&logoColor=white" alt="Folder Guide" /></a>
</p>

<p align="center">
  <a href="apps/data_cleaner_exe_app/EhazirShambles.exe">Open the EXE</a> ·
  <a href="apps/data_cleaner_exe_app/EhazirShambles_logo.png">View Logo</a> ·
  <a href="apps/data_cleaner_exe_app/ehazirshambles.ico">View Icon</a>
</p>

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Download Buttons](#download-buttons)
- [Quick Start](#quick-start)
- [How To Use](#how-to-use)
- [Folder Guide](#folder-guide)
- [Download The EXE](#download-the-exe)
- [Build It Yourself](#build-it-yourself)
- [Troubleshooting](#troubleshooting)
- [Notes](#notes)

---

## Overview

Ehazir Shambles is a Windows desktop app that cleans student data and turns it into ERP-ready output. It detects columns, normalizes names, handles DOB formatting, validates phone numbers, separates students by class, and exports final files in XLSX or CSV format.

## Features

- Auto column detection with manual mapping
- DOB parsing with source and desired format selection
- Empty DOB values can be filled as `0`
- Invalid phone numbers are cleared to plain empty values
- Per-class grouping and export
- Preset save, load, rename, and delete
- Startup splash with logo and icon branding

## Download Buttons

Use these if you want the README to feel like a landing page:

<table>
  <tr>
    <td><a href="apps/data_cleaner_exe_app/EhazirShambles.exe"><strong>Download EXE</strong></a></td>
    <td>Ready-to-run Windows build.</td>
  </tr>
  <tr>
    <td><a href="apps/data_cleaner_formatter.py"><strong>Open Source File</strong></a></td>
    <td>Main Python app source.</td>
  </tr>
  <tr>
    <td><a href="#folder-guide"><strong>See Folder Guide</strong></a></td>
    <td>Quick path map of the project.</td>
  </tr>
</table>

## Quick Start

1. Open [EhazirShambles.exe](apps/data_cleaner_exe_app/EhazirShambles.exe).
2. Import your student file in XLSX, XLS, or CSV format.
3. Check the detected columns and adjust mappings if needed.
4. Set DOB and address rules in the format customizer.
5. Clean the records, configure school settings, and export the final files.

## How To Use

### Step 1. Import File
Choose the student file you want to clean. The app supports XLSX, XLS, and CSV.

### Step 2. Review Column Detection
Check the automatically detected mapping. Correct any field that needs a manual override.

### Step 3. Adjust Format Rules
Use the DOB format controls and address options to match your source file.

### Step 4. Clean Data
Run cleaning to remove duplicates, normalize names, and convert invalid values.

### Step 5. Set Settings
Enter school name, domain, output format, and preset preferences.

### Step 6. Export
Choose your output folder, separate by class, and export the final files.

## Folder Guide

| Path | Purpose |
| --- | --- |
| [apps/data_cleaner_formatter.py](apps/data_cleaner_formatter.py) | Main source file for the desktop app. |
| [apps/data_cleaner_exe_app/EhazirShambles.exe](apps/data_cleaner_exe_app/EhazirShambles.exe) | Packaged Windows build you can run directly. |
| [apps/data_cleaner_exe_app/ehazirshambles.ico](apps/data_cleaner_exe_app/ehazirshambles.ico) | App icon used in the title bar and EXE. |
| [apps/data_cleaner_exe_app/EhazirShambles_logo.png](apps/data_cleaner_exe_app/EhazirShambles_logo.png) | Startup splash/logo image. |
| [apps/data_cleaner_exe_app/build/](apps/data_cleaner_exe_app/build/) | PyInstaller build workspace. |
| [html/](../html/) | Additional HTML assets in the workspace. |
| [.git/](.git/) | Git metadata for the repository. |

## Download The EXE

If you do not want to use GitHub Releases yet, follow these steps directly in the GitHub repository:

1. Open the GitHub repository page for this project.
2. Go to the folder [apps/data_cleaner_exe_app](apps/data_cleaner_exe_app/).
3. Click [EhazirShambles.exe](apps/data_cleaner_exe_app/EhazirShambles.exe) in the file list.
4. On the file page, use the download option shown by GitHub to save the file to your computer.
5. If GitHub opens the file view instead of downloading it, choose the raw/download button from the page and save the file manually.
6. Move the downloaded EXE to any folder you like on Windows.
7. Double-click the EXE to launch the app.

You do not need GitHub Releases for this workflow. The EXE can be downloaded straight from the repository file page as long as the file is present in the repo.

### Direct links

- [Download the EXE](apps/data_cleaner_exe_app/EhazirShambles.exe)
- [Open the source file](apps/data_cleaner_formatter.py)
- [Open the logo image](apps/data_cleaner_exe_app/EhazirShambles_logo.png)
- [Open the icon file](apps/data_cleaner_exe_app/ehazirshambles.ico)

## Build It Yourself

If you want to rebuild the EXE yourself, use the Python environment in `Env1` and install these packages:

- `customtkinter`
- `pandas`
- `openpyxl`
- `pyinstaller` for packaging the app again

Example build flow:

```powershell
Env1\Scripts\activate.ps1
pip install customtkinter pandas openpyxl pyinstaller
python apps\data_cleaner_formatter.py
```

If you build with PyInstaller, keep the icon and logo bundled so the EXE stays branded.

## Troubleshooting

- If the window does not stay maximized, close it and reopen it. The startup maximize is intentionally delayed.
- If the splash image is missing, confirm that [EhazirShambles_logo.png](apps/data_cleaner_exe_app/EhazirShambles_logo.png) is present beside the packaged EXE.
- If you rebuild the app, make sure the icon and splash image are included in the PyInstaller command.

## Notes

- This app is built for desktop use on Windows.
- The source version lives in [apps/data_cleaner_formatter.py](apps/data_cleaner_formatter.py).
- The packaged build lives in [apps/data_cleaner_exe_app/EhazirShambles.exe](apps/data_cleaner_exe_app/EhazirShambles.exe).
