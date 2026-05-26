# DMS/Tally Comparator (Privacy-first)

Compare two Excel exports—**DMS** and **Tally**—and generate reports that highlight **matched**, **unmatched**, and **mismatched** values.

## Features

- **Upload** DMS and Tally spreadsheets (`.xlsx`, `.xls`, `.csv`)
- **Automatic column cleanup** (drops fully blank / all-zero columns)
- **Column mapping** to define which fields should be compared
- **Generate report outputs**:
  - On-screen mismatch table
  - **PDF report**
  - **Excel comparison file**
- **Privacy-first**: files are processed in your browser; nothing is uploaded

## How it works (high level)

1. Load your **DMS** and **Tally** files.
2. Choose (or suggest) a set of **mappings** between DMS columns and Tally columns.
3. The **first mapping acts as the key** used to align records.
4. The remaining mapped columns are compared to detect differences.

## Inputs / Outputs

### Inputs

- Two spreadsheet files:
  - **DMS file**
  - **Tally file**
- Optional mappings between columns (used for matching + comparison)

### Outputs

- **PDF report** (`dmstally-report-YYYY-MM-DD.pdf`)
- **Excel comparison file** (`dmstally-comparison-YYYY-MM-DD.xlsx`)
- A mismatch table shown in the UI

## Privacy

This application is designed to work **entirely in the browser**. Uploaded files are read locally and are **not stored or sent to a server**.

## Usage

1. Install dependencies:
   ```bash
   npm install
   ```
2. Start the dev server:
   ```bash
   npm run dev
   ```
3. Open the app in your browser and follow the steps:
   - Load **DMS** Excel
   - Load **Tally** Excel
   - Click **Suggest mappings** (optional) or add your own mappings
   - Click **Generate report**
   - Download **PDF** and/or **Excel**

## Local Development

- Build:
  ```bash
  npm run build
  ```
- Preview production build:
  ```bash
  npm run preview
  ```
