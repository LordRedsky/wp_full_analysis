# WP Deep Analysis - Certificate Analysis Addition

## Overview
This project now includes a new certificate analysis feature that performs validation on the Sertifikat file before other processing steps.

## New Component Added

### `analysis_sert.py`
- Performs Certificate Analysis (Tahap 00)
- Analyzes only white/empty-filled NIB columns in the Sertifikat file
- Checks for empty NIB values and duplicate NIB values
- Colors rows red and adds descriptions for invalid entries
- Outputs the analyzed file as "Sertifikat_analysed.xlsx"

## Updated UI Flow in `app.py`
The Streamlit application now includes three processing stages:

1. **Tahap 00 - Analysis Sertifikat**
   - Upload Sertifikat file
   - Run "Proses Analysis Sertifikat"
   - Download the analyzed certificate file

2. **Tahap 01 - Exact Match**
   - Upload Wajib Pajak, Sertifikat (uses analyzed file if available), and Kendali files
   - Run "Proses Exact Match"
   - Download updated files

3. **Tahap 02 - Deep Analysis**
   - Run after Exact Match is completed
   - Run "Proses Deep Analysis"
   - Download final processed files

## Certificate Analysis Logic (Updated Requirements)
- Only processes white/empty-filled NIB columns
- Checks if NIB values exist (not empty)
- Checks for duplicate NIB values
- Colors rows red if NIB is empty or duplicate
- **IMPORTANT**: Only subsequent duplicates (not first occurrence) are marked as red
  - Example: If NIB "1234567890" appears in row 2 and 5, only row 5 will be marked as red
- Adds appropriate descriptions ("NIB kosong" for empty values, "NIB memiliki duplikate" for duplicates)
- Colors affected cells: Nama, NIB, and Luas columns

## Usage
```bash
streamlit run app.py
```

The application maintains all previous functionality while adding the new certificate analysis step as the first processing stage.