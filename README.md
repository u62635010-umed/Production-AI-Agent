#  Production AI Agent

A Streamlit application that extracts information from handwritten  production reports, lets an operator review and correct the results, and writes confirmed records to an Excel workbook.

## What it does

- Accepts production report photographs and prepares them for image analysis.
- Uses the Gemini API to extract report fields and hourly or shift entries.
- Provides a review step before any confirmed record is written.
- Validates values, flags missing fields, and checks for duplicate images or conflicting records.
- Maintains production history, an audit sheet, summaries, charts, and local workbook backups.
- Includes synthetic demo data for exploring the review interface; demo records are not saved as production records.

## Requirements

- Python 3.11 or later
- A Gemini API key for image extraction

## Set up on Windows

From the project directory, create and activate a virtual environment, then install the dependencies:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Create a local environment file and add your Gemini API key:

```powershell
Copy-Item .env.example .env
```

Set `GEMINI_API_KEY` in `.env`. The example file contains a placeholder only. `GEMINI_MODEL` and `SAVE_UPLOADS` can be adjusted there as needed. Keep `.env` private; Git ignores it.

Prepare a local workbook from the clean template:

```powershell
Copy-Item Templates/production_sheet.template.xlsx Templates/production_sheet.xlsx
```

The included workbook is empty. The local production workbook and generated outputs are excluded from Git.

## Run

```powershell
streamlit run app.py
```

Upload a report photograph, review the extracted fields and rows, resolve any flagged values, and select **Confirm and Update Excel** to save the confirmed record to the local workbook.

## Data and privacy

Production workbooks, generated backups and logs, uploaded photographs, `.env`, virtual environments, and cache folders are excluded by `.gitignore`. Do not add production data, photographs, or API keys to the repository. Uploaded originals are not retained by default; the application can retain them locally when that option is enabled.

## Tests

Run the automated suite with:

```powershell
python -m pytest
```
