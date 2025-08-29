# Paddy Collection → Excel App

A simple Streamlit app to collect Paddy Collection entries and append them to an Excel file `paddy_collection.xlsx`.

## Prerequisites

- Python 3.10 or newer

## Setup

1. (Optional) Create and activate a virtual environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Run the App

```bash
streamlit run app.py
```

- Your browser should open automatically at `http://localhost:8501`.

## Using the App

- Fill out all fields in the form and click Submit.
- Each valid submission appends exactly one new row to `paddy_collection.xlsx` on the `Submissions` sheet.
- If the file does not exist yet, it will be created with the correct headers on first save.
- Below the form, you can preview the latest 20 submissions (newest first) and download the full Excel file via the "Download Excel" button.

## File Location

- The Excel file `paddy_collection.xlsx` is saved in the same folder as `app.py`.

## Notes

- Concurrent writes are protected using a lock file `paddy_collection.xlsx.lock`.
- `Truck Number` and `SR Number` are saved as text to preserve leading zeros.
- `submitted_at` is recorded in your local time, ISO format.