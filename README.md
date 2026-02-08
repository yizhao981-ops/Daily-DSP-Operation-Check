# Daily DSP Operation Check (Streamlit)

## What it does
Upload an Excel file and generate an output workbook with:
- RouteMonitor (color highlighted)
- Summary
- Exceptions
- 3pm check (<50% at/after 3pm ET)
- 6pm check (<80% at/after 6pm ET)
- Meta

## Input Requirements
- Column B = Route
- Column J = Status
- Column L = Status timestamp

Optional:
- Column name contains "Flee" -> FleeName
- Column name contains "Driver" -> DriverName

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
