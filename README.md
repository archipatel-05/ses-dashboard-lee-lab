# SES Indices Dashboard

Quick Streamlit dashboard to run the existing SES indices pipeline for a single patient/address.

How to run

1. Install requirements (in your project folder):

```bash
pip install -r requirements.txt
```

2. Run the Streamlit app:

```bash
streamlit run dashboard.py
```

Notes
- The app imports the existing script `ses_indices_from_addresses_v22_final_for_real.py` and calls its `geocode_addresses` and `build_outputs` functions.
- The pipeline will look for reference data inside the project folder. Ensure the data subfolders/files from the repo are present.
- Geocoding may use the Census live geocoder (internet required) or fall back to offline heuristics.
- The pipeline now creates local parquet caches for large reference tables (file.parquet alongside CSV/XLSX) and supports a SQLite geocode cache (use `geocode_cache.sqlite`). These speed repeated runs.
