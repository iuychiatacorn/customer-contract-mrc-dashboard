# Customer Contract & MRC Dashboard

A Streamlit dashboard for exploring:
- customer contract status
- MRR / MRC trends
- project rates
- true-up planning
- hosting data

## Files
- `app.py` - Streamlit app
- `Customer Contract and MRC Tracking (1).xlsx` - workbook used by the dashboard
- `requirements.txt` - Python dependencies

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy publicly with Streamlit Community Cloud
1. Create a new public GitHub repository.
2. Upload `app.py`, `requirements.txt`, and the Excel file to the repo root.
3. Go to Streamlit Community Cloud and connect your GitHub account.
4. Choose the repo, branch, and `app.py`, then deploy.

Every new push to GitHub will update the deployed app.
