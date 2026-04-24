# Copper Forecast System

Standalone Streamlit app for copper bar forecasting and replenishment planning.

## What It Does

- Imports the current copper review workbook
- Tracks active copper sizes in pounds
- Blends plant inventory, distribution inventory, and inbound supply
- Uses reorder point plus safety stock logic
- Prefers direct mill replenishment from `Tecnofil`
- Allows manual demand and sourcing overrides
- Highlights shortage, distribution pull, and excess-risk items

## Run It

Open a terminal in this folder and run:

```powershell
python -m streamlit run app.py
```

Then open the local URL shown in the terminal.

## First Use

Local use:

1. Open the `Data Import / Refresh` tab
2. Click `Load default workbook: March Copper Review Final.xlsx`
3. Review `Dashboard`, `Supply Plan`, and `Reorder Recommendations`

Cloud use:

1. Deploy the app from this repo
2. Open the deployed app
3. Go to `Data Import / Refresh`
4. Upload your copper workbook manually
5. Review `Dashboard`, `Supply Plan`, and `Reorder Recommendations`

## Project Files

- `app.py`: Standalone copper forecasting app
- `requirements.txt`: Minimal dependencies
- `data/`: Saved import snapshot and manual overrides

## GitHub Note

The local workbook `March Copper Review Final.xlsx` is excluded from Git by default so you can keep sensitive planning data out of the repository. Upload the workbook through the app after deployment.
