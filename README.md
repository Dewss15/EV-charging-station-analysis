# Predictive Analysis of EV Charging Demand and Grid Load Dynamics (CE634 Case Study) ⚡️🔋

A reproducible case study that analyzes EV charging behavior and predicts peak grid load risk to support operational decisions. This repo contains the preprocessing and modeling pipeline; the first half of the pipeline (Steps 1–7) is complete and ready for handoff.

![Python](https://img.shields.io/badge/-Python-3776AB?logo=python&logoColor=white) ![Pandas](https://img.shields.io/badge/-Pandas-150458?logo=pandas&logoColor=white) ![Scikit-learn](https://img.shields.io/badge/-Scikit--Learn-F7931E?logo=scikit-learn&logoColor=white) ![Seaborn](https://img.shields.io/badge/-Seaborn-4B8BBE?logo=seaborn&logoColor=white) ![Kaggle](https://img.shields.io/badge/-Kaggle-20BEFF?logo=kaggle&logoColor=white)

---

## Table of Contents 📚
- [Project Architecture](#project-architecture)
- [My Completed Work (Steps 1–7)](#my-completed-work-steps-1-7)
- [Handoff / Next Steps for Partner (Steps 8–10)](#handoff--next-steps-for-partner-steps-8-10)
- [How to Run the Code](#how-to-run-the-code)
- [Files of Interest](#files-of-interest)
- [Contact & Notes](#contact--notes)

---

## Project Architecture 🏗️
This case study follows a 10-step pipeline designed for reproducible data science delivery:

1. Data acquisition (Kaggle dataset)
2. Initial data exploration & cleaning
3. Feature engineering (time features)
4. Encoding categorical variables
5. Scaling numerical features (StandardScaler)
6. Similarity & dissimilarity analyses (correlation, distance)
7. Baseline supervised model (Random Forest classification)
8. Unsupervised analysis: K-Means clustering across zones
9. Anomaly detection: Isolation Forest for outlier identification
10. Final dashboard & Key Findings (1x3 panels + written report)

---

## My Completed Work (Steps 1–7) ✅
Brief summary of what's already been implemented and validated in the notebook:

- Data ingestion and cleaning: duplicates removed, null checks performed, and datetime parsed.
- Feature engineering: `hour_of_day` and `day_of_week` derived from `date_time` for temporal signal.
- Categorical encoding: `station_type` and `city_zone` one-hot encoded (drop_first=True).
- Target encoding: `peak_load_risk` mapped to ordinal labels (Low → 0, Medium → 1, High → 2).
- Scaling: numerical features standardized using `StandardScaler` (`energy_dispensed_kwh`, `avg_charging_duration_minutes`, `grid_load_mw`, `vehicles_charged`, `renewable_energy_used_percent`).
- Data leakage protection: identifier column `record_id` removed from model features.
- Baseline model: `RandomForestClassifier(n_estimators=100)` trained with stratified 80/20 split. Baseline accuracy: **~45%** (diagnosed as severe class imbalance — next steps should address rebalancing or sampling).

Notes:
- The preprocessing and modeling pipeline is implemented in [data_prep.ipynb](data_prep.ipynb).
- The raw CSV is stored/loaded as [ev.csv](ev.csv) during notebook execution.

---

## Handoff / Next Steps for Partner (Steps 8–10) 🔁
Below is a targeted checklist for the teammate taking the project forward. Each item includes suggested code references and acceptance criteria.

- [ ] Step 8 — K-Means Clustering
  - Task: Run K-Means on the zone-level profiles (use the same scaled features used for distance analysis).
  - Deliverable: `kmeans_labels.csv` per zone, elbow plot to pick K, and short paragraph with interpretation.

- [ ] Step 9 — Isolation Forest Outlier Detection
  - Task: Fit `IsolationForest` on session-level data to flag anomalous charging events.
  - Deliverable: CSV of flagged outliers, example anomalous records, and brief remediation suggestions.

- [ ] Step 10 — Final 1x3 Dashboard Visualizations & Key Findings Report
  - Task: Create a 1x3 dashboard (single-row, three panels) showing:
    1. Temporal risk distribution (hourly heatmap / line)
    2. Zone clusters (K-Means colored map / scatter)
    3. Anomalies vs normal sessions (Isolation Forest overlay)
  - Deliverable: `final_dashboard.png` (or interactive `final_dashboard.html`) and `STEP_10_Key_Findings.md` (2–3 pages, bullet summary + recommendations).

Helpful tips:
- Address class imbalance before retraining for performance improvements: consider `SMOTE`, class-weight adjustments, or targeted sampling.
- For dashboarding, Plotly Dash or Streamlit work well for interactive outputs; Matplotlib/Seaborn + `plt.tight_layout()` is fine for static images.

---

## How to Run the Code ▶️
Quick setup to reproduce the environment and run the notebook locally or in Colab.

1. Create a virtual environment (recommended):

```bash
python -m venv .venv
source .venv/Scripts/activate    # Windows PowerShell: .venv\Scripts\Activate.ps1
```

2. Install dependencies:

```bash
pip install --upgrade pip
pip install pandas scikit-learn matplotlib seaborn plotly kagglehub jupyterlab
```

3. Kaggle authentication (required to download dataset):

```bash
kagglehub auth
```

4. Run the notebook:

Option A — Local Jupyter / VS Code:

```bash
jupyter lab
# open data_prep.ipynb and run cells in order
```

Option B — Google Colab:

1. Upload `data_prep.ipynb` to Colab.
2. Install dependencies in the first cell (use `!pip install ...`).
3. Run cells sequentially; the dataset download cell will pull from Kaggle via `kagglehub`.

Expected runtime notes:
- Preprocessing and EDA: a few seconds to a minute depending on I/O.
- Model training (Random Forest, 100 trees): typically under 1–2 minutes on a laptop.

---

## Files of Interest 📁
- [data_prep.ipynb](data_prep.ipynb) — main notebook (Steps 1–7 implemented)
- [ev.csv](ev.csv) — dataset loaded by the notebook (downloaded via Kaggle)
- [README.md](README.md) — this document

---

## Contact & Notes ✉️
Partner, welcome! I've left inline comments in `data_prep.ipynb` where I think you'll pick up work for Steps 8–10. If you want, I can also:

- add a starter `kmeans.ipynb` and `isolation_forest.ipynb` template,
- create a minimal Streamlit dashboard scaffold for the 1x3 visualization.

Happy to help with any of the handoff tasks — ping me when you want one of the templates created.

---

Good luck — this project is in a great spot to generate actionable operational insights for EV-grid planning! 🚀
