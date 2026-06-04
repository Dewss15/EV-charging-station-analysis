# Predictive Analysis of EV Charging Demand and Grid Load Dynamics ⚡️🔋

A comprehensive data mining case study (CE634) that analyzes Electric Vehicle (EV) charging behavior, predicts peak grid load risks, and identifies operational anomalies. 

This repository contains an end-to-end Machine Learning pipeline encompassing data preprocessing, supervised classification, unsupervised clustering, and anomaly detection to support data-driven smart grid management.

## Table of Contents 📚
- [Project Overview](#project-overview)
- [Key Insights & Results](#key-insights--results)
- [Pipeline Architecture](#pipeline-architecture)
- [How to Run the Code](#how-to-run-the-code)
- [Files of Interest](#files-of-interest)
- [Team](#team)

---

## Project Overview 🏗️
As EV adoption accelerates, unmanaged charging—especially fast charging—can destabilize local power grids. This project shifts grid management from reactive to proactive by analyzing a 2,800-record synthetic Kaggle dataset to predict grid stress and identify actionable charging archetypes.

**Core Objectives:**
1. Predict the level of peak load risk (Low, Medium, High).
2. Segment stations by operational behavior for dynamic power allocation.
3. Detect "black swan" charging events that could cause localized blackouts.

---

## Key Insights & Results 📊

![Final Dashboard](final_dashboard.jpg)
*(Dashboard showing Hourly Grid Load, Operational Clusters, and Anomaly Detection)*

* **Supervised Classification (Random Forest):** Achieved a baseline accuracy of 45.18% in predicting Peak Load Risk. We utilized **SMOTE** to handle severe class imbalance. The moderate accuracy highlights that grid failure is highly non-linear; future improvements require integrating external macroeconomic data (e.g., local weather, baseline neighborhood consumption).
* **Feature Importance:** The algorithm identified that while *Total Energy Dispensed* is a major factor, the *Percentage of Renewable Energy Used* acts as the ultimate buffer against grid stress, scoring the highest feature importance.
* **Unsupervised Clustering (K-Means):** Successfully partitioned the dataset into 3 distinct operational clusters (Residential, Urban Fast, Commercial Hubs) using `k=3`, providing actionable archetypes for tailored grid limits.
* **Anomaly Detection (Isolation Forest):** Flagged 84 critical anomalous events (3% of the dataset). These represent severe deviations (e.g., massive power draws during off-peak hours) that mimic real-world hardware malfunctions requiring immediate engineering response.

---

## Pipeline Architecture ⚙️
This case study follows a strict 10-step data science pipeline:
1. **Data Acquisition:** Kaggle dataset ingestion.
2. **Data Cleaning:** Duplicate removal, null checks, datetime parsing.
3. **Feature Engineering:** Deriving `hour_of_day` and `day_of_week` for temporal signals.
4. **Categorical Encoding:** One-hot encoding for `station_type` and `city_zone` (`drop_first=True`).
5. **Target Mapping:** Ordinal encoding for `peak_load_risk` (Low → 0, Medium → 1, High → 2).
6. **Feature Scaling:** `StandardScaler` applied to continuous variables to prevent magnitude bias.
7. **Similarity Analysis:** Pearson correlation and Euclidean distance mapping across city zones.
8. **Classification:** Random Forest with a stratified 80/20 split and SMOTE resampling.
9. **Clustering & Outliers:** K-Means clustering and Isolation Forest anomaly detection.
10. **Visualization:** 1x3 Matplotlib/Seaborn analytical dashboard.

---

## How to Run the Code ▶️

Quick setup to reproduce the environment and run the notebook locally or via Google Colab.

**1. Create a virtual environment (Recommended):**
```bash
python -m venv .venv
source .venv/bin/activate    # Mac/Linux
# OR Windows PowerShell: .venv\Scripts\Activate.ps1
2. Install dependencies:

Bash
pip install --upgrade pip
pip install pandas scikit-learn matplotlib seaborn plotly kagglehub jupyterlab
3. Kaggle authentication (Required to download dataset):

Bash
kagglehub auth
4. Run the notebook:

Local: Run jupyter lab, open data_prep.ipynb, and execute cells in order.

Colab: Upload data_prep.ipynb, run !pip install kagglehub in the first cell, and execute sequentially.

(Note: Model training for the Random Forest with 100 estimators typically takes under 1-2 minutes on a standard machine).

Files of Interest 📁
data_prep.ipynb — The complete end-to-end preprocessing, modeling, and visualization notebook.

ev.csv — Raw dataset loaded during execution (downloaded via Kaggle API).

final_dashboard.jpg — The generated 1x3 analytical dashboard.

DM CASE STUDY.docx — The comprehensive final written report and methodology defense.
Completed as part of the CE634 Data Mining and Data Warehousing coursework at Padre Conceição College of Engineering.
