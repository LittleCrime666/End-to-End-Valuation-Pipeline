# End-to-End Valuation Analysis Pipeline 📈

## Overview
Automated pipeline designed to streamline the investment valuation process for Indonesian public companies (IDX). This tool reduces data collection time by automating the extraction of quarterly financial reports, parsing unstructured Excel data, calculating fundamental ratios, and visualizing insights via Power BI.

## Key Features
* **Web Scraping Automation:** Utilizes **Selenium** to dynamically download financial statements based on user input (Year & Quarter).
* **Data Transformation:** Parses complex Excel financial reports using **Pandas**, handles currency conversion (USD to IDR), and standardizes account names.
* **Financial Modeling:** Automatically calculates key metrics:
    * Valuation: PER, PBV
    * Profitability: ROE, NPM, GPM
    * Solvency: DER
* **Visualization:** Interactive Power BI dashboard with DAX measures for sector comparison.

## Project Structure
* `scarper_lk.py`: Bot for downloading financial reports from IDX website.
* `rekap_fundamental.py`: Script for cleaning data & calculating ratios.
* `Konsolidasi.py`: Merges quarterly data into a master dataset.
* `end-to-end_valuation_analysis.py`: Orchestrator script to run the full ETL process.

## Dashboard Preview
![Dashboard Preview](dashboard_preview.png)
*(Displays financial performance trends and intrinsic value analysis)*

## Tech Stack
* **Language:** Python 3.10+
* **Libraries:** Selenium, Pandas, OpenPyXL, NumPy
* **BI Tool:** Microsoft Power BI (DAX)

## How to Run
1. Install requirements: `pip install selenium pandas openpyxl`
2. Run the orchestrator: `python end-to-end_valuation_analysis.py`
3. Input the target Year and Quarter when prompted.
