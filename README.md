# Market Benchmarking Dashboard

## Live Demo
https://market-benchmarking-dashboard.onrender.com

## Executive Summary
This dashboard benchmarks employee pay vs. national market percentiles (P10–P90), adjusts market reference by Area Differential, highlights below-market roles, and estimates the cost to move targeted employees to market reference points (P10/P25/P50). Breakouts provide job-level competitiveness views within a selected Job Famil

## Overview
This project is an interactive **market benchmarking dashboard** built to help compensation and HR teams understand how employee pay compares to external market data and to quantify the **cost of addressing below-market pay**.

The tool mirrors how internal compensation teams evaluate:
- Pay competitiveness by job
- Market percentile positioning
- Risk concentration across roles
- Budget tradeoffs for pay adjustments

This project is intended as a **portfolio demonstration** using fully synthetic data.

---

## Business Problem
Many organizations have never conducted formal market benchmarking, which creates several risks:
- Unknown exposure to under-market pay
- Hiring and retention challenges
- Inefficient or untargeted compensation spend
- Difficulty explaining pay decisions to leadership

This dashboard answers questions such as:
- Which jobs are below market?
- How far below market are they (by percentile)?
- How much would it cost to address pay gaps?
- Which roles should be prioritized?

---

## What the Dashboard Does

### Market Competitiveness Analysis
- Compares **employee pay percentiles** (P10–P90) to **market percentiles**
- Supports Base Pay, Bonus Paid, and Total Compensation
- Adjusts market data using **Area Differentials** (High / Mid / Low)
- Visualizes gaps using percentile line charts and delta tables

### Costing Scenarios
- Calculates the cost to bring employees to:
  - Market P10
  - Market P25
  - Market P50
- Allows filtering by **Year-End Performance Rating** to:
  - Prioritize higher performers
  - Phase compensation investment strategically

### Breakout Analysis
- Job-family–based breakouts independent of top-level filters
- Displays **all job titles** within a selected family
- Shows headcount and competitiveness deltas by percentile
- Enables sorting by severity of market gaps

---

## Key Features
- Interactive Dash (Plotly) web application
- Percentile-based benchmarking (not averages)
- Geo-adjusted market comparisons
- Performance-aware costing scenarios
- Executive-ready visuals and tables
- Designed for explainability and decision support

---

## Data Inputs (Demo / Mock)
This version of the tool uses **synthetic and sanitized data** for demonstration purposes.

### Employee Data (Excel)
- Job Title
- Job Family
- Area Differential
- Base Pay
- Bonus Paid
- Total Compensation
- Year-End Performance Rating

### Market Data (Excel)
- National market percentiles (P10–P90) by Job Title
- Separate percentiles for Base Pay, Bonus, and Total Compensation

> **Note:** No real employee or company data is used in this demo.

---

## Technology Stack
- Python
- Dash (Plotly)
- Pandas / NumPy
- Plotly Graph Objects & Express
- Excel (OpenPyXL)

---

## How to Run Locally

### 1. Clone the repository
```bash
git clone https://github.com/<your-username>/market-benchmarking-dashboard.git
cd market-benchmarking-dashboard

