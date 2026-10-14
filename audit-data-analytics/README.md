# Audit Data Analytics (free)
pip install pandas numpy scipy openpyxl
python audit_analytics.py
# Audit Data Analytics: Journal Entry Testing

Python tool that tests journal entries for fraud and error indicators.
Uses synthetic data; not real client data.

## What it does
Validates data, sets materiality, runs six tests (Benford's law,
duplicates, round amounts, weekend postings, segregation of duties,
threshold splitting), risk-scores entries and builds a reproducible sample.

## Key results
- 35 possible duplicate payments
- 30 segregation-of-duties breaches
- 1 threshold-splitting group (12 payments)

## Run it
pip install -r requirements.txt
python audit_analytics.py

Step-by-step working: audit_walkthrough.ipynb

## Limitations
Synthetic data; tests flag items for follow-up and are not proof of fraud.
