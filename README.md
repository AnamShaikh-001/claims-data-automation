# Claims Data Consolidation & Deduplication Automation

## Project Overview
This project automates the consolidation of multiple loss-run files into a single structured Excel file.

It eliminates manual effort by:
- Extracting summary details (Insured Name, Policy Number, Valuation Date)
- Identifying and cleaning claim-level data
- Skipping files with no valid claim tables
- Removing duplicate claim numbers
- Exporting a final ready-to-use dataset

---

## Tools & Technologies
- Python
- Pandas
- OS Module
- Excel

---

## Business Problem
Manually combining multiple loss-run reports was:
- Time-consuming
- Error-prone
- Difficult to standardize

---

## Solution
Developed a Python automation that:
1. Loops through multiple files in a folder
2. Dynamically detects claim table structure
3. Cleans and standardizes column names
4. Filters valid claim rows
5. Removes duplicate claims
6. Generates a final consolidated Excel report

---

## Business Impact
- Reduced manual processing time significantly
- Improved data accuracy
- Created a reusable reporting solution

---

## ▶️ How to Run

1. Place input files inside:

   data/input

2. Ensure the project structure is:

   project-folder
   │── data
   │   ├── input
   │   └── output
   │
   │── Hartford Merger.py

3. Run the script:

   python "Hartford Merger.py"
│   └── claims_merge.py
