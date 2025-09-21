# Payroll-Data-Processing-Demo

## Overview

This is a demo project showcasing how monthly payroll Excel exports can be transformed into a consolidated, analysis-ready dataset using Python (pandas).

The workflow reads several source files (employee codes, benefits, previous salary files), merges them, rearranges columns, computes new metrics, and writes a clean final Excel file.

**Note:**

This repo contains no real data. All file paths and column names are sanitized and generic. The script is provided purely as a demonstration of workflow automation with pandas.

## Data flow

1. Input files:
- Salaries.xls
- Codes.xls
- Benefits.xlsx
- Salaries_old.xls

2. Processing steps:
- Merge data
- Reorder columns
- Compute derived columns

3. Output
- Salaries_Final.xlsx

## Key Features

- Multiple joins with validate="m:1" to ensure clean merges
- Column reordering with DataFrame.insert for clear layout
- Derived payroll metrics:
  - Net salary
  - Staff benefits
  - Employee/Employer taxes
  - Gross totals
- Small helper function for systematic column insertion
- Final export to a single Excel file

## Requirements

This project depends on a few common Python libraries:

- pandas
- numpy
- openpyxl

Install them with:

```bash
pip install -r requirments.txt
```
