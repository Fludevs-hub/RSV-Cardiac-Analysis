# RSV Cardiac Disease Analysis

This project analyzes hospital-based SARI/ILI surveillance data to assess the burden and clinical impact of respiratory syncytial virus (RSV) among children with cardiac disease in Kenya.

## Overview

The script reads a Stata dataset, identifies children with cardiac disease, restricts the analysis to children with RSV RT-PCR results, compares RSV-positive and RSV-negative groups, and produces publication-ready tables for a conference or report.

The analysis focuses on three main questions:

1. How common is RSV among children with cardiac disease?
2. Is RSV associated with greater illness severity in this group?
3. What are the outcomes and resource-use differences by RSV status?

## Project purpose

The workflow in the script is designed for a secondary analysis of KEMRI/CGHR surveillance data and supports a conference-style output with tables on:

- baseline characteristics
- RSV prevalence by subgroup
- severity and outcomes
- hospital resource use
- interaction analysis for cardiac disease as an effect modifier

## Input data

The main input file is:

- `SARI_ILI_Dataset.dta`

This is a Stata dataset containing surveillance records for SARI/ILI cases. The script expects the data file in the same working directory.

## What the script does

The main analysis script performs the following steps:

### 1. Loads and cleans the dataset

- reads the Stata file using pandas
- replaces the sentinel missing code `9999` with `NaN`
- keeps a record of the study flow (how many records remain after each filter)

### 2. Defines cardiac disease

The script does not rely only on a single coded variable. It builds the cardiac-disease variable from:

- specific cardiac disease indicators already in the data
- a heart-failure flag created from keyword matches in free-text diagnoses
- broader keyword search terms such as `heart`, `cardi`, `rhd`, and `valv`

This allows the analysis to identify cardiac disease even when the dataset uses free-text diagnosis fields.

### 3. Restricts to RSV-tested children

The analysis keeps only records with a valid RSV RT-PCR result. It then derives variables for:

- RSV status (`rsv`)
- age groups
- sex
- enrollment site
- malnutrition status
- hypoxemia
- ICU / acute care admission
- mechanical ventilation
- in-hospital death
- composite severe outcome
- prior hospitalization
- length of stay

### 4. Creates the study population for cardiac disease

From the RSV-tested data, the script isolates the subset of children with cardiac disease and splits them into:

- RSV-positive
- RSV-negative

This allows direct comparison of baseline characteristics, severity markers, outcomes, and resource use.

### 5. Builds statistical summary tables

The script constructs multiple data tables, including:

- Table 1: baseline characteristics by RSV status
- Table 2: RSV prevalence by subgroup
- Table 3: severity and outcomes by RSV status
- Table 4: hospital resource use
- Table S1: interaction analysis to assess whether cardiac disease modifies the RSV-outcome association

These tables use statistical tests such as:

- Fisher's exact test
- chi-square tests
- Mann-Whitney U tests
- logistic regression models with adjusted odds ratios

### 6. Exports results

The script exports the results to:

- `outputs/RSV_Cardiac_Tables.docx`
- `outputs/RSV_Cardiac_Tables.xlsx`

The Word document is formatted for conference or poster-style presentation, and the Excel file stores each table in its own sheet.

### 7. Runs validation checks

Before printing results, the script performs a few assert-based checks to ensure the derived tables and counts are internally coherent. This helps prevent publication of inconsistent output.

## Dependencies

The analysis requires Python packages including:

- pandas
- numpy
- scipy
- statsmodels
- python-docx
- openpyxl

## How to run

From the project folder:

```bash
python main.py
```

or, if using `uv`:

```bash
uv run main.py
```


## Important notes from the analysis

The script includes clinical and methodological safeguards:

- cardiac disease is built from coded variables and free-text diagnosis fields
- heart-failure detection is based on keyword searches and should be manually reviewed for false positives
- prior hospitalization is treated as a descriptive proxy and not as readmission
- adjusted models are only fit when each RSV group has enough events
- ICU/acute care day totals are reported as recorded and not imputed

## Output folder

The script saves outputs in the `outputs` directory. If the folder does not exist, it should be created automatically by the program logic before writing files.

## Project status

This project is a focused epidemiologic analysis script for generating conference-ready summary tables and resource-use estimates for RSV in children with cardiac disease.
