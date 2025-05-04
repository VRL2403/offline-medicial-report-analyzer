# Medical Report Analyzer
A comprehensive offline Streamlit application that processes medical reports in PDF format, extracts structured data, converts it to Excel, and provides analysis based on Indian medical standards.
## Features

PDF Upload & Extraction: Accepts PDF uploads containing medical reports and extracts text/data
Data Conversion: Converts extracted data to Excel format with proper structure
Comprehensive Analysis: Analyzes a wide range of medical tests including CBC, Lipid Profile, Liver Function, etc.
Indian Medical Standards: Analysis follows Indian reference ranges
Report Generation: Generates summary reports highlighting abnormal values and potential concerns
Visualizations: Creates visual representations of key health parameters
Offline Operation: Functions entirely offline with no external API dependencies

# Installation
## Prerequisites

Python 3.8 or higher

## Setup

Clone this repository or download the files
Create a virtual environment (recommended):
python -m venv venv

Activate the virtual environment:

- Windows: venv\Scripts\activate
- Mac/Linux: source venv/bin/activate


Install the required dependencies:
pip install -r requirements.txt


# Usage

1. Run the Streamlit application:
streamlit run medical_report_analyzer.py

2. The application will open in your default web browser
3. Upload a medical report PDF file
4. View the analysis, visualizations, and download the Excel report

# Supported Test Categories
## Hematology

- Complete Blood Count (CBC)
- Differential Count
- Platelets

## Biochemistry

- Liver Function Tests (SGPT, Alkaline Phosphate)
- Kidney Function Tests (Creatinine)
- Lipid Profile (Cholesterol, Triglycerides, HDL, LDL)
- Thyroid Function (TSH)

## Other Tests

- Cancer Markers (PSA)
- Vitamins (B12, D3)
- Diabetes Markers (HbA1c)
- ECG Findings

# Reference Ranges
The application uses reference ranges based on Indian medical standards. These ranges are used to categorize test results as normal, borderline, or abnormal.
# Customization
You can customize the reference ranges and test patterns by modifying the REFERENCE_RANGES and TEST_PATTERNS dictionaries in the code