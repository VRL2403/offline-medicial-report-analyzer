import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import base64
import pdfplumber
from PIL import Image
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os

# Set page configuration
st.set_page_config(
    page_title="MedReport Analyzer",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Apply custom styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 1rem;
        font-weight: bold;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #0D47A1;
        margin-bottom: 1rem;
    }
    .section-header {
        font-size: 1.2rem;
        color: #1565C0;
        margin-top: 1rem;
        font-weight: bold;
    }
    .normal-value {
        color: #2E7D32;
        font-weight: bold;
    }
    .abnormal-value {
        color: #C62828;
        font-weight: bold;
    }
    .borderline-value {
        color: #FF8F00;
        font-weight: bold;
    }
    .info-box {
        background-color: #E3F2FD;
        padding: 1rem;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
    }
    .summary-box {
        background-color: #E8F5E9;
        padding: 1rem;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
        border-left: 5px solid #2E7D32;
    }
    .warning-box {
        background-color: #FFEBEE;
        padding: 1rem;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
        border-left: 5px solid #C62828;
    }
    .info-text {
        font-size: 0.9rem;
        color: #616161;
    }
</style>
""", unsafe_allow_html=True)

# Define medical reference ranges - based on Indian standards
# These would be more extensive in a full application
REFERENCE_RANGES = {
    # CBC
    "HAEMOGLOBIN": {"male": (13.0, 16.5), "female": (12.0, 15.5), "unit": "gm%", "low_concern": "Anemia possible", "high_concern": "Polycythemia possible"},
    "RBC": {"male": (4.2, 5.4), "female": (3.6, 5.0), "unit": "mill/c.mm.", "low_concern": "Low RBC count", "high_concern": "Elevated RBC count"},
    "PCV": {"male": (38, 42), "female": (36, 40), "unit": "%", "low_concern": "Low hematocrit", "high_concern": "Elevated hematocrit"},
    "MCV": {"male": (78, 100), "female": (78, 100), "unit": "fl", "low_concern": "Microcytic anemia possible", "high_concern": "Macrocytic anemia possible"},
    "MCH": {"male": (27, 31), "female": (27, 31), "unit": "pg", "low_concern": "Hypochromic anemia possible", "high_concern": "Hyperchromic anemia possible"},
    "MCHC": {"male": (32, 36), "female": (32, 36), "unit": "%", "low_concern": "Decreased hemoglobin concentration", "high_concern": "Increased hemoglobin concentration"},
    "WBC": {"male": (4500, 10000), "female": (4500, 10000), "unit": "/c.mm.", "low_concern": "Leukopenia - impaired immune response", "high_concern": "Leukocytosis - infection or inflammation possible"},
    "NEUTROPHILS": {"male": (40, 75), "female": (40, 75), "unit": "%", "low_concern": "Neutropenia", "high_concern": "Neutrophilia - bacterial infection possible"},
    "EOSINOPHILS": {"male": (0, 6), "female": (0, 6), "unit": "%", "low_concern": "", "high_concern": "Eosinophilia - allergy or parasitic infection possible"},
    "BASOPHILS": {"male": (0, 1), "female": (0, 1), "unit": "%", "low_concern": "", "high_concern": "Basophilia - inflammatory or allergic reaction possible"},
    "LYMPHOCYTES": {"male": (20, 45), "female": (20, 45), "unit": "%", "low_concern": "Lymphopenia", "high_concern": "Lymphocytosis - viral infection possible"},
    "MONOCYTES": {"male": (2, 10), "female": (2, 10), "unit": "%", "low_concern": "", "high_concern": "Monocytosis - chronic inflammation possible"},
    "PLATELETS": {"male": (140000, 450000), "female": (140000, 450000), "unit": "/c.mm.", "low_concern": "Thrombocytopenia - bleeding risk", "high_concern": "Thrombocytosis - clotting risk"},

    # Lipid Profile
    "CHOLESTEROL_TOTAL": {"male": (70, 200), "female": (70, 200), "unit": "mg/dl", "low_concern": "", "high_concern": "Hypercholesterolemia - increased cardiovascular risk"},
    "TRIGLYCERIDES": {"male": (40, 150), "female": (40, 150), "unit": "mg/dl", "low_concern": "", "high_concern": "Hypertriglyceridemia - increased cardiovascular risk"},
    "HDL": {"male": (35, 120), "female": (35, 120), "unit": "mg/dl", "low_concern": "Low HDL - increased cardiovascular risk", "high_concern": ""},
    "LDL": {
        "normal": (0, 100),
        "borderline": (101, 130),
        "high": (131, 170),
        "very_high": (171, 999),
        "unit": "mg/dl",
        "low_concern": "",
        "high_concern": "Elevated LDL - increased cardiovascular risk"
    },
    "VLDL": {"male": (5, 35), "female": (5, 35), "unit": "mg/dl", "low_concern": "", "high_concern": "Elevated VLDL - increased cardiovascular risk"},
    "CHO_HDL_RATIO": {"male": (3, 5), "female": (3, 5), "unit": "", "low_concern": "", "high_concern": "Elevated ratio - increased cardiovascular risk"},
    "LDL_HDL_RATIO": {"male": (2.5, 3.5), "female": (2.5, 3.5), "unit": "", "low_concern": "", "high_concern": "Elevated ratio - increased cardiovascular risk"},

    # Renal Function
    "CREATININE": {"male": (0.7, 1.3), "female": (0.6, 1.1), "unit": "mg/dl", "low_concern": "", "high_concern": "Elevated creatinine - possible kidney dysfunction"},

    # Liver Function
    "ALKALINE_PHOSPHATE": {"male": (15, 112), "female": (15, 112), "unit": "IU/L", "low_concern": "", "high_concern": "Elevated ALP - possible liver or bone disorder"},
    "SGPT": {"male": (0, 45), "female": (0, 45), "unit": "IU/L", "low_concern": "", "high_concern": "Elevated SGPT - possible liver damage"},

    # Thyroid Function
    "TSH": {"male": (0.39, 6.11), "female": (0.39, 6.11), "unit": "uIU/ml", "low_concern": "Low TSH - possible hyperthyroidism", "high_concern": "Elevated TSH - possible hypothyroidism"},

    # Cancer Screening
    "PSA": {"male": (0, 4.0), "female": None, "unit": "ng/ml", "low_concern": "", "high_concern": "Elevated PSA - prostate abnormality possible, including cancer"},

    # Vitamins
    "VITAMIN_B12": {
        "deficient": (0, 200),
        "normal": (200, 900),
        "excess": (900, 9999),
        "unit": "pg/ml",
        "low_concern": "B12 deficiency - neurological issues possible",
        "high_concern": ""
    },
    "VITAMIN_D3": {
        "deficient": (0, 20),
        "insufficient": (20, 30),
        "normal": (30, 80),
        "excess": (80, 9999),
        "unit": "ng/mL",
        "low_concern": "Vitamin D deficiency - bone health risk",
        "high_concern": "Vitamin D excess - hypercalcemia risk"
    },

    # Diabetes
    "HBA1C": {
        "normal": (0, 5.7),
        "prediabetes": (5.7, 6.5),
        "diabetes": (6.5, 20),
        "unit": "%",
        "low_concern": "",
        "high_concern": "Elevated HbA1c - diabetes or prediabetes"
    }
}

# Patterns for data extraction
PATTERNS = {
    "name": r"Name\s*:\s*([^\n]+)",
    "age": r"Age\s*:\s*(\d+)\s*Years",
    "sex": r"Sex\s*:\s*([A-Za-z]+)",
    "date": r"Date\s*:\s*(\d{1,2}-[A-Za-z]{3}-\d{4})",
    "lab_no": r"Lab No\.\s*:\s*([^\n]+)",
}

# Test patterns for extraction - these would be more extensive in a full application
TEST_PATTERNS = {
    # CBC
    "HAEMOGLOBIN": r"HAEMOGLOBIN\s+([\d\.]+)\s+gm%",
    "RBC": r"R\.B\.C\.\s+MILLIONS\s*/\s*CMM\s+([\d\.]+)",
    "PCV": r"P\.C\.V\s+%\s+([\d\.]+)",
    "MCV": r"M\.C\.V\s+FL\s+([\d\.]+)",
    "MCH": r"M\.C\.H\s+PG\s+([\d\.]+)",
    "MCHC": r"M\.\s*C\.\s*H\.\s*C\.\s+%\s+([\d\.]+)",
    "WBC": r"TOTAL\s+W\.B\.C\.\s+COUNT\s*/\s*CMM\s+(\d+)",
    "NEUTROPHILS": r"NEUTROPHILS\s+%\s+(\d+)",
    "EOSINOPHILS": r"EOSINOPHILS\s+%\s+(\d+)",
    "BASOPHILS": r"BASOPHILS\s+%\s+(\d+)",
    "LYMPHOCYTES": r"LYMPHOCYTES\s+%\s+(\d+)",
    "MONOCYTES": r"MONOCYTES\s+%\s+(\d+)",
    "PLATELETS": r"PLATELETS\s+COUNT\s+(\d+)",

    # Lipid Profile
    "CHOLESTEROL_TOTAL": r"CHOLESTEROL\s+TOTAL\s+([\d\.]+)",
    "TRIGLYCERIDES": r"TRIGLYCERIDES\s+([\d\.]+)",
    "HDL": r"CHOLESTEROL\s+-\s+HDL\s+([\d\.]+)",
    "LDL": r"CHOLESTEROL\s+-\s+LDL\s+([\d\.]+)",
    "VLDL": r"CHOLESTEROL\s+-\s+VLDL\s+([\d\.]+)",
    "CHO_HDL_RATIO": r"CHO\s*/\s*HDL\s+RATIO\s+([\d\.]+)",
    "LDL_HDL_RATIO": r"LDL\s*/\s*HDL\s+RATIO\s+([\d\.]+)",

    # Renal Function
    "CREATININE": r"S\.\s*CREATININE\s+([\d\.]+)",

    # Liver Function
    "ALKALINE_PHOSPHATE": r"ALKALINE\s+PHOSPHATE\s+(\d+)",
    "SGPT": r"S\.G\.P\.T\s+(\d+)",

    # Thyroid Function
    "TSH": r"TSH\s+([\d\.]+)",

    # Cancer Screening
    "PSA": r"PSA\s+-\s+Prostate\s+specific\s+Antigen\s+([\d\.]+)",

    # Vitamins
    "VITAMIN_B12": r"Vitamin\s+B12\s+([\d\.]+)",
    "VITAMIN_D3": r"Vitamin\s+D3\s+([\d\.]+)",

    # Diabetes
    "HBA1C": r"GLYCOSYLATED\s+HBA1c\s+%\s+([\d\.]+)",

    # ECG
    "ECG_RATE": r"Rate\s*:\s*(\d+)/min",
    "ECG_RHYTHM": r"Rhythm\s*:\s*([^\n]+)",
    "ECG_MECHANISM": r"Mechanism\s*:\s*([^\n]+)",
    "ECG_PR_INTERVAL": r"PR\s+Interval\s*:\s*([\d\.]+)",
    "ECG_QT_INTERVAL": r"QT\s+Interval\s*:\s*([\d\.]+)",
    "ECG_AXIS": r"Axis\s*:\s*([^\n]+)",
    "ECG_P_WAVE": r"P\s+Wave\s*:\s*([^\n]+)",
    "ECG_T_WAVE": r"T\s+Wave\s*:\s*([^\n]+)",
    "ECG_QRS_COMPLEX": r"QRS\s+Complex\s*:\s*([^\n]+)",
    "ECG_ST_SEGMENT": r"ST\s+Segment\s*:\s*([^\n]+)",
    "ECG_OTHER_FINDINGS": r"Other\s+Findings\s*:\s*([^\n]+)",
    "ECG_IMPRESSION": r"IMPRESSION\s*:\s*([^\n]+)"
}


def extract_text_from_pdf(pdf_file):
    """Extract text from a PDF file using pdfplumber"""
    text = ""
    try:
        # Create a BytesIO object for pdfplumber to use
        pdf_bytes = io.BytesIO(pdf_file.read())

        # Reset file pointer
        pdf_file.seek(0)

        with pdfplumber.open(pdf_bytes) as pdf:
            for page in pdf.pages:
                text += page.extract_text() + "\n"
        return text
    except Exception as e:
        st.error(f"Error extracting text from PDF: {e}")
        return None


def extract_patient_info(text):
    """Extract patient information from the text"""
    patient_info = {}
    for key, pattern in PATTERNS.items():
        match = re.search(pattern, text)
        if match:
            patient_info[key] = match.group(1).strip()
    return patient_info


def extract_test_results(text):
    """Extract test results from the text"""
    test_results = {}
    for test_name, pattern in TEST_PATTERNS.items():
        match = re.search(pattern, text)
        if match:
            try:
                value = float(match.group(1))
                test_results[test_name] = value
            except ValueError:
                # For non-numeric values like ECG findings
                test_results[test_name] = match.group(1).strip()
    return test_results


def categorize_results(test_results, patient_info):
    """Categorize test results as normal, abnormal, or borderline"""
    categorized_results = {}

    # Default to male if sex not specified
    sex = patient_info.get("sex", "MALE").lower()

    for test, value in test_results.items():
        if test in REFERENCE_RANGES:
            ref_range = REFERENCE_RANGES[test]

            # Handle special cases like LDL, HBA1C, Vitamin D3
            if test == "LDL":
                if value <= ref_range["normal"][1]:
                    categorized_results[test] = {
                        "value": value, "status": "normal", "concern": ""}
                elif value <= ref_range["borderline"][1]:
                    categorized_results[test] = {
                        "value": value, "status": "borderline", "concern": ref_range["high_concern"]}
                elif value <= ref_range["high"][1]:
                    categorized_results[test] = {
                        "value": value, "status": "abnormal_high", "concern": ref_range["high_concern"]}
                else:
                    categorized_results[test] = {
                        "value": value, "status": "very_abnormal_high", "concern": ref_range["high_concern"]}

            elif test == "HBA1C":
                if value < ref_range["normal"][1]:
                    categorized_results[test] = {
                        "value": value, "status": "normal", "concern": ""}
                elif value < ref_range["prediabetes"][1]:
                    categorized_results[test] = {
                        "value": value, "status": "borderline", "concern": "Prediabetes"}
                else:
                    categorized_results[test] = {
                        "value": value, "status": "abnormal_high", "concern": "Diabetes mellitus"}

            elif test == "VITAMIN_D3":
                if value < ref_range["deficient"][1]:
                    categorized_results[test] = {
                        "value": value, "status": "abnormal_low", "concern": "Severe Vitamin D deficiency"}
                elif value < ref_range["insufficient"][1]:
                    categorized_results[test] = {
                        "value": value, "status": "borderline_low", "concern": "Vitamin D insufficiency"}
                elif value <= ref_range["normal"][1]:
                    categorized_results[test] = {
                        "value": value, "status": "normal", "concern": ""}
                else:
                    categorized_results[test] = {
                        "value": value, "status": "abnormal_high", "concern": ref_range["high_concern"]}

            elif test == "VITAMIN_B12":
                if value < ref_range["deficient"][1]:
                    categorized_results[test] = {
                        "value": value, "status": "abnormal_low", "concern": ref_range["low_concern"]}
                elif value <= ref_range["normal"][1]:
                    categorized_results[test] = {
                        "value": value, "status": "normal", "concern": ""}
                else:
                    categorized_results[test] = {
                        "value": value, "status": "abnormal_high", "concern": "Elevated B12 levels"}

            # Regular numeric tests
            elif isinstance(value, (int, float)) and sex in ref_range:
                lower, upper = ref_range[sex]

                if value < lower:
                    categorized_results[test] = {
                        "value": value, "status": "abnormal_low", "concern": ref_range["low_concern"]}
                elif value > upper:
                    categorized_results[test] = {
                        "value": value, "status": "abnormal_high", "concern": ref_range["high_concern"]}
                else:
                    categorized_results[test] = {
                        "value": value, "status": "normal", "concern": ""}

            # Non-numeric values like ECG findings
            else:
                categorized_results[test] = {
                    "value": value, "status": "not_numeric", "concern": ""}
        else:
            # For tests without reference ranges
            categorized_results[test] = {
                "value": value, "status": "no_reference", "concern": ""}

    return categorized_results


def create_summary(categorized_results, patient_info):
    """Create a summary of the test results"""
    abnormal_results = {k: v for k, v in categorized_results.items(
    ) if 'abnormal' in v['status'] or 'borderline' in v['status']}

    summary = {
        "patient_info": patient_info,
        "abnormal_count": len(abnormal_results),
        "abnormal_tests": abnormal_results,
        "risk_factors": [],
        "recommendations": []
    }

    # Check for anemia
    if "HAEMOGLOBIN" in categorized_results and categorized_results["HAEMOGLOBIN"]["status"] == "abnormal_low":
        if "MCV" in categorized_results:
            if categorized_results["MCV"]["status"] == "abnormal_low":
                summary["risk_factors"].append(
                    "Microcytic anemia possible - consider iron deficiency")
                summary["recommendations"].append(
                    "Further investigation for iron deficiency anemia recommended")
            elif categorized_results["MCV"]["status"] == "abnormal_high":
                summary["risk_factors"].append(
                    "Macrocytic anemia possible - consider B12/folate deficiency")
                summary["recommendations"].append(
                    "Further investigation for vitamin B12 or folate deficiency recommended")
            else:
                summary["risk_factors"].append("Normocytic anemia possible")

    # Check for infection/inflammation
    if "WBC" in categorized_results and categorized_results["WBC"]["status"] == "abnormal_high":
        summary["risk_factors"].append(
            "Elevated white blood cell count - possible infection or inflammation")
        summary["recommendations"].append(
            "Monitor for signs of infection or inflammatory conditions")

    # Check cardiovascular risk
    has_lipid_concerns = False
    if "CHOLESTEROL_TOTAL" in categorized_results and categorized_results["CHOLESTEROL_TOTAL"]["status"] == "abnormal_high":
        has_lipid_concerns = True
    if "LDL" in categorized_results and (categorized_results["LDL"]["status"] == "abnormal_high" or categorized_results["LDL"]["status"] == "very_abnormal_high"):
        has_lipid_concerns = True
    if "HDL" in categorized_results and categorized_results["HDL"]["status"] == "abnormal_low":
        has_lipid_concerns = True
    if "TRIGLYCERIDES" in categorized_results and categorized_results["TRIGLYCERIDES"]["status"] == "abnormal_high":
        has_lipid_concerns = True

    if has_lipid_concerns:
        summary["risk_factors"].append(
            "Dyslipidemia - increased cardiovascular risk")
        summary["recommendations"].append(
            "Lifestyle modifications recommended; consider dietary changes and exercise")
        if "LDL" in categorized_results and categorized_results["LDL"]["status"] == "very_abnormal_high":
            summary["recommendations"].append(
                "Consider consultation with cardiologist for lipid management")

    # Check diabetes risk
    if "HBA1C" in categorized_results:
        if categorized_results["HBA1C"]["status"] == "borderline":
            summary["risk_factors"].append("Prediabetes")
            summary["recommendations"].append(
                "Lifestyle modifications recommended; consider dietary changes and exercise")
            summary["recommendations"].append(
                "Follow-up HbA1c test in 3-6 months")
        elif categorized_results["HBA1C"]["status"] == "abnormal_high":
            summary["risk_factors"].append("Diabetes mellitus")
            summary["recommendations"].append(
                "Consultation with endocrinologist recommended")
            summary["recommendations"].append(
                "Regular blood glucose monitoring")

    # Check prostate risk (for males)
    if patient_info.get("sex", "").upper() == "MALE":
        if "PSA" in categorized_results and categorized_results["PSA"]["status"] == "abnormal_high":
            summary["risk_factors"].append(
                "Elevated PSA - prostate abnormality possible")
            summary["recommendations"].append(
                "Urologist consultation recommended")

    # Check vitamin deficiencies
    if "VITAMIN_D3" in categorized_results:
        if categorized_results["VITAMIN_D3"]["status"] == "abnormal_low":
            summary["risk_factors"].append("Severe Vitamin D deficiency")
            summary["recommendations"].append(
                "Vitamin D supplementation recommended")
        elif categorized_results["VITAMIN_D3"]["status"] == "borderline_low":
            summary["risk_factors"].append("Vitamin D insufficiency")
            summary["recommendations"].append(
                "Consider Vitamin D supplementation")

    if "VITAMIN_B12" in categorized_results and categorized_results["VITAMIN_B12"]["status"] == "abnormal_low":
        summary["risk_factors"].append("Vitamin B12 deficiency")
        summary["recommendations"].append(
            "Vitamin B12 supplementation recommended")

    # Check kidney function
    if "CREATININE" in categorized_results and categorized_results["CREATININE"]["status"] == "abnormal_high":
        summary["risk_factors"].append(
            "Elevated creatinine - possible kidney dysfunction")
        summary["recommendations"].append(
            "Follow-up kidney function tests recommended")

    # Check liver function
    liver_concerns = False
    if "SGPT" in categorized_results and categorized_results["SGPT"]["status"] == "abnormal_high":
        liver_concerns = True
    if "ALKALINE_PHOSPHATE" in categorized_results and categorized_results["ALKALINE_PHOSPHATE"]["status"] == "abnormal_high":
        liver_concerns = True

    if liver_concerns:
        summary["risk_factors"].append("Possible liver function abnormalities")
        summary["recommendations"].append(
            "Follow-up liver function tests recommended")

    # Check thyroid function
    if "TSH" in categorized_results:
        if categorized_results["TSH"]["status"] == "abnormal_high":
            summary["risk_factors"].append(
                "Elevated TSH - possible hypothyroidism")
            summary["recommendations"].append(
                "Consider thyroid profile (T3, T4)")
        elif categorized_results["TSH"]["status"] == "abnormal_low":
            summary["risk_factors"].append(
                "Low TSH - possible hyperthyroidism")
            summary["recommendations"].append(
                "Consider thyroid profile (T3, T4)")

    # If no risk factors identified
    if not summary["risk_factors"]:
        summary["risk_factors"].append(
            "No significant risk factors identified")
        summary["recommendations"].append(
            "Routine follow-up as per age and gender appropriate guidelines")

    return summary


def create_excel_download_link(df, patient_name):
    """Generate a download link for the Excel file"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Test Results')

    b64 = base64.b64encode(output.getvalue()).decode()
    filename = f"{patient_name}_medical_report_{datetime.now().strftime('%Y%m%d')}.xlsx"
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Excel Report</a>'
    return href


def prepare_categorized_results(categorized_results, patient_info):
    """Prepare categorized test results with reference ranges for display"""
    # Group tests by category for display
    categories = {
        "Complete Blood Count": ["HAEMOGLOBIN", "RBC", "PCV", "MCV", "MCH", "MCHC", "WBC",
                                 "NEUTROPHILS", "EOSINOPHILS", "BASOPHILS", "LYMPHOCYTES",
                                 "MONOCYTES", "PLATELETS"],
        "Lipid Profile": ["CHOLESTEROL_TOTAL", "TRIGLYCERIDES", "HDL", "LDL", "VLDL",
                          "CHO_HDL_RATIO", "LDL_HDL_RATIO"],
        "Liver Function": ["ALKALINE_PHOSPHATE", "SGPT"],
        "Kidney Function": ["CREATININE"],
        "Thyroid Function": ["TSH"],
        "Cancer Markers": ["PSA"],
        "Vitamins": ["VITAMIN_B12", "VITAMIN_D3"],
        "Diabetes Markers": ["HBA1C"],
        "ECG Findings": ["ECG_RATE", "ECG_RHYTHM", "ECG_MECHANISM", "ECG_PR_INTERVAL",
                         "ECG_QT_INTERVAL", "ECG_AXIS", "ECG_P_WAVE", "ECG_T_WAVE",
                         "ECG_QRS_COMPLEX", "ECG_ST_SEGMENT", "ECG_OTHER_FINDINGS",
                         "ECG_IMPRESSION"]
    }

    sex = patient_info.get("sex", "MALE").lower()

    # Create DataFrames for each category
    dfs = {}
    all_results = []

    for category, tests in categories.items():
        rows = []

        for test in tests:
            if test in categorized_results:
                result = categorized_results[test]
                value = result["value"]
                status = result["status"]

                if test in REFERENCE_RANGES:
                    ref_range = REFERENCE_RANGES[test]

                    # Handle special cases
                    if test == "LDL":
                        ref_text = f"{ref_range['normal'][0]}-{ref_range['normal'][1]} (Normal), {ref_range['borderline'][0]}-{ref_range['borderline'][1]} (Borderline), {ref_range['high'][0]}-{ref_range['high'][1]} (High), >{ref_range['high'][1]} (Very High)"
                    elif test == "HBA1C":
                        ref_text = f"<{ref_range['normal'][1]} (Normal), {ref_range['prediabetes'][0]}-{ref_range['prediabetes'][1]} (Prediabetes), ≥{ref_range['diabetes'][0]} (Diabetes)"
                    elif test == "VITAMIN_D3":
                        ref_text = f"<{ref_range['deficient'][1]} (Deficient), {ref_range['insufficient'][0]}-{ref_range['insufficient'][1]} (Insufficient), {ref_range['normal'][0]}-{ref_range['normal'][1]} (Normal)"
                    elif test == "VITAMIN_B12":
                        ref_text = f"<{ref_range['deficient'][1]} (Deficient), {ref_range['normal'][0]}-{ref_range['normal'][1]} (Normal), >{ref_range['normal'][1]} (Excess)"
                    elif sex in ref_range:
                        ref_text = f"{ref_range[sex][0]}-{ref_range[sex][1]} {ref_range['unit']}"
                    else:
                        ref_text = "Not applicable"
                else:
                    ref_text = "Not available"

                # Format test name for display
                display_name = test.replace("_", " ").title()

                # Determine status for display
                if status == "normal":
                    status_display = "Normal"
                    status_class = "normal-value"
                elif status == "borderline" or status == "borderline_low":
                    status_display = "Borderline"
                    status_class = "borderline-value"
                elif "abnormal" in status:
                    status_display = "Abnormal"
                    status_class = "abnormal-value"
                else:
                    status_display = "Unknown"
                    status_class = ""

                # Add to rows for this category
                rows.append({
                    "Test": display_name,
                    "Value": value,
                    "Reference Range": ref_text,
                    "Status": status_display,
                    "Status Class": status_class,
                    "Concern": result.get("concern", "")
                })

                # Add to all results for Excel export
                all_results.append({
                    "Category": category,
                    "Test": display_name,
                    "Value": value,
                    "Reference Range": ref_text,
                    "Status": status_display,
                    "Concern": result.get("concern", "")
                })

        if rows:  # Only create DataFrame if we have results for this category
            dfs[category] = pd.DataFrame(rows)

    # Create DataFrame with all results for Excel export
    all_results_df = pd.DataFrame(all_results)

    return dfs, all_results_df


def create_visualizations(categorized_results, patient_info):
    """Create visualizations for test results"""
    visualizations = {}

    # Hematology Graph - CBC
    hematology_params = ["HAEMOGLOBIN", "RBC", "WBC", "PLATELETS"]
    hematology_data = []

    for param in hematology_params:
        if param in categorized_results:
            result = categorized_results[param]
            sex = patient_info.get("sex", "MALE").lower()

            if param in REFERENCE_RANGES and sex in REFERENCE_RANGES[param]:
                min_val, max_val = REFERENCE_RANGES[param][sex]
                value = result["value"]

                # Normalize values to 0-100% of reference range
                if param == "PLATELETS":  # Special handling for large values
                    norm_value = ((value - min_val) /
                                  (max_val - min_val)) * 100
                    norm_value = max(0, min(100, norm_value)
                                     )  # Clamp to 0-100%
                    display_value = value / 1000  # Convert to thousands for display
                else:
                    norm_value = ((value - min_val) /
                                  (max_val - min_val)) * 100
                    norm_value = max(0, min(100, norm_value)
                                     )  # Clamp to 0-100%
                    display_value = value

                hematology_data.append({
                    "Parameter": param.replace("_", " ").title(),
                    "Value": display_value,
                    "Normalized": norm_value,
                    "Min": min_val,
                    "Max": max_val,
                    "Status": result["status"]
                })

    if hematology_data:
        visualizations["hematology"] = hematology_data

    # Lipid Profile Graph
    lipid_params = ["CHOLESTEROL_TOTAL", "TRIGLYCERIDES", "HDL", "LDL"]
    lipid_data = []

    for param in lipid_params:
        if param in categorized_results:
            result = categorized_results[param]
            value = result["value"]

            if param == "LDL":  # Special handling for LDL
                min_val, max_val = REFERENCE_RANGES[param]["normal"]
                status = result["status"]
            elif param == "HDL":  # For HDL, higher is better
                sex = patient_info.get("sex", "MALE").lower()
                min_val, max_val = REFERENCE_RANGES[param][sex]
                status = result["status"]
            else:
                sex = patient_info.get("sex", "MALE").lower()
                if param in REFERENCE_RANGES and sex in REFERENCE_RANGES[param]:
                    min_val, max_val = REFERENCE_RANGES[param][sex]
                    status = result["status"]
                else:
                    continue

            lipid_data.append({
                "Parameter": param.replace("_", " ").title(),
                "Value": value,
                "Min": min_val,
                "Max": max_val,
                "Status": status
            })

    if lipid_data:
        visualizations["lipid"] = lipid_data

    return visualizations


def main():
    """Main function for the Streamlit application"""
    st.markdown('<h1 class="main-header">Medical Report Analyzer</h1>',
                unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.image(
            "https://img.icons8.com/color/96/000000/medical-doctor.png", width=100)
        st.markdown("## Upload Medical Report")
        uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

        # Patient details section
        st.markdown("## Patient Details")
        sex_options = ["Male", "Female"]
        default_sex = "Male"
        patient_sex = st.selectbox(
            "Sex", sex_options, index=sex_options.index(default_sex))
        patient_age = st.number_input(
            "Age", min_value=0, max_value=120, value=40)

        # Analysis options
        st.markdown("## Analysis Options")
        show_visuals = st.checkbox("Show Visualizations", value=True)
        show_recommendations = st.checkbox("Show Recommendations", value=True)
        show_detailed_analysis = st.checkbox(
            "Show Detailed Analysis", value=True)

    if uploaded_file is not None:
        # Extract text from PDF
        with st.spinner("Extracting text from PDF..."):
            extracted_text = extract_text_from_pdf(uploaded_file)

        if extracted_text:
            # Extract patient information
            patient_info = extract_patient_info(extracted_text)

            # Use user input if available
            if patient_info.get("sex") is None and patient_sex:
                patient_info["sex"] = patient_sex

            if patient_info.get("age") is None and patient_age:
                patient_info["age"] = str(patient_age)

            # Extract test results
            test_results = extract_test_results(extracted_text)

            if test_results:
                # Categorize results
                categorized_results = categorize_results(
                    test_results, patient_info)

                # Create summary
                summary = create_summary(categorized_results, patient_info)

                # Prepare results for display
                result_dfs, all_results_df = prepare_categorized_results(
                    categorized_results, patient_info)

                # Create Excel download link
                excel_link = create_excel_download_link(
                    all_results_df, patient_info.get("name", "patient"))

                # Create visualizations
                if show_visuals:
                    visualizations = create_visualizations(
                        categorized_results, patient_info)

                # Display patient information
                st.markdown(
                    '<h2 class="sub-header">Patient Information</h2>', unsafe_allow_html=True)
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.markdown(
                        f"**Name:** {patient_info.get('name', 'Not specified')}")

                with col2:
                    st.markdown(
                        f"**Age:** {patient_info.get('age', 'Not specified')} years")

                with col3:
                    st.markdown(
                        f"**Sex:** {patient_info.get('sex', 'Not specified')}")

                st.markdown(
                    f"**Date:** {patient_info.get('date', 'Not specified')}")
                st.markdown(
                    f"**Lab No.:** {patient_info.get('lab_no', 'Not specified')}")

                # Display summary
                st.markdown('<h2 class="sub-header">Summary</h2>',
                            unsafe_allow_html=True)

                if summary["abnormal_count"] == 0:
                    st.markdown(
                        '<div class="summary-box">All test results are within normal ranges. No significant abnormalities detected.</div>', unsafe_allow_html=True)
                else:
                    st.markdown(
                        f'<div class="warning-box">{summary["abnormal_count"]} abnormal test results detected.</div>', unsafe_allow_html=True)

                # Display risk factors
                if summary["risk_factors"]:
                    st.markdown(
                        '<h3 class="section-header">Potential Risk Factors</h3>', unsafe_allow_html=True)
                    for risk in summary["risk_factors"]:
                        st.markdown(f"- {risk}")

                # Display recommendations
                if show_recommendations and summary["recommendations"]:
                    st.markdown(
                        '<h3 class="section-header">Recommendations</h3>', unsafe_allow_html=True)
                    for rec in summary["recommendations"]:
                        st.markdown(f"- {rec}")

                # Display visualizations
                if show_visuals and visualizations:
                    st.markdown(
                        '<h2 class="sub-header">Visualizations</h2>', unsafe_allow_html=True)

                    col1, col2 = st.columns(2)

                    with col1:
                        if "hematology" in visualizations and visualizations["hematology"]:
                            st.markdown(
                                '<h3 class="section-header">Complete Blood Count</h3>', unsafe_allow_html=True)
                            hema_df = pd.DataFrame(
                                visualizations["hematology"])

                            # Create a horizontal bar chart for CBC
                            fig, ax = plt.subplots(figsize=(8, 3))

                            # Set colors based on status
                            colors = []
                            for status in hema_df["Status"]:
                                if status == "normal":
                                    colors.append("#2E7D32")  # Green
                                elif "borderline" in status:
                                    colors.append("#FF8F00")  # Orange
                                else:
                                    colors.append("#C62828")  # Red

                            # Plot normalized values as horizontal bars
                            bars = ax.barh(
                                hema_df["Parameter"], hema_df["Normalized"], color=colors)

                            # Add value labels
                            for i, bar in enumerate(bars):
                                ax.text(bar.get_width() + 5, bar.get_y() + bar.get_height()/2,
                                        f"{hema_df['Value'].iloc[i]}", va='center')

                            # Add a vertical line at 0% and 100%
                            ax.axvline(x=0, color='gray',
                                       linestyle='-', alpha=0.3)
                            ax.axvline(x=100, color='gray',
                                       linestyle='-', alpha=0.3)

                            # Set labels and title
                            ax.set_xlabel('% of Reference Range')
                            # Give some space for the labels
                            ax.set_xlim(0, 120)

                            # Remove top and right spines
                            ax.spines['top'].set_visible(False)
                            ax.spines['right'].set_visible(False)

                            st.pyplot(fig)

                    with col2:
                        if "lipid" in visualizations and visualizations["lipid"]:
                            st.markdown(
                                '<h3 class="section-header">Lipid Profile</h3>', unsafe_allow_html=True)
                            lipid_df = pd.DataFrame(visualizations["lipid"])

                            # Create bar chart for lipid profile
                            fig, ax = plt.subplots(figsize=(8, 4))

                            # Set colors based on status
                            colors = []
                            for status in lipid_df["Status"]:
                                if status == "normal":
                                    colors.append("#2E7D32")  # Green
                                elif "borderline" in status:
                                    colors.append("#FF8F00")  # Orange
                                else:
                                    colors.append("#C62828")  # Red

                            # Plot bars with actual values
                            bars = ax.bar(
                                lipid_df["Parameter"], lipid_df["Value"], color=colors)

                            # Add horizontal lines for max reference values
                            for i, param in enumerate(lipid_df["Parameter"]):
                                ax.axhline(y=lipid_df["Max"].iloc[i], xmin=i/len(lipid_df) + 0.05,
                                           xmax=(i+1)/len(lipid_df) - 0.05, color='black', linestyle='--', alpha=0.7)

                            # Add value labels
                            for bar in bars:
                                height = bar.get_height()
                                ax.text(bar.get_x() + bar.get_width()/2., height + 5,
                                        f"{height:.1f}", ha='center', va='bottom')

                            # Set labels and title
                            ax.set_ylabel('mg/dl')

                            # Remove top and right spines
                            ax.spines['top'].set_visible(False)
                            ax.spines['right'].set_visible(False)

                            st.pyplot(fig)

                # Display detailed test results
                if show_detailed_analysis:
                    st.markdown(
                        '<h2 class="sub-header">Detailed Test Results</h2>', unsafe_allow_html=True)

                    for category, df in result_dfs.items():
                        with st.expander(f"{category}"):
                            # Create a custom table with colored status cells
                            for i, row in df.iterrows():
                                cols = st.columns([3, 2, 3, 2])
                                cols[0].markdown(f"**{row['Test']}**")
                                cols[1].markdown(
                                    f"{row['Value']} {REFERENCE_RANGES.get(row['Test'].upper().replace(' ', '_'), {}).get('unit', '')}")
                                cols[2].markdown(f"{row['Reference Range']}")

                                # Apply color styling based on status
                                if row['Status Class'] == "normal-value":
                                    cols[3].markdown(
                                        f'<span class="normal-value">{row["Status"]}</span>', unsafe_allow_html=True)
                                elif row['Status Class'] == "borderline-value":
                                    cols[3].markdown(
                                        f'<span class="borderline-value">{row["Status"]}</span>', unsafe_allow_html=True)
                                elif row['Status Class'] == "abnormal-value":
                                    cols[3].markdown(
                                        f'<span class="abnormal-value">{row["Status"]}</span>', unsafe_allow_html=True)
                                else:
                                    cols[3].markdown(f"{row['Status']}")

                                # Add concern/notes if present
                                if row['Concern']:
                                    st.markdown(
                                        f"<span class='info-text'>Note: {row['Concern']}</span>", unsafe_allow_html=True)

                                st.markdown("---")

                # Excel download link
                st.markdown("### Download Report")
                st.markdown(excel_link, unsafe_allow_html=True)

            else:
                st.error("No test results could be extracted from the PDF.")
        else:
            st.error(
                "Failed to extract text from the PDF. Please check if the file is valid.")
    else:
        # Display welcome message and instructions when no file is uploaded
        st.markdown(
            """
            <div class="info-box">
                <h3>Welcome to Medical Report Analyzer!</h3>
                <p>Upload a medical report PDF to get started. The application will analyze the report and provide:</p>
                <ul>
                    <li>Analysis of test results based on Indian medical standards</li>
                    <li>Identification of abnormal values</li>
                    <li>Potential risk factors and recommendations</li>
                    <li>Visual representation of key parameters</li>
                    <li>Detailed breakdown of all test results</li>
                    <li>Excel export of structured data</li>
                </ul>
                <p>All processing happens offline - your data never leaves your computer.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            '<h2 class="sub-header">Supported Test Categories</h2>', unsafe_allow_html=True)

        # Display supported test categories
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("### Hematology")
            st.markdown("- Complete Blood Count (CBC)")
            st.markdown("- Differential Count")
            st.markdown("- Platelets")

        with col2:
            st.markdown("### Biochemistry")
            st.markdown("- Liver Function Tests")
            st.markdown("- Kidney Function Tests")
            st.markdown("- Lipid Profile")
            st.markdown("- Thyroid Function")

        with col3:
            st.markdown("### Other Tests")
            st.markdown("- Cancer Markers (PSA, etc.)")
            st.markdown("- Vitamins (B12, D3)")
            st.markdown("- Diabetes Markers (HbA1c)")
            st.markdown("- ECG Findings")


if __name__ == "__main__":
    main()
