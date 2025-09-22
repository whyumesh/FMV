#!/usr/bin/env python3
"""
FMV Calculator - Production Level Script
Automatically matches DVL doctors with CV survey data and updates FMV Calculator
Author: Production System
Version: 2.4 - Added Account:ID_18 from DVL as first column
"""

import pandas as pd
import os
import sys
import logging
from datetime import datetime
from typing import List, Optional
import traceback

# =============================================================================
# CONFIGURATION & LOGGING SETUP
# =============================================================================

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('fmv_calculator.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# File paths
FOLDER_PATH = r"C:/Users/PAWARUX1/Desktop/FMV"
FMV_FILE = os.path.join(FOLDER_PATH, "FMV_Calculator_Updated.xlsx")
CVDUMP_FILE = os.path.join(FOLDER_PATH, "CVdump.csv")
DVL_FILE = os.path.join(FOLDER_PATH, "DVL.csv")
MISSING_FILE = os.path.join(FOLDER_PATH, "Missing_Doctors.csv")
BACKUP_FILE = os.path.join(
    FOLDER_PATH,
    f"FMV_Calculator_Updated_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
)

# Required columns from CVdump
CVDUMP_COLUMNS = [
    "Start time", "HCP Email", "HCP Name",
    "Clinical Experience: i.e. Time Spent with Patients?",
    "Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...",
    "Geographic influence as a Key Opinion Leader.",
    "Highest Academic Position Held in past 10 years",
    "Educational Qualification", "Additional Educational Level ",
    "Specialty / Super Specialty",
    "Years of experience in the\xa0Specialty / Super Specialty?\n",
    "Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years",
    "Publication experience in the past 10 years",
    "Speaking experience (professional, academic, scientific, or media experience) in the past 10 years."
]

# Column mapping from CVdump to FMV Calculator
COLUMN_MAPPING = {
    "HCP Name": "HCP Name",
    "Years of experience in the\xa0Specialty / Super Specialty?\n": "Years of experience in the Specialty / Super Specialty?_x000D_\n",
    "Clinical Experience: i.e. Time Spent with Patients?": "Clinical Experience: i.e. Time Spent with Patients?",
    "Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...": "Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...",
    "Geographic influence as a Key Opinion Leader.": "Geographic influence as a Key Opinion Leader.",
    "Highest Academic Position Held in past 10 years": "Highest Academic Position Held in past 10 years",
    "Additional Educational Level ": "Additional Educational Level",
    "Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years": "Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years",
    "Publication experience in the past 10 years": "Publication experience in the past 10 years",
    "Speaking experience (professional, academic, scientific, or media experience) in the past 10 years.": "Speaking experience (professional, academic, scientific, or media experience) in the past 10 years.",
    "Specialty / Super Specialty": "Specialty / Super Specialty",
    "Educational Qualification": "Educational Qualification",
    "HCP Email": "HCP Email"
}

# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def safe_read_file(file_path: str, usecols: Optional[List[str]] = None, required: bool = True) -> pd.DataFrame:
    """Safely read CSV/Excel file with multiple encoding attempts"""
    if not os.path.exists(file_path):
        if required:
            raise FileNotFoundError(f"Required file not found: {file_path}")
        else:
            logger.warning(f"Optional file not found: {file_path}")
            return pd.DataFrame()
    
    if file_path.lower().endswith(('.xlsx', '.xls')):
        try:
            return pd.read_excel(file_path, usecols=usecols, dtype=str, engine='openpyxl')
        except Exception:
            return pd.read_excel(file_path, usecols=usecols, dtype=str)
    else:
        encodings = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']
        for encoding in encodings:
            try:
                return pd.read_csv(file_path, dtype=str, usecols=usecols, encoding=encoding)
            except Exception:
                continue
        raise ValueError(f"Could not read {file_path}")

# =============================================================================
# CORE FMV PROCESSING
# =============================================================================

def process_fmv():
    """Process FMV, CVdump, and bring in Account:ID_18 from DVL"""
    try:
        logger.info("Reading FMV, CVdump, and DVL files...")
        fmv_df = safe_read_file(FMV_FILE, required=True)
        cv_df = safe_read_file(CVDUMP_FILE, usecols=CVDUMP_COLUMNS, required=True)
        dvl_df = safe_read_file(DVL_FILE, usecols=["HCP Email", "Account:ID_18"], required=True)

        # Normalize headers
        cv_df = cv_df.rename(columns=COLUMN_MAPPING)

        # Merge FMV with CVdump
        merged = pd.merge(fmv_df, cv_df, on="HCP Email", how="left", suffixes=("", "_cv"))

        # Merge in Account:ID_18 from DVL
        merged = pd.merge(merged, dvl_df, on="HCP Email", how="left")

        # Reorder columns: Account:ID_18 first, then everything else
        cols = merged.columns.tolist()
        if "Account:ID_18" in cols:
            cols = ["Account:ID_18"] + [c for c in cols if c != "Account:ID_18"]
            merged = merged[cols]

        # Save backup
        fmv_df.to_excel(BACKUP_FILE, index=False, engine="openpyxl")
        logger.info(f"Backup created at {BACKUP_FILE}")

        # Overwrite FMV file with updated data
        merged.to_excel(FMV_FILE, index=False, engine="openpyxl")
        logger.info(f"FMV_Calculator_Updated.xlsx updated successfully with Account:ID_18 as first column")

        # Log missing doctors with Account:ID_18 included
        missing = merged[merged["HCP Name"].isna()]
        if not missing.empty:
            missing.to_csv(MISSING_FILE, index=False, encoding="utf-8")
            logger.warning(f"Missing doctors exported to {MISSING_FILE}")

    except Exception as e:
        logger.error(f"Error in FMV processing: {str(e)}")
        traceback.print_exc()

# =============================================================================
# MAIN
# =============================================================================

def main():
    try:
        logger.info("Starting FMV Calculator process...")

        # Process FMV + CVdump + Account:ID_18
        process_fmv()

        logger.info("FMV Calculator process completed successfully")

    except Exception as e:
        logger.error(f"Fatal error: {str(e)}")
        traceback.print_exc()


if __name__ == "__main__":
    main()
