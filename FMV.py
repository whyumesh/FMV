#!/usr/bin/env python3
"""
FMV Calculator - Production Level Script
Automatically matches DVL doctors with CV survey data and updates FMV Calculator
Author: Production System
Version: 2.0
"""

import pandas as pd
import os
import sys
import logging
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import traceback

# =============================================================================
# CONFIGURATION & LOGGING SETUP
# =============================================================================

# Setup logging
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
FOLDER_PATH = r"E:/FMV/FMV"
FMV_FILE = os.path.join(FOLDER_PATH, "FMV_Calculator_Updated.xlsx")  # Use your updated Excel file with formulas
CVDUMP_FILE = os.path.join(FOLDER_PATH, "CVdump.csv")
DVL_FILE = os.path.join(FOLDER_PATH, "DVL.csv")
MISSING_FILE = os.path.join(FOLDER_PATH, "Missing_Doctors.csv")
BACKUP_FILE = os.path.join(FOLDER_PATH, f"FMV_Calculator_Updated_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

# Required columns from CVdump (actual column names from the file)
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
    "Years of experience in the\xa0Specialty / Super Specialty?\n": "Years of experience in the\xa0Specialty / Super Specialty?_x000D_\n",
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
    """
    Safely read file (CSV or Excel) with multiple encoding attempts and error handling
    
    Args:
        file_path: Path to file
        usecols: List of columns to read
        required: Whether file is required (raises error if missing)
    
    Returns:
        DataFrame with loaded data
    """
    if not os.path.exists(file_path):
        if required:
            raise FileNotFoundError(f"Required file not found: {file_path}")
        else:
            logger.warning(f"Optional file not found: {file_path}")
            return pd.DataFrame()
    
    # Check file extension
    if file_path.lower().endswith('.xlsx') or file_path.lower().endswith('.xls'):
        try:
            logger.info(f"Reading Excel file: {file_path}")
            df = pd.read_excel(file_path, usecols=usecols, dtype=str)
            logger.info(f"Successfully read {file_path}")
            return df
        except Exception as e:
            logger.error(f"Error reading Excel file {file_path}: {str(e)}")
            raise
    else:
        # CSV file handling
        encodings = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                logger.info(f"Attempting to read {file_path} with {encoding} encoding")
                df = pd.read_csv(file_path, dtype=str, usecols=usecols, encoding=encoding)
                logger.info(f"Successfully read {file_path} with {encoding} encoding")
                return df
            except (UnicodeDecodeError, UnicodeError) as e:
                logger.warning(f"Failed to read with {encoding}: {str(e)}")
                continue
            except Exception as e:
                logger.error(f"Unexpected error reading {file_path} with {encoding}: {str(e)}")
                raise
        
        raise ValueError(f"Could not read {file_path} with any supported encoding")

def validate_dataframe(df: pd.DataFrame, name: str, required_columns: List[str]) -> bool:
    """
    Validate DataFrame has required columns and is not empty
    
    Args:
        df: DataFrame to validate
        name: Name for logging
        required_columns: List of required column names
    
    Returns:
        True if valid, False otherwise
    """
    if df.empty:
        logger.error(f"{name} is empty")
        return False
    
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        logger.error(f"{name} missing required columns: {missing_cols}")
        return False
    
    logger.info(f"{name} validation passed: {len(df)} rows, {len(df.columns)} columns")
    return True

def clean_email(email: str) -> str:
    """
    Clean and normalize email address
    
    Args:
        email: Raw email string
    
    Returns:
        Cleaned email string
    """
    if pd.isna(email) or email == 'nan' or email == '':
        return ''
    
    email = str(email).strip().lower()
    # Remove any non-printable characters
    email = ''.join(char for char in email if char.isprintable())
    return email

def parse_datetime_safe(date_str: str) -> Optional[pd.Timestamp]:
    """
    Safely parse datetime string with multiple format attempts
    
    Args:
        date_str: Date string to parse
    
    Returns:
        Parsed timestamp or None if parsing fails
    """
    if pd.isna(date_str) or date_str == 'nan' or date_str == '':
        return None
    
    date_str = str(date_str).strip()
    
    # Common datetime formats (prioritizing m/d/yyyy h:mm format as specified)
    formats = [
        '%m/%d/%y %H:%M',      # 10/25/24 17:32
        '%m/%d/%Y %H:%M',      # 10/25/2024 17:32
        '%m-%d-%y %H:%M',      # 10-25-24 17:32
        '%m-%d-%Y %H:%M',      # 10-25-2024 17:32
        '%m/%d/%y %H:%M:%S',   # 10/25/24 17:32:53
        '%m/%d/%Y %H:%M:%S',   # 10/25/2024 17:32:53
        '%m-%d-%y %H:%M:%S',   # 10-25-24 17:32:53
        '%m-%d-%Y %H:%M:%S',   # 10-25-2024 17:32:53
        '%Y-%m-%d %H:%M:%S',   # 2024-10-25 17:32:53
        '%d/%m/%Y %H:%M',      # 25/10/2024 17:32
        '%d-%m-%Y %H:%M'       # 25-10-2024 17:32
    ]
    
    for fmt in formats:
        try:
            return pd.to_datetime(date_str, format=fmt)
        except (ValueError, TypeError):
            continue
    
    # Fallback to pandas auto-detection
    try:
        return pd.to_datetime(date_str, errors='coerce')
    except:
        logger.warning(f"Could not parse datetime: {date_str}")
        return None

def create_backup(file_path: str) -> bool:
    """
    Create backup of file before modification
    
    Args:
        file_path: Path to file to backup
    
    Returns:
        True if backup successful, False otherwise
    """
    try:
        if os.path.exists(file_path):
            import shutil
            shutil.copy2(file_path, BACKUP_FILE)
            logger.info(f"Backup created: {BACKUP_FILE}")
            return True
    except Exception as e:
        logger.error(f"Failed to create backup: {str(e)}")
    return False

# =============================================================================
# MAIN PROCESSING FUNCTIONS
# =============================================================================

def load_data() -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Load all required data files with validation
    
    Returns:
        Tuple of (fmv_df, cvdump_df, dvl_df)
    """
    logger.info("=" * 60)
    logger.info("STARTING FMV CALCULATOR PROCESS")
    logger.info("=" * 60)
    
    # Load FMV Calculator (skip first row as it contains header info)
    logger.info("Loading FMV Calculator...")
    fmv_df = safe_read_file(FMV_FILE, required=True)
    
    # The FMV Calculator has a complex header structure - use row 0 as column names
    # But only if the first row doesn't look like actual data
    if len(fmv_df) > 0 and 'Unnamed:' in str(fmv_df.columns[0]):
        fmv_df.columns = fmv_df.iloc[0]
        fmv_df = fmv_df.drop(fmv_df.index[0]).reset_index(drop=True)
    
    # Load CVdump data
    logger.info("Loading CVdump data...")
    cvdump_df = safe_read_file(CVDUMP_FILE, usecols=CVDUMP_COLUMNS, required=True)
    
    # Load DVL data
    logger.info("Loading DVL data...")
    dvl_df = safe_read_file(DVL_FILE, required=True)
    
    # Validate data
    if not validate_dataframe(fmv_df, "FMV Calculator", ["HCP Email"]):
        raise ValueError("FMV Calculator validation failed")
    
    if not validate_dataframe(cvdump_df, "CVdump", ["HCP Email", "Start time"]):
        raise ValueError("CVdump validation failed")
    
    if not validate_dataframe(dvl_df, "DVL", ["Account: Email", "Customer Code"]):
        raise ValueError("DVL validation failed")
    
    return fmv_df, cvdump_df, dvl_df

def process_cvdump_data(cvdump_df: pd.DataFrame) -> pd.DataFrame:
    """
    Process CVdump data: clean emails, parse dates, handle duplicates
    
    Args:
        cvdump_df: Raw CVdump DataFrame
    
    Returns:
        Processed CVdump DataFrame
    """
    logger.info("Processing CVdump data...")
    
    # Create a copy to avoid modifying original
    df = cvdump_df.copy()
    
    # Clean email addresses
    df["HCP Email"] = df["HCP Email"].apply(clean_email)
    
    # Parse datetime with multiple format attempts
    logger.info("Parsing datetime fields...")
    df["Start time"] = df["Start time"].apply(parse_datetime_safe)
    
    # Remove rows with invalid emails or dates
    initial_count = len(df)
    df = df[df["HCP Email"] != '']
    df = df[df["Start time"].notna()]
    
    removed_count = initial_count - len(df)
    if removed_count > 0:
        logger.warning(f"Removed {removed_count} rows with invalid emails or dates")
    
    # Sort by email and datetime, then keep latest entry per email
    logger.info("Handling duplicate emails (keeping latest entry)...")
    df = df.sort_values(["HCP Email", "Start time"], na_position='last')
    df = df.drop_duplicates("HCP Email", keep="last")
    
    logger.info(f"CVdump processing complete: {len(df)} unique entries")
    return df

def match_doctors(dvl_df: pd.DataFrame, cvdump_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Match DVL doctors with CVdump survey data
    
    Args:
        dvl_df: DVL DataFrame
        cvdump_df: Processed CVdump DataFrame
    
    Returns:
        Tuple of (matched_df, missing_df)
    """
    logger.info("Matching doctors with survey data...")
    
    # Clean DVL emails
    dvl_df = dvl_df.copy()
    dvl_df["Account: Email"] = dvl_df["Account: Email"].apply(clean_email)
    
    # Remove invalid emails from DVL
    dvl_df = dvl_df[dvl_df["Account: Email"] != '']
    
    # Create email lookup from CVdump
    cvdump_lookup = cvdump_df.set_index("HCP Email")
    
    matched_data = []
    missing_data = []
    
    total_dvl = len(dvl_df)
    logger.info(f"Processing {total_dvl} DVL entries...")
    
    for idx, row in dvl_df.iterrows():
        email = row["Account: Email"]
        dvl_code = row["Customer Code"]
        
        if email in cvdump_lookup.index:
            # Found match - get the survey data
            survey_data = cvdump_lookup.loc[email].to_dict()
            
            # Create combined record
            combined_record = {
                "DVL Code": dvl_code,
                "HCP Email": email,
                **{COLUMN_MAPPING.get(k, k): v for k, v in survey_data.items() if k in COLUMN_MAPPING}
            }
            
            matched_data.append(combined_record)
        else:
            # No match found
            missing_data.append({
                "DVL Code": dvl_code,
                "HCP Email": email
            })
    
    matched_df = pd.DataFrame(matched_data)
    missing_df = pd.DataFrame(missing_data)
    
    logger.info(f"Matching complete:")
    logger.info(f"   - Matched: {len(matched_df)} doctors")
    logger.info(f"   - Missing: {len(missing_df)} doctors")
    
    return matched_df, missing_df

def update_fmv_calculator(fmv_df: pd.DataFrame, matched_df: pd.DataFrame) -> pd.DataFrame:
    """
    Update FMV Calculator with new matched data and update existing records with missing data
    
    Args:
        fmv_df: Current FMV Calculator DataFrame
        matched_df: New matched data to add
    
    Returns:
        Updated FMV Calculator DataFrame
    """
    logger.info("Updating FMV Calculator...")
    
    if matched_df.empty:
        logger.info("No new data to add to FMV Calculator")
        return fmv_df
    
    # Get existing emails to avoid duplicates
    existing_emails = set(fmv_df["HCP Email"].dropna().apply(clean_email))
    
    # Filter out already existing emails
    new_emails = matched_df["HCP Email"].apply(clean_email)
    new_data = matched_df[~new_emails.isin(existing_emails)]
    
    # Also check for existing emails that need updating (missing years data)
    existing_data = matched_df[new_emails.isin(existing_emails)]
    
    updated_count = 0
    if not existing_data.empty:
        logger.info(f"Found {len(existing_data)} existing doctors to update with missing data")
        
        # Update existing records with missing data
        for idx, row in existing_data.iterrows():
            email = clean_email(row["HCP Email"])
            fmv_idx = fmv_df[fmv_df["HCP Email"].apply(clean_email) == email].index
            
            if len(fmv_idx) > 0:
                fmv_idx = fmv_idx[0]
                # Update only if the target field is empty/NaN
                years_col = "Years of experience in the\xa0Specialty / Super Specialty?_x000D_\n"
                if pd.isna(fmv_df.loc[fmv_idx, years_col]) and not pd.isna(row.get(years_col)):
                    for col in matched_df.columns:
                        if col in fmv_df.columns and not pd.isna(row[col]):
                            fmv_df.loc[fmv_idx, col] = row[col]
                    updated_count += 1
    
    if new_data.empty and updated_count == 0:
        logger.info("All matched doctors already exist in FMV Calculator with complete data")
        return fmv_df
    
    # Ensure all required columns exist in new data
    for col in fmv_df.columns:
        if col not in new_data.columns:
            new_data[col] = None
    
    # Reorder columns to match FMV Calculator
    new_data = new_data[fmv_df.columns]
    
    # Append new data
    updated_fmv = pd.concat([fmv_df, new_data], ignore_index=True)
    
    logger.info(f"Added {len(new_data)} new doctors to FMV Calculator")
    logger.info(f"Updated {updated_count} existing doctors with missing data")
    return updated_fmv

def save_results(fmv_df: pd.DataFrame, missing_df: pd.DataFrame) -> bool:
    """
    Save all results to files
    
    Args:
        fmv_df: Updated FMV Calculator DataFrame
        missing_df: Missing doctors DataFrame
    
    Returns:
        True if all saves successful, False otherwise
    """
    logger.info("Saving results...")
    
    success = True
    
    try:
        # Create backup before saving
        create_backup(FMV_FILE)
        
        # Save updated FMV Calculator
        if FMV_FILE.lower().endswith('.xlsx'):
            fmv_df.to_excel(FMV_FILE, index=False, engine='openpyxl')
            logger.info(f"FMV Calculator saved as Excel: {len(fmv_df)} total records")
        else:
            fmv_df.to_csv(FMV_FILE, index=False, encoding="utf-8-sig")
            logger.info(f"FMV Calculator saved as CSV: {len(fmv_df)} total records")
        
    except Exception as e:
        logger.error(f"Failed to save FMV Calculator: {str(e)}")
        success = False
    
    try:
        # Save missing doctors
        if not missing_df.empty:
            missing_df.to_csv(MISSING_FILE, index=False, encoding="utf-8-sig")
            logger.info(f"Missing doctors saved: {len(missing_df)} records")
        else:
            # Create empty file if no missing doctors
            pd.DataFrame(columns=["DVL Code", "HCP Email"]).to_csv(MISSING_FILE, index=False, encoding="utf-8-sig")
            logger.info("Missing doctors file created (empty)")
            
    except Exception as e:
        logger.error(f"Failed to save missing doctors: {str(e)}")
        success = False
    
    return success

# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    """
    Main execution function with comprehensive error handling
    """
    try:
        # Load all data
        fmv_df, cvdump_df, dvl_df = load_data()
        
        # Process CVdump data
        processed_cvdump = process_cvdump_data(cvdump_df)
        
        # Match doctors
        matched_df, missing_df = match_doctors(dvl_df, processed_cvdump)
        
        # Update FMV Calculator
        updated_fmv = update_fmv_calculator(fmv_df, matched_df)
        
        # Save results
        if save_results(updated_fmv, missing_df):
            logger.info("=" * 60)
            logger.info("FMV CALCULATOR PROCESS COMPLETED SUCCESSFULLY")
            logger.info("=" * 60)
            logger.info(f"Final Statistics:")
            logger.info(f"   - Total FMV records: {len(updated_fmv)}")
            logger.info(f"   - New doctors added: {len(matched_df)}")
            logger.info(f"   - Missing doctors: {len(missing_df)}")
            logger.info(f"   - Backup created: {BACKUP_FILE}")
        else:
            logger.error("Process completed with errors during save")
            return False
            
    except Exception as e:
        logger.error("=" * 60)
        logger.error("CRITICAL ERROR - PROCESS FAILED")
        logger.error("=" * 60)
        logger.error(f"Error: {str(e)}")
        logger.error(f"Traceback: {traceback.format_exc()}")
        return False
    
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)

