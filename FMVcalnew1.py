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
FOLDER_PATH = r"C:/Users/PAWARUX1/Desktop/FMV"
FMV_FILE = os.path.join(FOLDER_PATH, "FMV_Calculator_Updated.xlsx")  # Use your original file with headers
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
# SCORING FUNCTIONS
# =============================================================================

def load_scoring_criteria():
    """Load scoring criteria from the scoring_criteria.xlsx file"""
    try:
        # Load Details 1 sheet for scoring lookup
        details_df = pd.read_excel("scoring_criteria.xlsx", sheet_name="Details 1", header=None)
        
        # Load OUS FMV Rates sheet for honorarium calculation
        rates_df = pd.read_excel("scoring_criteria.xlsx", sheet_name="OUS FMV Rates", header=1)
        
        logger.info("Scoring criteria loaded successfully")
        return details_df, rates_df
    except Exception as e:
        logger.error(f"Error loading scoring criteria: {str(e)}")
        raise

def create_scoring_lookup(details_df):
    """Create lookup dictionaries from the Details 1 sheet"""
    scoring_lookup = {}
    
    # Manually create the lookup dictionaries based on the known structure
    # Years of experience
    scoring_lookup["years_experience"] = {
        "1-2 years of experience": 0,
        "3-7 years of experience": 2,
        "8-14 years of experience": 4,
        "15+ years of experience": 6
    }
    
    # Clinical experience
    scoring_lookup["clinical_experience"] = {
        "Minimal patient interactions and predominantly administrative/academic work": 0,
        "Less than half the time spent with patients in clinical setting and higher focus on academic/administrative work": 2,
        "Equal amount of time spent with patients in clinical setting and equal amount of time spent in academic/administrative work": 4,
        "Significant time spent with patients in clinical setting and minimal time spent in academic/administrative work": 6
    }
    
    # Leadership position
    scoring_lookup["leadership"] = {
        "Not applicable, as not a part of any society or leadership roles in hospital": 0,
        "1-2 years in a leadership position(s) eg. HOD of a particular speciality in Hospital or other Patient Care Setting and/or serving as a President, Vice president, Secretary,Treasurer, Board member in a Professional or Scientific Society.": 2,
        "3-7 years in a leadership position(s) eg HOD of a particular speciality in Hospital or other Patient Care Setting and/or serving as a national/regional leader in a Professional or Scientific Society.": 4,
        "8 or more years in a leadership position(s) eg HOD for a specialty in Hospital or other Patient Care Setting and/or serving as an international leader in a Professional or Scientific Society.": 6
    }
    
    # Geographical reach
    scoring_lookup["geographical_reach"] = {
        "Local Influence": 0,
        "National Influence": 2,
        "Multi-Country Influence": 4,
        "Global/Worldwide Influence": 6
    }
    
    # Academic position
    scoring_lookup["academic_position"] = {
        "None or N/A": 0,
        "Professor (including Associate / Assistant Professor)": 2,
        "Professor or Adjunct/Additional/Emeritus Professor": 4,
        "Department Chair/ HOD (or similar position)": 6
    }
    
    # Additional education
    scoring_lookup["additional_education"] = {
        "None or N/A": 0,
        "1 Additional degree, fellowship, or advanced training certification.": 2,
        "2 Additional degrees, fellowship, or advanced training certification.": 4,
        "3 or More Additional degrees, fellowship, or advanced training certification.": 6
    }
    
    # Research experience
    scoring_lookup["research_experience"] = {
        "None or N/A": 0,
        "Participation as an Investigator or Sub-Investigator in 1 to 4 clinical trials or research studies.": 2,
        "Participation as an Investigator or Sub-Investigator in 5 to 9 clinical trials or research studies.": 4,
        "Participation as an Investigator of Sub-Investigator in 10 or more clinical trials or research studies or Principal Investigator for two or more clinical trials or research studies or serving as the Principal Investigator for a clinical trial or research study that led to important medical innovations or significant medical technology breakthroughs.": 6
    }
    
    # Publication experience
    scoring_lookup["publication_experience"] = {
        "None or N/A": 0,
        "Co-authorship or participation as contributing author on 1 to 4 publications.": 2,
        "First authorship (if known) on 1 to 5 publications and/or co-authorship or participation as contributing author on 6 to 10 publications": 4,
        "First authorship (if known) on 6 or more publications and/or co-authorship or participation as contributing author on 11 or more publications": 6
    }
    
    # Speaking experience
    scoring_lookup["speaking_experience"] = {
        "Local speaking engagements and the scientific work done for the specialty is near to the practice location": 0,
        "Most of the speaking engagements are directed nationally for the conferences, symposia or national webinars in the designated specialty and the scientific work done is not restricted for the local audience": 2,
        "The speaking experiences are not restricted nationally but to a group of specified countries and the scientific work is directed to the same group of countries": 4,
        "The speaking engagements and the scinetific work carried out is across the globe": 6
    }
    
    return scoring_lookup

def calculate_individual_scores(row, scoring_lookup):
    """Calculate individual scores (Score 1-9) for a doctor"""
    scores = {}
    
    # Score 1: Years of experience
    years_text = str(row.get("Years of experience in the Specialty / Super Specialty?_x000D_\n", "")).strip()
    scores["Score 1"] = scoring_lookup.get("years_experience", {}).get(years_text, 0)
    
    # Score 2: Clinical Experience
    clinical_text = str(row.get("Clinical Experience: i.e. Time Spent with Patients?", "")).strip()
    scores["Score 2"] = scoring_lookup.get("clinical_experience", {}).get(clinical_text, 0)
    
    # Score 3: Leadership position
    leadership_text = str(row.get("Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...", "")).strip()
    scores["Score 3"] = scoring_lookup.get("leadership", {}).get(leadership_text, 0)
    
    # Score 4: Geographical influence
    geo_text = str(row.get("Geographic influence as a Key Opinion Leader.", "")).strip()
    scores["Score 4"] = scoring_lookup.get("geographical_reach", {}).get(geo_text, 0)
    
    # Score 5: Highest Academic Position
    academic_text = str(row.get("Highest Academic Position Held in past 10 years", "")).strip()
    scores["Score 5"] = scoring_lookup.get("academic_position", {}).get(academic_text, 0)
    
    # Score 6: Additional Educational Level
    add_edu_text = str(row.get("Additional Educational Level", "")).strip()
    scores["Score 6"] = scoring_lookup.get("additional_education", {}).get(add_edu_text, 0)
    
    # Score 7: Research Experience
    research_text = str(row.get("Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years", "")).strip()
    scores["Score 7"] = scoring_lookup.get("research_experience", {}).get(research_text, 0)
    
    # Score 8: Publication experience
    pub_text = str(row.get("Publication experience in the past 10 years", "")).strip()
    scores["Score 8"] = scoring_lookup.get("publication_experience", {}).get(pub_text, 0)
    
    # Score 9: Speaking experience
    speaking_text = str(row.get("Speaking experience (professional, academic, scientific, or media experience) in the past 10 years.", "")).strip()
    scores["Score 9"] = scoring_lookup.get("speaking_experience", {}).get(speaking_text, 0)
    
    return scores

def calculate_tier(total_score):
    """Calculate tier based on total score"""
    if total_score <= 13:
        return "Tier 1"
    elif total_score <= 26:
        return "Tier 2"
    elif total_score <= 40:
        return "Tier 3"
    else:
        return "Tier 4"

def calculate_honorarium_rate(specialty, tier, rates_df):
    """Calculate honorarium rate based on specialty and tier using precise OUS FMV Rates"""
    try:
        # Clean specialty name for better matching
        specialty_clean = str(specialty).strip()
        
        # Find matching specialty in rates (exact match first)
        specialty_row = rates_df[rates_df["HCP Specialty"] == specialty_clean]
        
        if specialty_row.empty:
            # Try case-insensitive exact match
            specialty_row = rates_df[rates_df["HCP Specialty"].str.lower() == specialty_clean.lower()]
        
        if specialty_row.empty:
            # Try partial matching for specialties that might have slight variations
            specialty_row = rates_df[rates_df["HCP Specialty"].str.contains(specialty_clean, case=False, na=False)]
        
        if specialty_row.empty:
            # Log the specialty that wasn't found for debugging
            logger.warning(f"Specialty not found in rates table: '{specialty_clean}'")
            # Use a conservative default rate based on tier
            default_rates = {
                "Tier 1": 5000,   # Conservative default
                "Tier 2": 7000,   
                "Tier 3": 9000,   
                "Tier 4": 12000   
            }
            return default_rates.get(tier, 5000)
        
        # Get the rate for the tier - use exact column names with spaces
        if tier in specialty_row.columns:
            rate = specialty_row[tier].iloc[0]
            # Ensure we return a whole number
            return int(rate) if not pd.isna(rate) else 0
        else:
            # Log the tier that wasn't found
            logger.warning(f"Tier column not found: '{tier}' in columns: {list(specialty_row.columns)}")
            return 0
            
    except Exception as e:
        logger.warning(f"Error calculating honorarium rate for '{specialty}', '{tier}': {str(e)}")
        return 0

def apply_scoring_to_fmv(fmv_df, scoring_lookup, rates_df):
    """Apply scoring calculations to the FMV DataFrame"""
    logger.info("Applying scoring calculations...")
    
    # Debug: Check if scoring columns exist
    required_columns = ["Score based on selection mentioned criteria", "Score 1", "Score 2", "Tier", "Rate of Honorarium"]
    missing_columns = [col for col in required_columns if col not in fmv_df.columns]
    if missing_columns:
        logger.warning(f"Missing scoring columns: {missing_columns}")
        # Add missing columns
        for col in missing_columns:
            fmv_df[col] = None
    
    for idx, row in fmv_df.iterrows():
        # Calculate individual scores
        scores = calculate_individual_scores(row, scoring_lookup)
        
        # Calculate total score
        total_score = sum(scores.values())
        
        # Calculate tier
        tier = calculate_tier(total_score)
        
        # Calculate honorarium rate
        specialty = str(row.get("Specialty / Super Specialty", "")).strip()
        honorarium_rate = calculate_honorarium_rate(specialty, tier, rates_df)
        
        # Debug: Log first few calculations
        if idx < 3:
            logger.info(f"Row {idx}: {row['HCP Name']} - Total: {total_score}, Tier: {tier}")
        
        # Update the DataFrame
        fmv_df.loc[idx, "Score based on selection mentioned criteria"] = total_score
        fmv_df.loc[idx, "Score 1"] = scores["Score 1"]
        fmv_df.loc[idx, "Score 2"] = scores["Score 2"]
        fmv_df.loc[idx, "Score 3"] = scores["Score 3"]
        fmv_df.loc[idx, "Score 4"] = scores["Score 4"]
        fmv_df.loc[idx, "Score 5"] = scores["Score 5"]
        fmv_df.loc[idx, "Score 6"] = scores["Score 6"]
        fmv_df.loc[idx, "Score 7"] = scores["Score 7"]
        fmv_df.loc[idx, "Score 8"] = scores["Score 8"]
        fmv_df.loc[idx, "Score 9"] = scores["Score 9"]
        fmv_df.loc[idx, "Range"] = f"{total_score}-{total_score}"  # Single score range
        fmv_df.loc[idx, "Tier"] = tier
        fmv_df.loc[idx, "Rate of Honorarium"] = honorarium_rate
    
    logger.info("Scoring calculations completed")
    
    # Debug: Check final results
    logger.info(f"Final sample - Row 0: {fmv_df.iloc[0]['HCP Name']} - Score: {fmv_df.iloc[0]['Score based on selection mentioned criteria']}")
    
    return fmv_df

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
            # Try reading with data_only=True to ignore formatting issues
            df = pd.read_excel(file_path, usecols=usecols, dtype=str, engine='openpyxl')
            logger.info(f"Successfully read {file_path}")
            return df
        except Exception as e:
            logger.warning(f"Error reading Excel file with openpyxl: {str(e)}")
            try:
                # Try with xlrd engine as fallback
                logger.info("Trying with xlrd engine...")
                df = pd.read_excel(file_path, usecols=usecols, dtype=str, engine='xlrd')
                logger.info(f"Successfully read {file_path} with xlrd")
                return df
            except Exception as e2:
                logger.warning(f"Error reading with xlrd: {str(e2)}")
                try:
                    # Try reading without usecols to avoid column issues
                    logger.info("Trying to read all columns...")
                    df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
                    if usecols:
                        # Filter to only the columns we need
                        available_cols = [col for col in usecols if col in df.columns]
                        df = df[available_cols]
                    logger.info(f"Successfully read {file_path} with all columns")
                    return df
                except Exception as e3:
                    logger.error(f"All methods failed to read Excel file {file_path}: {str(e3)}")
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
    # For FMV Calculator, empty is okay if it has the right columns
    if df.empty and name == "FMV Calculator":
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            logger.error(f"{name} missing required columns: {missing_cols}")
            return False
        logger.info(f"{name} validation passed: empty file with correct columns")
        return True
    
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
    
    # If the file has only headers (no data), create an empty DataFrame with proper columns
    if len(fmv_df) == 0:
        logger.info("FMV Calculator file has only headers - will add new data")
        # Create empty DataFrame with the same column structure
        fmv_df = pd.DataFrame(columns=fmv_df.columns)
    
    # For empty files, we need to ensure we have the required columns
    if 'HCP Email' not in fmv_df.columns:
        # Add the required column if it doesn't exist
        fmv_df['HCP Email'] = None
    
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
    
    if not validate_dataframe(dvl_df, "DVL Code", ["Account: Email", "Customer Code"]):
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
                "DVL Code": dvl_code,  # DVL Code column is named "i" in the FMV Calculator
                "HCP Email": email,
                **{COLUMN_MAPPING.get(k, k): v for k, v in survey_data.items() if k in COLUMN_MAPPING}
            }
            
            matched_data.append(combined_record)
        else:
            # No match found
            missing_data.append({
                "DVL Code": dvl_code,  # DVL Code column is named "i" in the FMV Calculator
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
                # Update DVL code and other missing fields
                years_col = "Years of experience in the Specialty / Super Specialty?_x000D_\n"
                dvl_code_col = "DVL Code"  # DVL Code column
                
                # Always update DVL code if it's missing
                if pd.isna(fmv_df.loc[fmv_idx, dvl_code_col]) and not pd.isna(row.get(dvl_code_col)):
                    fmv_df.loc[fmv_idx, dvl_code_col] = row[dvl_code_col]
                    updated_count += 1
                
                # Update other fields if years column is empty
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
            fmv_df.to_excel(FMV_FILE, index=False, engine='openpyxl', sheet_name='HCP Database')
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
        
        # Load scoring criteria and apply scoring
        logger.info("Loading scoring criteria...")
        details_df, rates_df = load_scoring_criteria()
        scoring_lookup = create_scoring_lookup(details_df)
        
        logger.info("Applying scoring calculations...")
        updated_fmv = apply_scoring_to_fmv(updated_fmv, scoring_lookup, rates_df)
        
        # Debug: Check if scoring worked
        logger.info(f"Sample scoring results: {updated_fmv.iloc[0]['Score based on selection mentioned criteria']}")
        
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
