#!/usr/bin/env python3
"""
FMV Calculator - Production Level Script
Automatically matches DVL doctors with CV survey data and updates FMV Calculator
Author: Production System
Version: 2.1 - Added DVL_updated.xlsx export functionality
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
DVL_UPDATED_FILE = os.path.join(FOLDER_PATH, "DVL_updated.xlsx")  # New output file
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
      
