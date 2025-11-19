#!/usr/bin/env python3
"""
CV FMV Calculator - Production Level Script
Calculates Fair Market Value (FMV) for all doctor entries from CVdump.csv
Based on scoring_criteria.xlsx and OUS FMV Rates
Author: AI Assistant
Version: 1.0
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
        logging.FileHandler('cv_fmv_calculator.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# File paths
CVDUMP_FILE = "CVdump.csv"
SCORING_CRITERIA_FILE = "scoring_criteria.xlsx"
OUTPUT_FILE = f"CV_FMV_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# =============================================================================
# SCORING CRITERIA LOOKUP DICTIONARIES
# =============================================================================

def create_scoring_lookup():
    """Create comprehensive scoring lookup dictionaries"""
    
    # Years of Experience scoring
    years_experience_scores = {
        "1-2 years of experience": 0,
        "3-7 years of experience": 2,
        "8-14 years of experience": 4,
        "15+ years of experience": 6
    }
    
    # Clinical Experience scoring
    clinical_experience_scores = {
        "Minimal patient interactions and predominantly administrative/academic work": 0,
        "Less than half the time spent with patients in clinical setting and higher focus on academic/administrative work": 2,
        "Equal amount of time spent with patients in clinical setting and equal amount of time spent in academic/administrative work": 4,
        "Significant time spent with patients in clinical setting and minimal time spent in academic/administrative work": 6
    }
    
    # Leadership position scoring
    leadership_scores = {
        "Not applicable, as not a part of any society or leadership roles in hospital": 0,
        "1-2 years in a leadership position(s) eg. HOD of a particular speciality in Hospital or other Patient Care Setting and/or serving as a President, Vice president, Secretary,Treasurer, Board member in a Professional or Scientific Society.": 2,
        "3-7 years in a leadership position(s) eg HOD of a particular speciality   in Hospital or other Patient Care Setting and/or serving as a national/regional leader in a Professional or Scientific Society.": 4,
        "8 or more years in a leadership position(s) eg HOD for a specialty in Hospital or other Patient Care Setting and/or serving as an international leader in a Professional or Scientific Society.": 6
    }
    
    # Geographical Reach scoring
    geographical_reach_scores = {
        "Local Influence": 0,
        "National Influence": 2,
        "Multi-Country Influence": 4,
        "Global/Worldwide Influence": 6
    }
    
    # Highest Academic Position scoring
    academic_position_scores = {
        "None or N/A": 0,
        "Professor (including Associate / Assistant Professor)": 2,
        "Professor or Adjunct/Additional/Emeritus Professor": 4,
        "Department Chair/ HOD (or similar position)": 6
    }
    
    # Additional Educational Level scoring
    educational_level_scores = {
        "None or N/A": 0,
        "1 Additional degree, fellowship, or advanced training certification.": 2,
        "2 Additional degrees, fellowship, or advanced training certification.": 4,
        "3 or More Additional degrees, fellowship, or advanced training certification.": 6
    }
    
    # Research Experience scoring
    research_experience_scores = {
        "None or N/A": 0,
        "Participation as an Investigator or Sub-Investigator in 1 to 4 clinical trials or research studies.": 2,
        "Participation as an Investigator or Sub-Investigator in 5 to 9 clinical trials or research studies.": 4,
        "Participation as an Investigator of Sub-Investigator in 10 or more clinical trials or research studies or Principal Investigator for two or more clinical trials or research studies or serving as the Principal Investigator for a clinical trial or research study that led to important medical innovations or significant medical technology breakthroughs.": 6
    }
    
    # Publication Experience scoring
    publication_experience_scores = {
        "None or N/A": 0,
        "Co-authorship or participation as contributing author on 1 to 4 publications.": 2,
        "First authorship (if known) on 1 to 5 publications and/or co-authorship or participation as contributing author on 6 to 10 publications": 4,
        "First authorship (if known) on 6 or more publications and/or co-authorship or participation as contributing author on 11 or more publications": 6
    }
    
    # Speaking Experience scoring
    speaking_experience_scores = {
        "Local speaking engagements and the scientific work done for the specialty is near to the practice location": 0,
        "Most of the speaking engagements are directed nationally for the conferences, symposia or national webinars in the designated specialty and the scientific work done is not restricted for the local audience": 2,
        "The speaking experiences are not restricted nationally but to a group of specified countries and the scientific work is directed to the same group of countries": 4,
        "The speaking engagements and the scinetific work carried out is across the globe": 6
    }
    
    return {
        "years_experience": years_experience_scores,
        "clinical_experience": clinical_experience_scores,
        "leadership": leadership_scores,
        "geographical_reach": geographical_reach_scores,
        "academic_position": academic_position_scores,
        "educational_level": educational_level_scores,
        "research_experience": research_experience_scores,
        "publication_experience": publication_experience_scores,
        "speaking_experience": speaking_experience_scores
    }

# =============================================================================
# FMV RATES LOADING
# =============================================================================

def load_fmv_rates():
    """Load FMV rates from OUS FMV Rates sheet"""
    try:
        rates_df = pd.read_excel(SCORING_CRITERIA_FILE, sheet_name="OUS FMV Rates", header=1)
        
        # Filter for India rates
        india_rates = rates_df[rates_df['Country'] == 'India'].copy()
        
        # Create specialty to rates mapping
        specialty_rates = {}
        for _, row in india_rates.iterrows():
            specialty = row['HCP Specialty']
            if pd.notna(specialty):
                specialty_rates[specialty] = {
                    'Tier 1': row['Tier 1'],
                    'Tier 2': row['Tier 2'],
                    'Tier 3': row['Tier 3'],
                    'Tier 4': row['Tier 4']
                }
        
        logger.info(f"Loaded FMV rates for {len(specialty_rates)} specialties")
        return specialty_rates
    except Exception as e:
        logger.error(f"Error loading FMV rates: {str(e)}")
        raise

# =============================================================================
# SCORING FUNCTIONS
# =============================================================================

def calculate_individual_scores(row, scoring_lookup):
    """Calculate individual scores for each criterion"""
    scores = {}
    
    # Years of Experience (Score 1)
    years_exp = str(row.get("Years of experience in the Specialty / Super Specialty?\n", "")).strip()
    scores["score_1"] = scoring_lookup["years_experience"].get(years_exp, 0)
    
    # Clinical Experience (Score 2)
    clinical_exp = str(row.get("Clinical Experience: i.e. Time Spent with Patients?", "")).strip()
    scores["score_2"] = scoring_lookup["clinical_experience"].get(clinical_exp, 0)
    
    # Leadership position (Score 3)
    leadership = str(row.get("Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...", "")).strip()
    scores["score_3"] = scoring_lookup["leadership"].get(leadership, 0)
    
    # Geographical Reach (Score 4)
    geo_reach = str(row.get("Geographic influence as a Key Opinion Leader.", "")).strip()
    scores["score_4"] = scoring_lookup["geographical_reach"].get(geo_reach, 0)
    
    # Highest Academic Position (Score 5)
    academic_pos = str(row.get("Highest Academic Position Held in past 10 years", "")).strip()
    scores["score_5"] = scoring_lookup["academic_position"].get(academic_pos, 0)
    
    # Additional Educational Level (Score 6)
    edu_level = str(row.get("Additional Educational Level ", "")).strip()
    scores["score_6"] = scoring_lookup["educational_level"].get(edu_level, 0)
    
    # Research Experience (Score 7)
    research_exp = str(row.get("Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years", "")).strip()
    scores["score_7"] = scoring_lookup["research_experience"].get(research_exp, 0)
    
    # Publication Experience (Score 8)
    pub_exp = str(row.get("Publication experience in the past 10 years", "")).strip()
    scores["score_8"] = scoring_lookup["publication_experience"].get(pub_exp, 0)
    
    # Speaking Experience (Score 9)
    speaking_exp = str(row.get("Speaking experience (professional, academic, scientific, or media experience) in the past 10 years.", "")).strip()
    scores["score_9"] = scoring_lookup["speaking_experience"].get(speaking_exp, 0)
    
    # Calculate total score
    scores["total_score"] = sum(scores.values())
    
    return scores

def determine_tier(total_score):
    """Determine tier based on total score"""
    if total_score <= 13:
        return "Tier 1"
    elif total_score <= 26:
        return "Tier 2"
    elif total_score <= 40:
        return "Tier 3"
    else:
        return "Tier 4"

def calculate_fmv_amount(specialty, tier, fmv_rates):
    """Calculate FMV amount based on specialty and tier"""
    if specialty in fmv_rates and tier in fmv_rates[specialty]:
        return fmv_rates[specialty][tier]
    else:
        # Default to Cardiologist rates if specialty not found
        default_specialty = "Cardiologist"
        if default_specialty in fmv_rates and tier in fmv_rates[default_specialty]:
            return fmv_rates[default_specialty][tier]
        else:
            logger.warning(f"No FMV rate found for specialty: {specialty}, tier: {tier}")
            return 0

# =============================================================================
# DATA PROCESSING FUNCTIONS
# =============================================================================

def load_cvdump_data():
    """Load and clean CVdump.csv data"""
    try:
        logger.info("Loading CVdump.csv data...")
        
        # Try different encodings
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        df = None
        
        for encoding in encodings:
            try:
                df = pd.read_csv(CVDUMP_FILE, encoding=encoding)
                logger.info(f"Successfully loaded CVdump.csv with {encoding} encoding")
                break
            except UnicodeDecodeError:
                continue
        
        if df is None:
            raise Exception("Could not load CVdump.csv with any supported encoding")
        
        # Clean email addresses
        df["HCP Email"] = df["HCP Email"].astype(str).str.strip().str.lower()
        
        # Remove rows with invalid emails
        df = df[df["HCP Email"] != "nan"]
        df = df[df["HCP Email"] != ""]
        
        logger.info(f"Loaded {len(df)} records from CVdump.csv")
        return df
    except Exception as e:
        logger.error(f"Error loading CVdump data: {str(e)}")
        raise

def process_doctor_data(df, scoring_lookup, fmv_rates):
    """Process each doctor's data and calculate FMV"""
    results = []
    
    for index, row in df.iterrows():
        try:
            # Calculate individual scores
            scores = calculate_individual_scores(row, scoring_lookup)
            total_score = scores["total_score"]
            
            # Determine tier
            tier = determine_tier(total_score)
            
            # Get specialty
            specialty = str(row.get("Specialty / Super Specialty", "")).strip()
            
            # Calculate FMV amount
            fmv_amount = calculate_fmv_amount(specialty, tier, fmv_rates)
            
            # Create result record matching FMV_Calculator_Updated.xlsx structure
            result = {
                "i": index + 1,  # Sequential number
                "HCP Name": row.get("HCP Name", ""),
                "Years of experience in the Specialty / Super Specialty?_x000D_\n": row.get("Years of experience in the Specialty / Super Specialty?\n", ""),
                "Clinical Experience: i.e. Time Spent with Patients?": row.get("Clinical Experience: i.e. Time Spent with Patients?", ""),
                "Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...": row.get("Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...", ""),
                "Geographic influence as a Key Opinion Leader.": row.get("Geographic influence as a Key Opinion Leader.", ""),
                "Highest Academic Position Held in past 10 years": row.get("Highest Academic Position Held in past 10 years", ""),
                "Additional Educational Level": row.get("Additional Educational Level ", ""),
                "Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years": row.get("Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years", ""),
                "Publication experience in the past 10 years": row.get("Publication experience in the past 10 years", ""),
                "Speaking experience (professional, academic, scientific, or media experience) in the past 10 years.": row.get("Speaking experience (professional, academic, scientific, or media experience) in the past 10 years.", ""),
                "Score based on selection mentioned criteria": total_score,
                "Score 1": scores["score_1"],
                "Score 2": scores["score_2"],
                "Score 3": scores["score_3"],
                "Score 4": scores["score_4"],
                "Score 5": scores["score_5"],
                "Score 6": scores["score_6"],
                "Score 7": scores["score_7"],
                "Score 8": scores["score_8"],
                "Score 9": scores["score_9"],
                "Range": f"{total_score}-{total_score}",  # Individual score range
                "Tier": tier,
                "Rate of Honorarium": fmv_amount,
                "Specialty / Super Specialty": specialty,
                "HCP Email": row.get("HCP Email", ""),
                "Educational Qualification": row.get("Educational Qualification", "")
            }
            
            results.append(result)
            
        except Exception as e:
            logger.error(f"Error processing doctor {row.get('HCP Name', 'Unknown')}: {str(e)}")
            continue
    
    return results

def save_results(results):
    """Save results to Excel file"""
    try:
        results_df = pd.DataFrame(results)
        
        # Create Excel file with single sheet matching FMV_Calculator_Updated.xlsx structure
        results_df.to_excel(OUTPUT_FILE, sheet_name='Sheet1', index=False)
        
        logger.info(f"Results saved to {OUTPUT_FILE}")
        return OUTPUT_FILE
        
    except Exception as e:
        logger.error(f"Error saving results: {str(e)}")
        raise

# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    """Main execution function"""
    try:
        logger.info("Starting CV FMV Calculator...")
        
        # Load scoring criteria
        logger.info("Loading scoring criteria...")
        scoring_lookup = create_scoring_lookup()
        
        # Load FMV rates
        logger.info("Loading FMV rates...")
        fmv_rates = load_fmv_rates()
        
        # Load CVdump data
        logger.info("Loading CVdump data...")
        cvdump_df = load_cvdump_data()
        
        # Process doctor data
        logger.info("Processing doctor data and calculating FMV...")
        results = process_doctor_data(cvdump_df, scoring_lookup, fmv_rates)
        
        # Save results
        logger.info("Saving results...")
        output_file = save_results(results)
        
        logger.info(f"✅ CV FMV Calculator completed successfully!")
        logger.info(f"📊 Processed {len(results)} doctors")
        logger.info(f"📁 Results saved to: {output_file}")
        
        # Print summary
        if results:
            total_fmv = sum(r['Rate of Honorarium'] for r in results)
            avg_score = sum(r['Score based on selection mentioned criteria'] for r in results) / len(results)
            tier_counts = {}
            for r in results:
                tier = r['Tier']
                tier_counts[tier] = tier_counts.get(tier, 0) + 1
            
            print(f"\n📈 SUMMARY:")
            print(f"   Total Doctors: {len(results)}")
            print(f"   Average Score: {avg_score:.2f}")
            print(f"   Total FMV Amount: ₹{total_fmv:,}")
            print(f"   Tier Distribution:")
            for tier, count in sorted(tier_counts.items()):
                print(f"     {tier}: {count} doctors")
        
    except Exception as e:
        logger.error(f"❌ Error in main execution: {str(e)}")
        logger.error(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main()
