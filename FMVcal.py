import pandas as pd
import os

# Define the folder path
folder_path = os.path.expanduser("C:/Users/PAWARUX1/Desktop/FMV Automation")

# File paths
doctor_data_path = os.path.join(folder_path, "Extracted_Doctor_Data.xlsx")
cvdump_path = os.path.join(folder_path, "CVdump.xlsx")
fmv_calc_path = os.path.join(folder_path, "FMV_Calculator.xlsx")

# Load Excel files
doctor_df = pd.read_excel(doctor_data_path)
cvdump_df = pd.read_excel(cvdump_path)
fmv_df = pd.read_excel(fmv_calc_path)

# Clean emails (strip spaces, lowercase)
doctor_df["Email"] = doctor_df["Email"].astype(str).str.strip().str.lower()
cvdump_df["HCP Email"] = cvdump_df["HCP Email"].astype(str).str.strip().str.lower()

# Mapping from CVdump column names to FMV_Calculator column names
column_mapping = {
    "Clinical Experience: i.e. Time Spent with Patients?": "Clinical Experience",
    "Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...": "Leadership position in scientific Society / Hospital / Patient care",
    "Geographic influence as a Key Opinion Leader.": "Geographical Reach",
    "Highest Academic Position Held in past 10 years": "Highest Academic position ",
    "Additional Educational Level": "Additional Educational Level",
    "Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years": "Research Experience",
    "Publication experience in the past 10 years": "Publication Experience",
    "Speaking experience (professional, academic, scientific, or media experience) in the past 10 years.": "Speaking Experience",
    "Years of experience in the Specialty / Super Specialty?": "Years of Experience "
}

# Rename CVdump columns according to mapping
cvdump_renamed = cvdump_df.rename(columns=column_mapping)

# Merge doctor_df with cvdump_df on email
merged_df = pd.merge(
    doctor_df,
    cvdump_renamed,
    how="left",
    left_on="Email",
    right_on="HCP Email"
)

# Build final dataframe
final_columns = [
    "HCP Name", "DVL Code", "Email",
    "Years of Experience ", "Clinical Experience",
    "Leadership position in scientific Society / Hospital / Patient care",
    "Geographical Reach", "Highest Academic position ",
    "Additional Educational Level", "Research Experience",
    "Publication Experience", "Speaking Experience"
]

# Ensure missing columns exist
for col in final_columns:
    if col not in merged_df.columns:
        merged_df[col] = None

# Select only final columns
matched_df = merged_df[final_columns].copy()

# Optional: drop duplicates if multiple matches
matched_df = matched_df.drop_duplicates()

# Append to existing FMV data
updated_fmv_df = pd.concat([fmv_df, matched_df], ignore_index=True)

# Save the updated file
updated_fmv_df.to_excel(fmv_calc_path, index=False)

print(f"✅ Successfully appended {len(matched_df)} matched records to FMV_Calculator.xlsx.")
