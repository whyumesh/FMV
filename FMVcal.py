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

# Mapping from CVdump column names to FMV_Calculator column names
column_mapping = {
    "Clinical Experience: i.e. Time Spent with Patients?": "Clinical Experience",
    "Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...": "Leadership position in scientific Society / Hospital / Patient care",
    "Geographic influence as a Key Opinion Leader.": "Geographical Reach",
    "Highest Academic Position Held in past 10 years": "Highest Academic position",
    "Additional Educational Level": "Additional Educational Level",
    "Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years": "Research Experience",
    "Publication experience in the past 10 years": "Publication Experience",
    "Speaking experience (professional, academic, scientific, or media experience) in the past 10 years.": "Speaking Experience"
}

# Prepare list to collect matched rows
matched_rows = []

# Iterate over each email in doctor_df
for _, row in doctor_df.iterrows():
    email = str(row["Email"]).strip().lower()
    match = cvdump_df[cvdump_df["HCP Email"].astype(str).str.strip().str.lower() == email]

    if not match.empty:
        matched_data = {
            "Doctor Name": row.get("Doctor Name", ""),
            "DVL Code": row.get("DVL Code", ""),
            "Email": email,
            "Years of Experience": None  # Placeholder if not available
        }
        for cv_col, fmv_col in column_mapping.items():
            matched_data[fmv_col] = match.iloc[0].get(cv_col, None)
        matched_rows.append(matched_data)

# Convert matched rows to DataFrame
matched_df = pd.DataFrame(matched_rows)

# Ensure all required columns exist
final_columns = [
    "Doctor Name", "DVL Code", "Email",
    "Years of Experience", "Clinical Experience",
    "Leadership position in scientific Society / Hospital / Patient care",
    "Geographical Reach", "Highest Academic position",
    "Additional Educational Level", "Research Experience",
    "Publication Experience", "Speaking Experience"
]

for col in final_columns:
    if col not in matched_df.columns:
        matched_df[col] = None

# Reorder columns
matched_df = matched_df[final_columns]

# Append to existing FMV data
updated_fmv_df = pd.concat([fmv_df, matched_df], ignore_index=True)

# Save the updated file
updated_fmv_df.to_excel(fmv_calc_path, index=False)

print(f"✅ Successfully appended {len(matched_df)} matched records to FMV_Calculator.xlsx.")
