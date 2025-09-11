import pandas as pd
from datetime import datetime
import os

# -------- File Paths --------
base_path = r"C:\Users\PAWARUX1\Desktop\FMV Automation"
fmv_file = os.path.join(base_path, "FMV_Calculator.csv")
cvdump_file = os.path.join(base_path, "CVdump.csv")
dvl_file = os.path.join(base_path, "DVL.csv")
missing_file = os.path.join(base_path, "Missing_Doctors.csv")

# -------- Columns of interest --------
cvdump_cols = [
    "Start time", "HCP Email",
    "Clinical Experience: i.e. Time Spent with Patients?",
    "Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...",
    "Geographic influence as a Key Opinion Leader.",
    "Highest Academic Position Held in past 10 years",
    "Educational Qualification", "Additional Educational Level",
    "Specialty / Super Specialty",
    "Years of experience in the Specialty / Super Specialty?",
    "Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years",
    "Publication experience in the past 10 years",
    "Speaking experience (professional, academic, scientific, or media experience) in the past 10 years."
]

# -------- Load files --------
print("📂 Loading files...")
fmv = pd.read_csv(fmv_file, dtype=str)
cvdump = pd.read_csv(cvdump_file, dtype=str, usecols=lambda x: x in cvdump_cols)
dvl = pd.read_csv(dvl_file, dtype=str)

# Ensure consistent column names
fmv.columns = fmv.columns.str.strip()
cvdump.columns = cvdump.columns.str.strip()
dvl.columns = dvl.columns.str.strip()

# -------- Process Start time --------
cvdump["Start time"] = pd.to_datetime(cvdump["Start time"], errors="coerce")

# Keep only latest entry per HCP Email
cvdump = cvdump.sort_values("Start time").drop_duplicates("HCP Email", keep="last")

# -------- Matching --------
missing_emails = []
new_rows = []

for _, row in dvl.iterrows():
    email = row["Account: Email"]
    dvl_code = row["Customer Code"]

    match = cvdump[cvdump["HCP Email"].str.lower() == str(email).lower()]

    if match.empty:
        missing_emails.append({"Account: Email": email, "Customer Code": dvl_code})
        continue

    match = match.iloc[0]  # latest entry

    # Build new FMV row
    new_entry = {col: "" for col in fmv.columns}  # blank row with same headers
    for col in cvdump_cols:
        if col in fmv.columns:
            new_entry[col] = match[col]
    new_entry["HCP Email"] = email
    new_entry["DVL Code"] = dvl_code

    # Append only if not already in FMV_Calculator
    if not (fmv["HCP Email"].str.lower() == email.lower()).any():
        new_rows.append(new_entry)

# -------- Append and Save --------
if new_rows:
    fmv = pd.concat([fmv, pd.DataFrame(new_rows)], ignore_index=True)

fmv.to_csv(fmv_file, index=False, encoding="utf-8-sig")

if missing_emails:
    pd.DataFrame(missing_emails).to_csv(missing_file, index=False, encoding="utf-8-sig")

print(f"✅ FMV_Calculator updated. Added {len(new_rows)} new rows.")
print(f"🚨 Missing doctors saved to {missing_file} ({len(missing_emails)} emails).")
