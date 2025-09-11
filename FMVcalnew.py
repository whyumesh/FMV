import pandas as pd
import os

# ========== CONFIG ==========
folder_path = r"C:\Users\PAWARUX1\Desktop\FMV Automation"
fmv_file = os.path.join(folder_path, "FMV_Calculator.csv")
cvdump_file = os.path.join(folder_path, "CVdump.csv")
dvl_file = os.path.join(folder_path, "DVL.csv")
missing_log_file = os.path.join(folder_path, "Missing_Doctors.csv")

# Required columns
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
dvl_cols = ["Customer Code", "Account: Email", "Tier Type", "Account: Account Name"]

# ========== LOAD FILES ==========
print("📂 Loading CSV files...")
fmv = pd.read_csv(fmv_file, dtype=str).fillna("")
cvdump = pd.read_csv(cvdump_file, dtype=str, usecols=lambda x: x in cvdump_cols).fillna("")
dvl = pd.read_csv(dvl_file, dtype=str, usecols=lambda x: x in dvl_cols).fillna("")

# Normalize emails
for df, col in [(cvdump, "HCP Email"), (dvl, "Account: Email"), (fmv, "HCP Email")]:
    if col in df.columns:
        df[col] = df[col].str.strip().str.lower()

# Deduplicate
cvdump = cvdump.drop_duplicates(subset=["HCP Email"], keep="first")
dvl = dvl.drop_duplicates(subset=["Account: Email"], keep="first")

# Merge DVL + CVdump on email
merged = pd.merge(dvl, cvdump, left_on="Account: Email", right_on="HCP Email", how="inner")
merged["DVL Code"] = merged["Customer Code"]

# Ensure FMV has all required columns
for col in list(cvdump_cols) + ["DVL Code"]:
    if col not in fmv.columns:
        fmv[col] = ""

# Prepare rows in FMV format
cols_to_add = ["DVL Code"] + [c for c in cvdump_cols if c in merged.columns]
new_rows = merged[cols_to_add]

# Deduplication check (skip already present doctors)
if "DVL Code" in fmv.columns:
    new_rows = new_rows[~new_rows["DVL Code"].isin(fmv["DVL Code"])]
if "HCP Email" in fmv.columns and "HCP Email" in new_rows.columns:
    new_rows = new_rows[~new_rows["HCP Email"].isin(fmv["HCP Email"])]

# Append into FMV
before = len(fmv)
fmv = pd.concat([fmv, new_rows], ignore_index=True)
added = len(fmv) - before

# Save back into SAME FMV file
fmv.to_csv(fmv_file, index=False)

# Track missing doctors (present in DVL but not in CVdump)
missing = dvl[~dvl["Account: Email"].isin(cvdump["HCP Email"])]
if not missing.empty:
    missing.to_csv(missing_log_file, index=False)
    print(f"⚠️ Missing doctors logged: {len(missing)} (saved to {missing_log_file})")

print(f"✅ FMV_Calculator updated successfully. Added {added} new rows.")
