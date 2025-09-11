import pandas as pd
import os
from openpyxl import load_workbook

# ========== CONFIG ==========
folder_path = r"C:\Users\PAWARUX1\Desktop\FMV Automation"
fmv_file = os.path.join(folder_path, "FMV_Calculator.xlsx")
cvdump_file = os.path.join(folder_path, "CVdump.xlsx")
dvl_file = os.path.join(folder_path, "DVL.xlsx")
missing_log_file = os.path.join(folder_path, "Missing_Doctors.xlsx")

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

# ========== SAFE READER ==========
def safe_read_excel(file, usecols=None):
    """
    Try reading Excel normally; if it fails (due to corrupted styles),
    fallback to openpyxl value-only reader (ignores styles).
    """
    try:
        return pd.read_excel(file, dtype=str, usecols=usecols, engine="openpyxl")
    except Exception as e:
        print(f"⚠️ Normal read failed for {file}: {e}")
        print("👉 Retrying with style-stripped mode...")
        wb = load_workbook(file, read_only=True, data_only=True)
        sheet = wb.active
        data = sheet.values
        cols = next(data)
        df = pd.DataFrame(data, columns=cols)
        if usecols:
            df = df[[c for c in df.columns if c in usecols]]
        return df.astype(str).fillna("")

# ========== MAIN PIPELINE ==========
print("📂 Loading files...")
fmv = safe_read_excel(fmv_file)
cvdump = safe_read_excel(cvdump_file, usecols=cvdump_cols)
dvl = safe_read_excel(dvl_file, usecols=dvl_cols)

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
        fmv[col] = None

# Prepare rows in FMV format
cols_to_add = ["DVL Code"] + [c for c in cvdump_cols if c in merged.columns]
new_rows = merged[cols_to_add]

# Deduplication check (skip existing doctors)
if "DVL Code" in fmv.columns:
    new_rows = new_rows[~new_rows["DVL Code"].isin(fmv["DVL Code"])]
if "HCP Email" in fmv.columns and "HCP Email" in new_rows.columns:
    new_rows = new_rows[~new_rows["HCP Email"].isin(fmv["HCP Email"])]

# Append into FMV
before = len(fmv)
fmv = pd.concat([fmv, new_rows], ignore_index=True)
added = len(fmv) - before

# Save back into SAME FMV file
fmv.to_excel(fmv_file, index=False)

# Track missing doctors (present in DVL but not in CVdump)
missing = dvl[~dvl["Account: Email"].isin(cvdump["HCP Email"])]
if not missing.empty:
    missing.to_excel(missing_log_file, index=False)
    print(f"⚠️ Missing doctors logged: {len(missing)} (saved to {missing_log_file})")

print(f"✅ FMV_Calculator updated successfully. Added {added} new rows.")
