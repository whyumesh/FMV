import pandas as pd
import os
from datetime import datetime

# ------------------ CONFIG ------------------
folder_path = r"C:\Users\PAWARUX1\Desktop\FMV Automation"
fmv_file = os.path.join(folder_path, "FMV_Calculator.xlsx")
cvdump_file = os.path.join(folder_path, "CVdump.xlsx")
dvl_file = os.path.join(folder_path, "DVL.xlsx")
missing_file = os.path.join(folder_path, "Missing_Doctors.xlsx")
backup_file = fmv_file.replace(".xlsx", f"_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
# --------------------------------------------

def safe_read_excel(file, **kwargs):
    """Read Excel safely, ignoring style corruption issues."""
    try:
        return pd.read_excel(file, dtype=str, engine="openpyxl", **kwargs)
    except Exception:
        return pd.read_excel(file, dtype=str, **kwargs)

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

# ------------------ LOAD FILES ------------------
print("📂 Loading Excel files...")
fmv = safe_read_excel(fmv_file)
cvdump = safe_read_excel(cvdump_file, usecols=lambda x: x in cvdump_cols)
dvl = safe_read_excel(dvl_file, usecols=lambda x: x in dvl_cols)

# ------------------ NORMALIZE EMAILS ------------------
for df, col in [(cvdump, "HCP Email"), (dvl, "Account: Email"), (fmv, "HCP Email")]:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.lower()

# ------------------ CLEAN DUPLICATES ------------------
cvdump = cvdump.drop_duplicates(subset=["HCP Email"], keep="first")
dvl = dvl.drop_duplicates(subset=["Account: Email"], keep="first")

# ------------------ MERGE ------------------
merged = pd.merge(
    dvl, cvdump,
    left_on="Account: Email", right_on="HCP Email",
    how="left"   # left join so we can track missing ones too
)

# ------------------ SPLIT FOUND vs MISSING ------------------
found = merged.dropna(subset=["HCP Email"]).copy()
missing = merged[merged["HCP Email"].isna()].copy()

# Add DVL Code
if "Customer Code" in found.columns:
    found["DVL Code"] = found["Customer Code"]
else:
    found["DVL Code"] = None

# ------------------ ALIGN WITH FMV ------------------
for col in list(cvdump_cols) + ["DVL Code"]:
    if col not in fmv.columns:
        fmv[col] = None

# Prepare rows in FMV format
cols_to_add = ["DVL Code"] + [c for c in cvdump_cols if c in found.columns]
new_rows = found[cols_to_add]

# ------------------ DEDUPLICATION ------------------
before = len(new_rows)
if "DVL Code" in fmv.columns:
    new_rows = new_rows[~new_rows["DVL Code"].isin(fmv["DVL Code"])]
if "HCP Email" in fmv.columns and "HCP Email" in new_rows.columns:
    new_rows = new_rows[~new_rows["HCP Email"].isin(fmv["HCP Email"])]
after = len(new_rows)
skipped_duplicates = before - after

# ------------------ APPEND ------------------
added_count = len(new_rows)
fmv = pd.concat([fmv, new_rows], ignore_index=True)

# ------------------ SAVE FILES ------------------
# Backup original FMV
os.rename(fmv_file, backup_file)
print(f"🛡️ Backup created: {backup_file}")

# Save updated FMV
fmv.to_excel(fmv_file, index=False)
print(f"✅ FMV_Calculator updated in place. Added {added_count} new rows, skipped {skipped_duplicates} duplicates.")

# Save missing doctors log
if not missing.empty:
    missing[["Customer Code", "Account: Email", "Tier Type", "Account: Account Name"]].to_excel(missing_file, index=False)
    print(f"⚠️ {len(missing)} doctors from DVL not found in CVdump. Logged in {missing_file}")
else:
    print("🎉 All DVL doctors matched in CVdump. No missing entries.")
