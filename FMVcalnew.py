import pandas as pd
import os

# -------- File paths --------
folder = r"C:\Users\PAWARUX1\Desktop\FMV Automation"
fmv_file = os.path.join(folder, "FMV_Calculator.csv")
cvdump_file = os.path.join(folder, "CVdump.csv")
dvl_file = os.path.join(folder, "DVL.csv")
missing_file = os.path.join(folder, "Missing_Doctors.csv")

# -------- Columns we need --------
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

# -------- Safe CSV loader --------
def safe_read_csv(path, usecols=None):
    try:
        return pd.read_csv(path, dtype=str, usecols=usecols, encoding="utf-8")
    except UnicodeDecodeError:
        return pd.read_csv(path, dtype=str, usecols=usecols, encoding="latin1")

print("📂 Loading files...")
fmv = safe_read_csv(fmv_file)
cvdump = safe_read_csv(cvdump_file, usecols=lambda x: x in cvdump_cols)
dvl = safe_read_csv(dvl_file)

# -------- Clean data --------
cvdump["Start time"] = pd.to_datetime(cvdump["Start time"], errors="coerce")
cvdump = cvdump.sort_values("Start time").drop_duplicates("HCP Email", keep="last")

# -------- Merge --------
print("🔎 Matching emails...")
merged = pd.merge(
    dvl.rename(columns={"Account: Email": "HCP Email", "Customer Code": "DVL Code"}),
    cvdump,
    on="HCP Email",
    how="left"
)

# -------- Find missing --------
missing = merged[merged["Start time"].isna()][["DVL Code", "HCP Email"]]
if not missing.empty:
    print(f"⚠️ {len(missing)} emails not found in CVdump → saved to Missing_Doctors.csv")
    missing.to_csv(missing_file, index=False, encoding="utf-8-sig")

# -------- Keep only found entries --------
new_data = merged[merged["Start time"].notna()].copy()

# -------- Ensure same column order as FMV --------
cols_to_add = [c for c in fmv.columns if c in new_data.columns]
final_new = new_data[cols_to_add]

# -------- Append only new emails --------
existing_emails = set(fmv["HCP Email"].dropna())
final_new = final_new[~final_new["HCP Email"].isin(existing_emails)]

if not final_new.empty:
    print(f"✅ Appending {len(final_new)} new rows to FMV_Calculator...")
    fmv = pd.concat([fmv, final_new], ignore_index=True)
    fmv.to_csv(fmv_file, index=False, encoding="utf-8-sig")
    print("💾 FMV_Calculator updated successfully.")
else:
    print("ℹ️ No new doctors to append. FMV_Calculator unchanged.")
