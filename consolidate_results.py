"""
Consolidate all Clearing_Results_*.xlsx files into one file.
"""
import glob
import os
import re
import sys
import pandas as pd
from datetime import datetime

folder = os.path.dirname(os.path.abspath(__file__))
pattern = os.path.join(folder, "Clearing_Results_*.xlsx")
files = sorted(glob.glob(pattern))

# Exclude any already-consolidated file
files = [f for f in files if "CONSOLIDATED" not in os.path.basename(f).upper()]

print(f"Found {len(files)} file(s):")
for f in files:
    print(f"  {os.path.basename(f)}")

if not files:
    print("Nothing to consolidate.")
    sys.exit(0)

frames = []
for f in files:
    name = os.path.basename(f)
    m = re.match(r"Clearing_Results_([A-Z]+)_(\d{8})_(\d{6})\.xlsx", name)
    system = m.group(1) if m else "UNKNOWN"
    ts_str = (m.group(2) + "_" + m.group(3)) if m else ""
    try:
        ts = datetime.strptime(ts_str, "%Y%m%d_%H%M%S").strftime("%Y-%m-%d %H:%M:%S") if ts_str else ""
    except Exception:
        ts = ts_str
    try:
        df = pd.read_excel(f, sheet_name="Clearing Results")
        df.insert(0, "System", system)
        df.insert(1, "Run Timestamp", ts)
        df.insert(2, "Source File", name)
        frames.append(df)
        print(f"  OK: {name}  ({len(df)} row(s))")
    except Exception as e:
        print(f"  ERROR: {name}: {e}")

if not frames:
    print("No data to combine.")
    sys.exit(1)

combined = pd.concat(frames, ignore_index=True)
out_name = f"Clearing_Results_CONSOLIDATED_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
out_path = os.path.join(folder, out_name)

with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    combined.to_excel(writer, index=False, sheet_name="All Results")

print(f"\nSaved: {out_name}")
print(f"Total rows: {len(combined)}")
