import pandas as pd
import os
import warnings
import time

# === SUPPRESS WARNINGS ===
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# === CONFIGURATION ===
base_path = "./data/raw"
output_file = "./data/output/Final.xlsx"

all_data = []

# === WALK THROUGH ALL FILES ===
for root, dirs, files in os.walk(base_path):
    for file in files:
        if file.endswith(".xlsx") and file != "Final.xlsx" and not file.startswith("~$"):
            file_path = os.path.join(root, file)
            try:
                # Optimized read
                df = pd.read_excel(file_path, header=8, usecols="C:AQ", engine="openpyxl")
                
                if not df.empty:
                    # Filter: Remove blanks and 'Total' rows
                    mask = df.iloc[:, 0].notna() & ~df.iloc[:, 0].astype(str).str.contains("(?i)total")
                    df = df[mask].copy()
                    
                    df.insert(0, "Source File", file)
                    all_data.append(df)
                    print(f"✅ Processed: {file}")
                
            except Exception as e:
                print(f"⚠️ Skipped {file} - error: {e}")

# === COMBINE & EXPORT ===
if all_data:
    final_df = pd.concat(all_data, ignore_index=True)
    
    # Attempt to save with retry logic for OneDrive locks
    for attempt in range(3):
        try:
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False, sheet_name="Combined_Data")
            print(f"\n✅ Success! Combined file saved to: {output_file}")
            break
        except PermissionError:
            print(f"⚠️ File locked (attempt {attempt+1}/3). Retrying in 5s...")
            time.sleep(5)
    else:
        print("\n❌ Failed to save: File is being used by another process.")
else:
    print("\n⚠️ No valid data found.")
