import pandas as pd
import os

# Get CSV path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
csv_path = os.path.join(SCRIPT_DIR, "output_categorized(2025).csv")

# Read CSV
df = pd.read_csv(csv_path)

# Create NET_FINAL
df["NET_FINAL"] = (
    df["NET_Total (Hours)"].fillna(0)
    + df["NET_Rework (Hours)"].fillna(0)
)

# Create DCOM_DEM_FINAL
df["DCOM_DEM_FINAL"] = (
    df["DCOM_Total (Hours)"].fillna(0)
    + df["DEM_Total (Hours)"].fillna(0)
    + df["DCOM_Rework (Hours)"].fillna(0)
    + df["DEM_Rework (Hours)"].fillna(0)
)

# Save updated file
output_path = os.path.join(
    SCRIPT_DIR,
    "output_categorized_with_final.csv"
)

df.to_csv(output_path, index=False)

print("Done!")
print(f"Saved to: {output_path}")
