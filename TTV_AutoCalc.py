import os, sys
import pandas as pd
from openpyxl import load_workbook

# File paths in same directory as script
current_dir = os.getcwd()
print(f"Current directory: {current_dir}")
template_path = current_dir + "\TTV_template.xlsx"
side1_path = current_dir + "\side1.txt"
side2_path = current_dir + "\side2.txt"

# Load TXT files (space separated XYZ)
side1 = pd.read_csv(side1_path, sep=r"\s+", header=None, names=["x", "y", "z"])
side2 = pd.read_csv(side2_path, sep=r"\s+", header=None, names=["x", "y", "z"])

# Open Excel template
wb = load_workbook(template_path)

# Sheet names
ws1 = wb["Side 1"]  # change if needed
ws2 = wb["Side 2"]  # change if needed

# Write Z values
# Change column letter + start row to match template
start_row = 3
column_letter = "D"   # change if needed

for i, value in enumerate(side1["z"], start=start_row):
    ws1[f"{column_letter}{i}"] = float(value)

for i, value in enumerate(side2["z"], start=start_row):
    ws2[f"{column_letter}{i}"] = float(value)

# Save output under new name
output_path = current_dir + "\TTV_populated.xlsx"
wb.save(output_path)

print("TTV file populated successfully.")
