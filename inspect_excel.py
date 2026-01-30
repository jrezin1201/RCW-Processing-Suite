"""Inspect an Excel file to see its structure."""
import sys
sys.path.append('.')
import pandas as pd

filepath = 'data/uploads/ee2ccca9-177a-49ff-8da7-1c90187139cd.xlsx'

# Read the file
df = pd.read_excel(filepath, nrows=10)  # Just read first 10 rows

print(f"File: {filepath}")
print(f"Shape: {df.shape[0]} rows, {df.shape[1]} columns")
print("\n=== Column Headers ===")
for i, col in enumerate(df.columns, 1):
    print(f"  {i}. '{col}'")

print("\n=== First 5 Rows (sample data) ===")
for idx, row in df.head().iterrows():
    print(f"\nRow {idx + 1}:")
    for col in df.columns[:5]:  # Show first 5 columns only
        val = row[col]
        if pd.notna(val):
            print(f"  {col}: {val}")

# Check for headers in first 10 rows
print("\n=== Checking for Lennar headers in data ===")
required = ["lot/block", "plan", "task", "task start date"]
for row_idx in range(min(10, len(df))):
    row = df.iloc[row_idx]
    row_str = ' '.join(str(v).lower() for v in row.values if pd.notna(v))
    matches = [h for h in required if h in row_str]
    if matches:
        print(f"Row {row_idx}: Found {matches}")

print("\n=== Required Lennar Headers ===")
print("The service expects these column headers:")
print("  1. Lot/Block")
print("  2. Plan")
print("  3. Task")
print("  4. Task Start Date")
print("\nIf your file doesn't have these exact headers, it won't be processed.")