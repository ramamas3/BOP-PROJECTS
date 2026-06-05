import pandas as pd

# File configuration
file_path = r'C:\Users\RAMAMANIS\OneDrive - IBM\Desktop\python\BOB\Stocking Strategy Tracking.xlsx'
sheet_name = "Stocking stratergy by PN -New"

# Load data
print("Loading data...")
df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

# Clean column names
df.columns = df.columns.str.strip()

print(f"\nData loaded successfully!")
print(f"   - Total rows: {len(df)}")
print(f"   - Total columns: {len(df.columns)}")

print("\nAvailable columns:")
for i, col in enumerate(df.columns, 1):
    print(f"   {i}. {col}")

print("\nSample data (first 3 rows):")
print(df.head(3))

# Made with Bob
