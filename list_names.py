import pandas as pd

file_path = "attendance_template.xlsx"

df = pd.read_excel(file_path, sheet_name="Employee Master", header=None)

# employee names are in column 1
names = (
    df.iloc[7:, 1]   # skip headers
    .dropna()
    .astype(str)
    .str.replace(r"\t", " ", regex=True)
    .str.replace(r"\s+", " ", regex=True)
    .str.strip()
    .unique()
)

print("CLEAN STAFF NAMES:\n")
for n in names:
    print(n)