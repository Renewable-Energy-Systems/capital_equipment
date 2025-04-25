import pandas as pd, pathlib, tempfile, shutil
xlsx = pathlib.Path(r"C:/Users/PRANAY-RES/OneDrive - Renewable Energy Systems Limited/RES/Capital Equipment 2025/Capital_Equipment.xlsx")

with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
    tmp_path = pathlib.Path(tmp.name)
shutil.copy2(xlsx, tmp_path)
df = pd.read_excel(tmp_path, sheet_name="cp_list", header=2).iloc[1:]
tmp_path.unlink(missing_ok=True)

print("Columns:", list(df.columns))
print("\nSample rows:")
print(df.head(2).to_string(index=False))
