import pandas as pd

csv_path = r"C:\Users\2750834\Downloads\b1b5cbd0-b789-4062-a850-e15759f9db7d-export_11818295_1771568422.csv"
jh_path = r"C:\Users\2750834\Downloads\JH_7204862_200226_11-32-05.xlsx"
wh_path = r"C:\Users\2750834\Downloads\Warehouse Associate - Delhi-NCR (4).xlsx"
out_path = r"C:\Users\2750834\Downloads\Candidates_Combined.xlsx"

csv_df = pd.read_csv(csv_path, usecols=["Full Name", "Mobile No."])
csv_df.columns = ["Candidate Name", "Candidate No."]

jh_df = pd.read_excel(jh_path, usecols=["Full Name", "Phone Number"])
jh_df.columns = ["Candidate Name", "Candidate No."]

wh_df = pd.read_excel(wh_path, usecols=["Candidate Name", "Phone number"])
wh_df.columns = ["Candidate Name", "Candidate No."]

combined = pd.concat([csv_df, jh_df, wh_df], ignore_index=True)
combined.to_excel(out_path, index=False)

print(f"Done! Total rows: {len(combined)}")
print(f"Saved to: {out_path}")
