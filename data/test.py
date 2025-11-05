import pandas as pd

# Load your Excel file
df = pd.read_excel(r"backend\data\projects.xlsx")  # make sure the path is correct

# Print all column names
print(df.columns.tolist())
