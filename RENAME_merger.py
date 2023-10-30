import pandas as pd

# Normalize column names by removing extra spaces
def normalize_column_name(col_name):
    return col_name.replace(" ", "").strip()

# Load the Excel file
file_path = "output.xlsx"
xls = pd.ExcelFile(file_path)

# Read all sheets into a list of dataframes
all_dfs = [xls.parse(sheet_name, header=None) for sheet_name in xls.sheet_names]

# Group dataframes by the first header in row 2
grouped_dfs = {}
for df in all_dfs:
    # Normalize column names
    df.iloc[1] = df.iloc[1].apply(normalize_column_name)
    first_header = df.iloc[1, 0]
    
    if first_header not in grouped_dfs:
        grouped_dfs[first_header] = []
    
    # Set the headers correctly and drop the header row
    df.columns = df.iloc[1]
    df = df.drop([0, 1]).reset_index(drop=True)
    grouped_dfs[first_header].append(df)

# Merge dataframes within each group
merged_dataframes = []
for key, group in grouped_dfs.items():
    # Concatenate dataframes vertically
    merged = pd.concat(group, axis=0, ignore_index=True)
    merged_dataframes.append(merged)

# Save the merged tables to a new Excel file
output_file_path = "merged_output.xlsx"
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    for idx, df in enumerate(merged_dataframes):
        df.to_excel(writer, sheet_name=f'Merged_Table_{idx+1}', index=False)
