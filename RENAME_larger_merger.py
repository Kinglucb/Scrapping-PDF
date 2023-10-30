import pandas as pd

def get_largest_table_from_excel(file_path):
    # Load the Excel file
    xls = pd.ExcelFile(file_path)
    
    # Track which sheet has the most rows
    max_rows = 0
    largest_table_sheet_name = None
    
    # Iterate over sheets to find the one with the most rows
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        if len(df) > max_rows:
            max_rows = len(df)
            largest_table_sheet_name = sheet_name
    
    # Read the largest table
    largest_table = pd.read_excel(file_path, sheet_name=largest_table_sheet_name)
    
    return largest_table

# Path to the merged Excel file
file_path = "merged_output.xlsx"

# Get the largest table
largest_table_df = get_largest_table_from_excel(file_path)

# If needed, you can save this largest table to a new Excel or CSV file
largest_table_df.to_excel("largest_table.xlsx", index=False)
# largest_table_df.to_csv("largest_table.csv", index=False)
