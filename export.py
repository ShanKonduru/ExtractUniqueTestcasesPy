import pandas as pd
from datetime import datetime
import openpyxl

def get_merged_cell_dimensions(file_path, sheet_name):    
    # Load the Excel workbook
    wb = openpyxl.load_workbook(file_path)
    
    # Select the specified worksheet
    ws = wb[sheet_name]
    
    # Initialize an empty dictionary to store merged cell dimensions
    merged_cell_dimensions = {}
    
    # Iterate through all merged cells and get their dimensions
    for merged_cell_range in ws.merged_cells.ranges:
        # Get the coordinates of the merged cell
        min_col, min_row, max_col, max_row = merged_cell_range.min_col, merged_cell_range.min_row, merged_cell_range.max_col, merged_cell_range.max_row
        
        # Calculate the number of rows in the merged cell
        num_rows = max_row - min_row + 1
        
        # Store the dimensions in the dictionary
        merged_cell_dimensions[(min_row, min_col)] = num_rows
    
    # Close the workbook
    wb.close()
    
    print (merged_cell_dimensions)
    return merged_cell_dimensions

# Load the Excel workbook and sheets
file_path = './Train_Regression_Testcases.xlsx'
src_sheet_name = 'Train_Regression_TestCases'
unique_sheet_name = 'Unique names'
unique_column_name = 'Test_Case_Name(Unique Names)'
testcase_column_name = 'Test_Case_Name'

train_df = pd.read_excel(file_path, sheet_name=src_sheet_name)
unique_df = pd.read_excel(file_path, sheet_name=unique_sheet_name)

# Extract unique test case names
unique_test_cases = unique_df[unique_column_name].tolist()

# Initialize an empty DataFrame for the output
output_df = pd.DataFrame(columns=train_df.columns)

# Get merged cell dimensions
merged_cell_dimensions = get_merged_cell_dimensions(file_path, src_sheet_name)

# Iterate through the unique test cases and filter the rows
for test_case in unique_test_cases:
    filtered_rows = train_df[train_df[testcase_column_name] == test_case]
    
    # Exclude empty or all-NA columns from filtered_rows
    filtered_rows = filtered_rows.dropna(axis=1, how='all')
    
    # Get the number of rows for merged cells in specified columns
    for col in ['Test_Step_Description', 'Step_Name', 'Type', 'Status', 'Test_Step_Expected_Result', 'Assigned_To', 'TS_CREATION_DATE']:
        if col in filtered_rows.columns:
            col_index = filtered_rows.columns.get_loc(col)
            row_index = filtered_rows.index[0]  # Get the index of the first row
            num_rows = merged_cell_dimensions.get((row_index + 1, col_index + 1), 1)  # Adjust row and column indices
            print (num_rows)
            # Expand the rows for the merged cell
            filtered_rows.loc[:, col] = filtered_rows.iloc[0][col]
            filtered_rows = filtered_rows.iloc[:num_rows]
    
    output_df = pd.concat([output_df, filtered_rows], ignore_index=True)

# Append date-time stamp to the sheet name
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_sheet_name = f'output_{timestamp}'

# Write the output DataFrame to a new sheet with the unique sheet name
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    output_df.to_excel(writer, sheet_name=output_sheet_name, index=False)

print(f"Extraction completed and saved to '{output_sheet_name}' sheet.")
