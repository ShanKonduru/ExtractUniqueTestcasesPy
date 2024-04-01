import pandas as pd

# Load the Excel workbook and sheets
file_path = './Train_Regression_Testcases.xlsx'
src_sheet_name='Train_Regression_TestCases'
unique_sheet_name='Unique names'
unique_column_name = 'Test_Case_Name(Unique Names)'
testcase_column_name = 'Test_Case_Name'

train_df = pd.read_excel(file_path, sheet_name=src_sheet_name)
unique_df = pd.read_excel(file_path, sheet_name=unique_sheet_name)

# Extract unique test case names
unique_test_cases = unique_df[unique_column_name].tolist()

# Initialize an empty DataFrame for the output
output_df = pd.DataFrame(columns=[
    'Test_ID', 'Test_Plan_Folder', 'Test_Case_Name', 'Description', 
    'Test_Step_Description', 'Step_Name', 'Type', 'Status', 
    'Test_Step_Expected_Result', 'Assigned_To', 'TS_CREATION_DATE', 
    'Priority (P1- High priority, P2, P3, P4-Low Priority)', 
    'Interfacing application', 'Smoke Testing', 'Mock0', 'SIT0', 
    'SIT1', 'SIT2', 'Priority'
])

# Iterate through unique test cases and extract corresponding rows from train_df
for test_case in unique_test_cases:
    extracted_rows = train_df[train_df[testcase_column_name] == test_case]
    output_df = pd.concat([output_df, extracted_rows], ignore_index=True)

# Write the output DataFrame to a new sheet called 'output'
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    output_df.to_excel(writer, sheet_name='output', index=False)

print("Extraction completed and saved to 'output' sheet.")
