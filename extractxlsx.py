import openpyxl

def extract_test_case_details(input_file, output_file):
    # Load the input Excel workbook
    wb_in = openpyxl.load_workbook(input_file)
    ws_test_cases = wb_in['Train_Regression_TestCases']
    ws_test_names = wb_in['UniqueTestcaseName']

    # Create a new Excel workbook and sheet for the output
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active

    # Copy header from input to output
    for col_num, cell in enumerate(ws_test_cases[1], 1):
        ws_out.cell(row=1, column=col_num, value=cell.value)

    # Read test case names from "UniqueTestcaseName" sheet, skipping the first row
    test_names = [cell.value for cell in ws_test_names['A'][1:] if cell.value]

    # Initialize a flag to check if a test case name is found
    found_test_case = False

    # Iterate through the "Train_Regression_TestCases" sheet to extract rows
    for row in ws_test_cases.iter_rows(min_row=2, values_only=True):
        if len(row) > 1 and row[1] in test_names:
            found_test_case = True
            ws_out.append(row)
        elif not row[0] and not row[1] and found_test_case:
            ws_out.append(row)
        elif row[0] and row[1]:
            found_test_case = False

    # Save the output Excel workbook
    wb_out.save(output_file)

if __name__ == '__main__':
    input_file = './SampleCSVFile.xlsx'
    output_file = './output_SampleCSVFile.xlsx'
    
    extract_test_case_details(input_file, output_file)
    print(f"Test case details extracted and saved to {output_file}.")
