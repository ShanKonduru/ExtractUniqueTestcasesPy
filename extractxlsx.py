import openpyxl

def extract_test_case_details(input_file, output_file, test_names):
    # Load the input Excel workbook and get the active sheet
    wb_in = openpyxl.load_workbook(input_file)
    ws_in = wb_in.active

    # Create a new Excel workbook and sheet for the output
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active

    # Copy header from input to output
    for col_num, cell in enumerate(ws_in[1], 1):
        ws_out.cell(row=1, column=col_num, value=cell.value)

    for test_name in test_names:
        # Initialize a flag and buffer to track whether to extract data
        extract = False
        buffer = []

        for row in ws_in.iter_rows(min_row=2, values_only=True):
            if row[1] == test_name:
                extract = True
                buffer.append(row)
            elif not row[0] and not row[1] and extract:
                buffer.append(row)
            elif extract and row[0] and row[1]:
                extract = False
                for buffered_row in buffer:
                    ws_out.append(buffered_row)
                buffer = []
                if row[1] == test_name:
                    extract = True
                    buffer.append(row)

        # Write remaining rows from the buffer for the last test case
        if buffer:
            for buffered_row in buffer:
                ws_out.append(buffered_row)
            buffer = []

    # Save the output Excel workbook
    wb_out.save(output_file)

if __name__ == '__main__':
    input_file = './SampleCSVFile.xlsx'
    output_file = './output_SampleCSVFile.xlsx'
    test_names = ['Name 6', 'Name 2', 'Name 5']  # List of test case names in random order
    
    extract_test_case_details(input_file, output_file, test_names)
    print(f"Test case details for {test_names} extracted and saved to {output_file}.")
