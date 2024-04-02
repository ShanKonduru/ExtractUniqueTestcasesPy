import openpyxl

def get_merged_cell_dimensions(file_path, sheet_name, column_identifier):
    """
    Get the dimensions (number of rows and columns) of merged cells in a specified column in an Excel worksheet.
    
    Args:
    - file_path (str): Path to the Excel file.
    - sheet_name (str): Name of the worksheet.
    - column_identifier (str or int): Column Name (e.g., 'A', 'B', 'C') or Column Number (e.g., 1, 2, 3) of the cell.
    
    Returns:
    - merged_cell_dimensions_list (list): List of tuples containing the number of rows and columns in each merged cell.
    """
    
    # Load the Excel workbook
    wb = openpyxl.load_workbook(file_path)
    
    # Select the specified worksheet
    ws = wb[sheet_name]
    
    # Convert column identifier to column index if it's a column name
    if isinstance(column_identifier, str):
        col_index = openpyxl.utils.column_index_from_string(column_identifier)
    else:
        col_index = column_identifier
    
    # Initialize merged_cell_dimensions_list
    merged_cell_dimensions_list = []
    
    # Check each row in the specified column for merged cells
    for row in range(1, ws.max_row + 1):
        cell = ws.cell(row=row, column=col_index)
        if cell.coordinate in ws.merged_cells:
            # Get the dimensions of the merged cell
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    min_col, min_row, max_col, max_row = merged_range.min_col, merged_range.min_row, merged_range.max_col, merged_range.max_row
                    num_rows = max_row - min_row + 1
                    num_columns = max_col - min_col + 1
                    merged_cell_dimensions_list.append((num_rows, num_columns))
                    break
    
    # Close the workbook
    wb.close()
    
    return merged_cell_dimensions_list

# Input parameters
file_path = './SampleFile.xlsx'
sheet_name = 'TestData'
column_identifier = 'A'  # Can also be an integer (e.g., 1 for column A)

# Get merged cell dimensions
merged_cell_dimensions_list = get_merged_cell_dimensions(file_path, sheet_name, column_identifier)

# Print the list of merged cell dimensions
for idx, (num_rows, num_columns) in enumerate(merged_cell_dimensions_list, start=1):
    print(f"Merged cell {idx} spans {num_rows} rows and {num_columns} columns.")
