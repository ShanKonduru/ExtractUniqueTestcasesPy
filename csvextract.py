import csv

def extract_test_case_details(input_file, output_file, test_name):
    with open(input_file, 'r') as infile, open(output_file, 'w', newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        # Write header to the output CSV file
        header = next(reader)
        writer.writerow(header)

        # Initialize a flag and buffer to track whether to extract data
        extract = False
        buffer = []

        for row in reader:
            if row[1] == test_name:
                extract = True
                buffer.append(row)
            elif not row[0] and not row[1] and extract:
                buffer.append(row)
            elif extract and row[0] and row[1]:
                extract = False
                for buffered_row in buffer:
                    writer.writerow(buffered_row)
                buffer = []
                if row[1] == test_name:
                    extract = True
                    buffer.append(row)

if __name__ == '__main__':
    input_file = './SampleCSVFile.csv'
    output_file = './output.csv'
    test_name = 'Name 3'
    
    extract_test_case_details(input_file, output_file, test_name)
    print(f"Test case details for {test_name} extracted and saved to {output_file}.")
