import pandas as pd
import openpyxl # needed for pandas to work properly with Excel files
import argparse
import os

def main(xlsx_file, txt_file):
    csv_file = "./temp.csv"
    report_file = "./report.txt"

    # Convert Excel file to CSV file
    excel_file = pd.read_excel(xlsx_file)
    excel_file.to_csv(csv_file, index = None, header = True)

    # Read the target comments
    comments = []
    with open(txt_file, "r") as f:
        comments = f.readlines()

    # Read CSV file and extract headers
    df = pd.read_csv(csv_file)
    headers = list(filter(lambda header: "taskResults Comments" in header, df.columns.tolist()))

    # Write row numbers into text file
    with open(report_file, "w") as f: 
        for index, row in df.iterrows():
            for header in headers:
                if row[header] in comments:
                    f.write(str(index + 2) + "\n")
                break
    
    # Remove CSV file
    os.remove(csv_file)

if __name__ == "__main__":
    # Create parser and add arguments
    parser = argparse.ArgumentParser(description = "Retrieve violating rows")
    parser.add_argument("xlsx_file", type = str, help = 'Path to Excel file')
    parser.add_argument("txt_file", type = str, help = "Path to text file with target comments")

    # Retrieve arguments
    args = parser.parse_args()
    xlsx_file = args.xlsx_file
    txt_file = args.txt_file

    # Execute extraction
    main(xlsx_file, txt_file)