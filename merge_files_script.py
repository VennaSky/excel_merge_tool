import pandas as pd
import os

def combine_excel_files_to_sheets():
    print("Enter the full paths of Excel files to combine into one workbook.")
    print("Each file will become its own sheet. Type 'done' when you're finished.")

    file_paths = []
    while True:
        path = input("Excel file path: ").strip()
        if path.lower() == 'done':
            break
        if os.path.exists(path) and (path.endswith('.xlsx') or path.endswith('.xls')):
            file_paths.append(path)
        else:
            print("File does not exist or is not an Excel file. Try again.")

    if not file_paths:
        print("No valid Excel files provided.")
        return

    output_path = input("Enter output filename (e.g., combined_sheets.xlsx): ").strip()
    if not output_path.endswith(".xlsx"):
        output_path += ".xlsx"

    with pd.ExcelWriter(output_path) as writer:
        for path in file_paths:
            sheet_name = os.path.splitext(os.path.basename(path))[0][:31]  # Excel sheet names max 31 chars
            print(f"Adding '{sheet_name}' to workbook...")
            df = pd.read_excel(path)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\n✅ Combined workbook saved as: {output_path}")

combine_excel_files_to_sheets()
