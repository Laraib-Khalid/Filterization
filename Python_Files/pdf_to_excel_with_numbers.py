import pdfplumber
import pandas as pd

def pdf_to_excel(pdf_path, excel_path, start_page, end_page):
    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []
        # Adjust loop to only process pages in the specified range
        for i in range(start_page - 1, end_page):  # pdfplumber pages are zero-indexed
            page = pdf.pages[i]
            tables = page.extract_tables()
            for table in tables:
                if table:
                    all_tables.append(table)

    # Debugging: Print the raw extracted tables
    print("Extracted Tables:")
    for table in all_tables:
        print(table)

    # Define the unwanted row pattern as a set of possible values to be removed
    unwanted_patterns = [
        "Terumo Product Code", "Description", "Price Rule", "EA/BX",
        "List Price (per pc)", "List Price (per box)",
        "Net Price (per pc)", "Net Price (per box)"
    ]

    # Flatten all tables into a single list of rows
    combined_data = [row for table in all_tables for row in table]

    # Clean headers by removing newlines
    headers = [header.replace('\n', ' ') for header in combined_data[0]]

    # Filter out rows that contain any of the unwanted patterns
    filtered_data = [headers] + [
        row for row in combined_data[1:]
        if not any(pattern in row for pattern in unwanted_patterns)
    ]

    # Debugging: Print the filtered data
    print("Filtered Data:")
    for row in filtered_data:
        print(row)

    # Remove newlines and spaces from the "Terumo Product Code" column
    if filtered_data:
        terumo_col_index = headers.index("Terumo Product Code") if "Terumo Product Code" in headers else -1
        for row in filtered_data[1:]:
            if terumo_col_index != -1 and len(row) > terumo_col_index:
                row[terumo_col_index] = row[terumo_col_index].replace('\n', ' ').replace(' ', '')

        # Create a DataFrame
        df = pd.DataFrame(filtered_data[1:], columns=filtered_data[0])
        df.to_excel(excel_path, index=False)
    else:
        print("No valid tables found in the PDF.")

# Define file paths and page range
pdf_path = 'D:/compare/data/pdf_number/2268 - 0 - Terumo Medical Corporation - ORLANDO HEALTH - Outpatient.pdf'
excel_path = 'D:/compare/data/pdf_number/2268 - 0 - Terumo Medical Corporation - ORLANDO HEALTH - Outpatient_processed_with_numbers.xlsx'
start_page = 9
end_page = 20

# Call the function
pdf_to_excel(pdf_path, excel_path, start_page, end_page)
