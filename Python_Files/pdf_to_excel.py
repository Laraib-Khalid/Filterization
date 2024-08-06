import pdfplumber
import pandas as pd


def pdf_to_excel(pdf_path, excel_path):
    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:
                    all_tables.append(table)

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


# Define file paths
pdf_path = 'D:/compare/data/pdf/PriceExhibit_27690.pdf'
excel_path = 'D:/compare/data/pdf/PriceExhibit_27690_processed.xlsx'

# Call the function
pdf_to_excel(pdf_path, excel_path)
