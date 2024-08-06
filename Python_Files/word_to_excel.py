from docx import Document
import pandas as pd

# List of unwanted patterns
unwanted_patterns = [
    "Terumo Product Code", "Description", "Price Rule", "EA/BX",
    "List Price (per pc)", "List Price (per box)",
    "Net Price (per pc)", "Net Price (per box)"
]

def extract_table_from_word(doc_path):
    # Load the Word document
    doc = Document(doc_path)

    # Initialize variables
    data = []

    for table in doc.tables:
        # Extract table rows
        for row in table.rows:
            data.append([cell.text.strip() for cell in row.cells])  # Use strip() to clean up whitespace

    return data

def clean_data(data):
    if not data:
        return []

    # Identify header
    header = data[0]

    # Remove completely blank rows
    data = [row for row in data if any(cell.strip() for cell in row)]

    # Remove rows with unwanted patterns (excluding header)
    data = [row for row in data if row != header and not any(pattern in row for pattern in unwanted_patterns)]

    # Remove rows where all columns have the same value
    data = [row for row in data if len(set(row)) > 1]

    # Ensure the header remains unchanged
    data.insert(0, header)

    return data

def save_to_excel(data, output_path):
    if not data:
        print("No data found to save.")
        return

    # Clean the data
    cleaned_data = clean_data(data)

    # Print the cleaned data to check
    for row in cleaned_data:
        print(row)

    # Create a DataFrame from the cleaned data
    try:
        df = pd.DataFrame(cleaned_data[1:], columns=cleaned_data[0])  # Assume first row is header
        df.to_excel(output_path, index=False)
    except ValueError as e:
        print(f"Error creating DataFrame: {e}")

# File paths
word_file_path = 'D:/compare/data/word/Contract_27701.0.docx'
excel_file_path = 'D:/compare/data/word/Contract_27701.0_processed.xlsx'

# Extract table and save to Excel
table_data = extract_table_from_word(word_file_path)
save_to_excel(table_data, excel_file_path)
