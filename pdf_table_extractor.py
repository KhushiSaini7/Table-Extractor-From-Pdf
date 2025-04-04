import pdfplumber
import pandas as pd
import os
import sys

def extract_tables_from_pdf(pdf_path):
    """
    Extracts tables from a PDF file.
    It first attempts to extract tables using pdfplumber's extract_tables()
    function (good for bordered tables) and if not found, it uses a custom
    heuristic based on word positions.
    """
    extracted_tables = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            # Attempt to detect tables using pdfplumber's built-in table extraction.
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    if any(row for row in table if any(cell is not None for cell in row)):
                        extracted_tables.append(table)
            else:
                # Fallback: custom detection for borderless or irregular tables.
                text_objs = page.extract_words()
                rows = group_text_into_rows(text_objs)
                if rows:
                    table = build_table_from_rows(rows)
                    extracted_tables.append(table)
                else:
                    print(f"No table detected on page {page_num}.")
    
    return extracted_tables

def group_text_into_rows(text_objs, tolerance=3):
    """
    Groups text objects (with positional info) into rows based on a tolerance
    on their 'top' value. This is a simple heuristic to cluster words that 
    appear on the same horizontal line.
    """
    rows = {}
    for obj in text_objs:
        # Round the 'top' coordinate based on tolerance.
        y = round(obj['top'] / tolerance) * tolerance
        rows.setdefault(y, []).append(obj)
    
    # Create a list of rows sorted by y-coordinate.
    sorted_rows = []
    for y in sorted(rows.keys()):
        # Sort each row's words by the x coordinate (left to right).
        row = sorted(rows[y], key=lambda x: x['x0'])
        sorted_rows.append(row)
    return sorted_rows

def build_table_from_rows(rows):
    """
    Builds a table (list of lists) from rows of text objects.
    Each inner list represents a row of cell values.
    """
    table = []
    for row in rows:
        # Concatenate words that are close to each other if needed.
        row_data = []
        current_cell = ""
        last_x = None
        # Adjust the tolerance for column separation as needed.
        col_tolerance = 10  
        for word in row:
            # If the x-coordinate gap is larger than the tolerance,
            # assume a new cell is starting.
            if last_x is None or word['x0'] - last_x > col_tolerance:
                if current_cell:
                    row_data.append(current_cell.strip())
                current_cell = word['text'] + " "
            else:
                current_cell += word['text'] + " "
            last_x = word['x1']
        if current_cell:
            row_data.append(current_cell.strip())
        table.append(row_data)
    return table

def export_tables_to_excel(tables, excel_path):
    """
    Exports a list of tables to an Excel workbook.
    Each table is written to a separate sheet.
    """
    writer = pd.ExcelWriter(excel_path, engine='openpyxl')
    for idx, table in enumerate(tables):
        # Create a DataFrame from the table data.
        df = pd.DataFrame(table)
        sheet_name = f"Table_{idx+1}"
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
    writer.close()
    print(f"Tables successfully exported to {excel_path}")

def main():
    if len(sys.argv) < 2:
        print("Usage: python pdf_table_extractor.py <pdf_path>")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    
    if not os.path.exists(pdf_path):
        print(f"File {pdf_path} does not exist.")
        sys.exit(1)
    
    print(f"Processing PDF: {pdf_path}")
    tables = extract_tables_from_pdf(pdf_path)
    
    if not tables:
        print("No tables detected in the PDF.")
        sys.exit(0)
    
    # Create an output Excel file in the same directory as the PDF.
    output_file = os.path.splitext(pdf_path)[0] + "_tables.xlsx"
    export_tables_to_excel(tables, output_file)

if __name__ == "__main__":
    main()
