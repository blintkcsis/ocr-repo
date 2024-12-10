import pandas as pd
import re
import argparse
from datetime import datetime
from docling.document_converter import DocumentConverter

def extract_metadata(markdown_text):
    """Extract metadata from the markdown notes section."""
    # Initialize default values
    currency = "EUR"
    valid_from = ""
    valid_until = ""
    commodity = "General Cargo"
    
    # Find dates using regex
    date_pattern = r'(\d{2}/\d{2}/\d{4})'
    dates = re.findall(date_pattern, markdown_text)
    if len(dates) >= 2:
        valid_from = dates[0]
        valid_until = dates[1]
    
    return {
        'currency': currency,
        'valid_from': valid_from,
        'valid_until': valid_until,
        'commodity': commodity
    }

def extract_table_data(markdown_text):
    """Extract the table data from markdown text."""
    lines = markdown_text.split('\n')
    table_data = []
    current_section = []
    in_table = False
    
    for line in lines:
        # Skip empty lines
        if not line.strip():
            continue
            
        # Check if this is a table row
        if '|' in line:
            # Skip separator lines (containing only dashes and pipes)
            if re.match(r'^[\|\s\-]+$', line):
                continue
            # Skip header lines
            if 'Origin' in line and 'Destination' in line:
                in_table = True
                continue
                
            if in_table:
                # Process the table row
                cells = [cell.strip() for cell in line.split('|')[1:-1]]
                if len(cells) >= 4:  # Ensure we have minimum required columns
                    try:
                        # Clean and validate the data
                        cells = [c.strip() for c in cells]
                        # Convert numeric values and validate
                        numeric_values = [
                            cells[0],  # Origin
                            cells[1],  # Destination
                            float(cells[2].replace(',', '.')),  # Min
                            float(cells[3].replace(',', '.')),  # <45
                            float(cells[4].replace(',', '.')),  # >45
                            float(cells[5].replace(',', '.')) if len(cells) > 5 else None  # >100
                        ]
                        table_data.append(numeric_values)
                    except (ValueError, IndexError):
                        continue  # Skip rows with invalid data
    
    return table_data

def format_date_for_filename(date_str):
    """Convert date from DD/MM/YYYY to DDMMYY format for filename"""
    if not date_str:
        return "unknown_date"
    try:
        # Parse the date string
        day, month, year = date_str.split('/')
        # Return in DDMMYY format
        return f"{day}{month}{year[-2:]}"
    except:
        return "unknown_date"

def create_excel_from_markdown(markdown_text, airline_name, output_dir="."):
    """Convert markdown text to Excel file with specified format."""
    # Extract metadata and table data
    metadata = extract_metadata(markdown_text)
    table_data = extract_table_data(markdown_text)
    
    # Format dates for filename
    valid_from = format_date_for_filename(metadata['valid_from'])
    valid_until = format_date_for_filename(metadata['valid_until'])
    
    # Create output filename
    output_filename = f"{airline_name}_{valid_from}-{valid_until}.xlsx"
    output_path = f"{output_dir}/{output_filename}"
    
    # Create DataFrame
    df_data = []
    for row in table_data:
        df_row = {
            'Airline': airline_name,
            'Origin': row[0],
            'Destination': row[1],
            'Commodity': metadata['commodity'],
            'Min': float(row[2]),
            '<45': float(row[3]),
            '>45': float(row[4]),
            '>100': float(row[5]) if len(row) > 5 else None,
            '>300': None,
            '>500': None,
            '>1000': None,
            'Currency': metadata['currency'],
            'Valid from': metadata['valid_from'],
            'Valid until': metadata['valid_until'],
            'Notes': ''
        }
        df_data.append(df_row)
    
    df = pd.DataFrame(df_data)
    
    # Write to Excel
    df.to_excel(output_path, index=False)
    return df, output_path

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Convert PDF rate sheet to Excel format')
    parser.add_argument('pdf_path', help='Path to the PDF file')
    parser.add_argument('--airline', '-a', default='Turkish', help='Airline name (default: Turkish)')
    parser.add_argument('--output-dir', '-o', default='.', help='Output directory (default: current directory)')
    
    # Parse arguments
    args = parser.parse_args()
    
    # Convert PDF to markdown
    converter = DocumentConverter()
    result = converter.convert(args.pdf_path)
    markdown_text = result.document.export_to_markdown()
    
    # Convert markdown to Excel
    df, output_path = create_excel_from_markdown(markdown_text, args.airline, args.output_dir)
    print(f"Excel file created successfully at: {output_path}")

if __name__ == "__main__":
    main()