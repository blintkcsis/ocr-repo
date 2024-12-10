# PDF Rate Sheet to Excel Converter

This tool converts PDF rate sheets into a structured Excel format. It handles multi-page PDFs with repeated headers and extracts metadata like validity dates and currency information.

## Installation

1. Install the required packages:
```bash
pip install pandas docling
```

## Usage

### Command Line Interface (CLI)

Basic usage with just the PDF path:
```bash
python ocr.py path/to/RatesFile.pdf
```

Specify a different airline name:
```bash
python ocr.py path/to/RatesFile.pdf --airline "Lufthansa"
```

Specify an output directory:
```bash
python ocr.py path/to/RatesFile.pdf --output-dir ./rates
```

Use all options:
```bash
python ocr.py path/to/RatesFile.pdf -a "Lufthansa" -o ./rates
```

The output filename will be automatically generated in the format:
`{Airline}_{valid-from}-{valid-until}.xlsx`

For example: `Turkish_011024-010325.xlsx`

### Using in Python/Jupyter

```python
from docling.document_converter import DocumentConverter
from ocr import create_excel_from_markdown

# Convert PDF to markdown
source = "./RatesFile.pdf"
converter = DocumentConverter()
result = converter.convert(source)
markdown_text = result.document.export_to_markdown()

# Convert markdown to Excel
airline_name = "Turkish"
output_dir = "./rates"
df, output_path = create_excel_from_markdown(markdown_text, airline_name, output_dir)
print(f"Excel file created successfully at: {output_path}")
```

## Output Format

The Excel file will contain the following columns:
- Airline
- Origin
- Destination
- Commodity
- Min
- <45
- >45
- >100
- >300
- >500
- >1000
- Currency
- Valid from
- Valid until
- Notes

## Features

- Handles multi-page PDFs with repeated headers
- Extracts metadata (dates, currency, commodity type)
- Converts all rates to proper numeric format
- Automatically generates filename based on metadata
- Supports both CLI and programmatic usage
- Handles malformed data and table structure variations

## Error Handling

The script includes error handling for:
- Invalid PDF files
- Malformed table data
- Missing metadata
- Invalid numeric values
- File permission issues

If any rows contain invalid data, they will be skipped and processing will continue with the next valid row.