# Excelify

Converts CSV files to formatted Excel files.

## Features

- Converts CSV files to Excel (.xlsx) format
- Intelligently formats column headers (title case with configurable uppercase words)
- Freezes the top row for easier navigation
- Adds filtering capability to all columns
- Auto-adjusts column widths based on content
- Left-aligns all cell content for better readability
- Removes default cell borders for a cleaner look

## Installation

1. Clone this repository or download the source code
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Configuration

The application uses a `config.json` file in the same directory as the script to specify which words should remain uppercase in column headers. The configuration file should have the following format:

```json
{
  "uppercase_words": [
    "API",
    "CSS",
    "DL",
    "DOB",
    "HTML",
    "HTTP",
    "HTTPS",
    "ID",
    "JSON",
    "SQL",
    "SSN",
    "URL",
    "XML"
  ]
}
```

You can modify this list to include any words that should remain uppercase in your column headers.

If the configuration file is missing or invalid, the application will use an empty list for uppercase words and display a warning message. This means all words in column headers will be converted to title case.

## Usage

Run the script from the command line, providing the path to your CSV file:

```bash
python excelify.py path/to/your/file.csv
```

The script will generate an Excel file with the same name in the same location as your CSV file.

## Example

If you have a file named `data.csv`, running:

```bash
python excelify.py data.csv
```

Will create a formatted Excel file named `data.xlsx` in the same directory.

## Requirements

- Python 3.6 or higher
- pandas
- openpyxl

## Error Handling

The script includes error handling for:

- Missing command-line arguments
- Non-CSV input files
- File reading/writing errors
- Missing or invalid configuration file (uses an empty list for uppercase words)

## License

The MIT License

Copyright 2025 Travis Horn

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
