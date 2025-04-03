# Document Keyword Parser

A comprehensive document processing system that replaces keywords in Word documents with data from Excel spreadsheets and other sources. This system enables dynamic document generation with data-driven content and interactive user inputs.

## Core Components

### Excel Manager (`excel_manager.py`)

The `excelManager` class provides a robust interface for interacting with Excel workbooks, handling all aspects of workbook manipulation with special attention to formula preservation and consistent formatting.

#### Architecture

The class maintains two parallel workbook objects:
- **Data-only workbook**: For reading calculated values from formulas
- **Formula workbook**: For preserving formulas and formatting when writing

This dual approach ensures you get accurate data while maintaining the original structure and formatting.

#### Key Methods

| Method | Description | Parameters | Return Value |
|--------|-------------|------------|--------------|
| `__init__(file_path=None)` | Constructor | Optional path to Excel file | None |
| `create_workbook(file_path=None)` | Creates a new workbook | Optional file path | Workbook object |
| `load_workbook(file_path=None)` | Loads an existing workbook | Optional file path | Workbook object |
| `save(file_path=None)` | Saves the workbook | Optional file path | None |
| `close()` | Closes the workbook | None | None |
| `count_sheets()` | Counts sheets | None | Integer |
| `get_sheet_names()` | Gets all sheet names | None | List of strings |
| `create_sheet(sheet_name)` | Creates a new sheet | Sheet name | Sheet object |
| `get_sheet(sheet_name)` | Gets a sheet | Sheet name | Sheet object |
| `delete_sheet(sheet_name)` | Deletes a sheet | Sheet name | None |
| `read_cell(sheet_name, row_or_cell, column=None)` | Reads a cell | Sheet name, cell reference or row/column | Cell value (formatted) |
| `write_cell(sheet_name, row_or_cell, column=None, value=None)` | Writes to a cell | Sheet name, cell reference or row/column, value | None |
| `read_range(sheet_name, start_cell_or_row, ...)` | Reads a range | Sheet name, range reference or coordinates | 2D list of values |
| `write_range(sheet_name, start_cell_or_row, ...)` | Writes to a range | Sheet name, start cell, data | None |
| `read_total(sheet_name, row_or_cell, column=None)` | Finds total value | Sheet name, start position | Value |
| `read_items(sheet_name, row_or_cell, column=None, offset=0)` | Reads consecutive items | Sheet name, start position, offset | List of values |

#### Special Features

##### `read_total()` Method

This method traverses down a column starting from a specified cell, looking for the last non-empty value before an empty cell. Particularly useful for finding totals at the end of financial data columns.

```python
# Find total in column F starting from row 25
total = manager.read_total("Sheet1", "F25")

# Using row and column numbers instead of cell reference
total = manager.read_total("Sheet1", 25, 6)
```

##### `read_items()` Method

Collects all consecutive non-empty values in a column starting from a given cell until it encounters an empty cell. The offset parameter allows excluding a specified number of rows from the end (useful for omitting summary rows).

```python
# Get all items in column A starting from row 1
items = manager.read_items("Sheet1", "A1")

# Get items excluding the last row (typically a total)
items = manager.read_items("Sheet1", "A1", offset=1)
```

##### Formatting Support

The class automatically handles various formatting considerations:
- Preserves currency symbols and formatting
- Applies proper rounding to numeric values
- Maintains formula references and calculated values

#### Usage Examples

```python
from excel_manager import excelManager

# Initialize and load a workbook
manager = excelManager("financial_data.xlsx")

# Get all sheet names
sheets = manager.get_sheet_names()
print(f"Available sheets: {sheets}")

# Read a range and its total
monthly_data = manager.read_range("Income", "B5:B16")
annual_total = manager.read_total("Income", "B17")
print(f"Monthly data: {monthly_data}")
print(f"Annual total: {annual_total}")

# Write new values
manager.write_cell("Income", "C5", 1500.75)
manager.write_range("Income", "D5", [[100], [200], [300]])

# Save changes
manager.save()
manager.close()
```

### Keyword Parser (`keyword_parser.py`)

The `keywordParser` class processes various templating keywords in text, replacing them with values from Excel, user input, JSON files, and more. It serves as the core templating engine for the document processor.

#### Architecture

The parser recognizes keywords in double curly braces `{{keyword}}` and processes them based on their type. It can handle nested references and has special support for interactive user inputs and table generation.

#### Keyword Types

##### Excel Data Keywords

| Keyword Pattern | Description | Example |
|-----------------|-------------|---------|
| `{{XL:A1}}` | Basic cell reference | `{{XL:A1}}` |
| `{{XL:Sheet2!B5}}` | Cell with sheet name | `{{XL:Sales!C10}}` |
| `{{XL::A1}}` | Find total value (traverses down) | `{{XL::F25}}` |
| `{{XL:Sales!C3:C7}}` | Range of cells (formats as table) | `{{XL:Expenses!A5:D10}}` |

##### User Input Keywords

| Keyword Pattern | Description | Example |
|-----------------|-------------|---------|
| `{{INPUT:text:label:value}}` | Text input field | `{{INPUT:text:Your Name:John Doe}}` |
| `{{INPUT:area:label:value:height}}` | Text area (multi-line) | `{{INPUT:area:Comments::200}}` |
| `{{INPUT:date:label:value:format}}` | Date picker | `{{INPUT:date:Select Date:today:YYYY/MM/DD}}` |
| `{{INPUT:select:label:options}}` | Dropdown selection | `{{INPUT:select:Choose:Red,Green,Blue}}` |
| `{{INPUT:check:label:value}}` | Checkbox | `{{INPUT:check:Agree:false}}` |

##### Range Processing Keywords

| Keyword Pattern | Description | Example |
|-----------------|-------------|---------|
| `{{SUM:XL:A1:A10}}` | Sum of range | `{{SUM:XL:Sales!B5:B15}}` |
| `{{AVG:XL:Sheet1!B1:B10}}` | Average of range | `{{AVG:XL:Metrics!C1:C20}}` |

##### Template Keywords

| Keyword Pattern | Description | Example |
|-----------------|-------------|---------|
| `{{TEMPLATE:filename.docx}}` | Include entire template | `{{TEMPLATE:disclaimer.txt}}` |
| `{{TEMPLATE:file.docx:section=name}}` | Include specific section | `{{TEMPLATE:report.docx:section=conclusion}}` |
| `{{TEMPLATE:file.txt:line=5}}` | Include specific line | `{{TEMPLATE:data.txt:line=10}}` |
| `{{TEMPLATE:file.docx:VARS(name=X,date=Y)}}` | Template with variables | `{{TEMPLATE:letter.txt:VARS(name=John,date=2025-04-01)}}` |

##### JSON Data Keywords

| Keyword Pattern | Description | Example |
|-----------------|-------------|---------|
| `{{JSON:filename.json:$.key}}` | Access property | `{{JSON:config.json:$.settings.theme}}` |
| `{{JSON:data.json:$.users[0].name}}` | Access array element | `{{JSON:employees.json:$.staff[2].department}}` |
| `{{JSON:data.json:$.values:SUM}}` | Sum array values | `{{JSON:sales.json:$.quarterly:SUM}}` |
| `{{JSON:data.json:$.names:JOIN(,)}}` | Join array items | `{{JSON:teams.json:$.members:JOIN(, )}}` |

#### Word Document Integration

The class can integrate directly with python-docx:
- Generates well-formatted tables from Excel ranges
- Applies proper styling (headers, alternating rows, etc.)
- Handles table insertion within the document flow

#### Usage Example

```python
from keyword_parser import keywordParser
from excel_manager import excelManager

# Initialize with an Excel manager
excel_mgr = excelManager("financial_report.xlsx")
parser = keywordParser(excel_mgr)

# Basic keyword replacement
template = "The total revenue for Q1 is {{XL:Summary!B5}}."
result = parser.parse(template)
print(result)  # "The total revenue for Q1 is $127,350.00."

# Complex example with table and calculations
report_template = """
# Financial Summary
{{XL:Summary!A1:D5}}

The average monthly expense was {{AVG:XL:Expenses!B2:B13}}.
"""
processed_report = parser.parse(report_template)
```

## Applications

### Main Document Processor (`main.py`)

A Streamlit web application for processing Word documents with keyword replacements. This is the primary user-facing application of the system.

#### Features

- **Document Upload**: Support for Word documents (.docx) and Excel spreadsheets (.xlsx)
- **Interactive Form Generation**: Automatically detects `INPUT` keywords and creates forms
- **Excel Data Integration**: All Excel-related keywords are processed using the Excel manager
- **Table Generation**: Creates properly formatted tables from Excel ranges
- **Progress Tracking**: Visual progress bar and status updates during processing
- **Error Handling**: Robust error handling for various failure scenarios
- **Download**: Processed documents available for download

#### Processing Flow

1. User uploads Word document and Excel spreadsheet
2. System scans the document for all keywords
3. If INPUT keywords are found, a form is generated for user input
4. The system processes all keywords, replacing them with appropriate values
5. Special handling is applied for table generation from Excel ranges
6. The processed document is saved and made available for download

#### Implementation Details

- Uses `streamlit` for the web interface
- Employs `python-docx` for Word document manipulation
- Handles both paragraph and table cell replacements
- Supports direct table insertion into the document
- Processes keywords in a specific order to handle dependencies
- Preserves formatting of the original document

#### Usage

```bash
streamlit run main.py
```

### Tester Application (`tester_app.py`)

A development and testing UI for exploring the Excel Manager and Keyword Parser capabilities.

#### Features

- **File Management**: Create, load, save, and download Excel files
- **Excel Operations Tab**: Test various Excel operations
  - Count/list sheets
  - Create new sheets
  - Read cells and ranges
  - Find totals and item sequences
- **Sheet Management Tab**: Add/delete sheets
- **Keyword Testing Tab**: Test keyword parsing with live preview
  - Visual keyword parsing results
  - Input form generation testing
  - Clear input cache functionality

#### Implementation Details

- Provides a user-friendly interface for testing all class methods
- Displays tabular data using pandas DataFrames for better visualization
- Implements tabs for organizing different functional areas
- Supports transient file operations with tempfile for clean testing

#### Usage

```bash
streamlit run tester_app.py
```

## Installation and Setup

### Prerequisites

- Python 3.6 or higher
- pip package manager

### Dependencies

The following Python packages are required:

```
streamlit       # Web interface framework
openpyxl        # Excel file handling
pandas          # Data manipulation and display
python-docx     # Word document processing
```

### Installation Steps

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/document-keyword-parser.git
   cd document-keyword-parser
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   streamlit run main.py
   ```

## Use Cases

### Financial Reporting

Generate customized financial reports by pulling data from Excel spreadsheets:
- Insert current financial data into standardized templates
- Create dynamic tables of financial metrics
- Calculate summaries and averages on-the-fly

### Contract Generation

Create personalized legal documents:
- Insert client-specific information
- Include appropriate clauses based on conditions
- Generate professional-looking agreements with consistent formatting

### Data-Driven Communications

Produce customized communications:
- Insert recipient-specific information into form letters
- Include relevant data tables based on recipient category
- Allow for user input to personalize messages

## Best Practices

1. **Structuring Excel Data**:
   - Use consistent layouts with clear labels
   - Place totals at the end of data columns
   - Use descriptive sheet names

2. **Document Templates**:
   - Use clear, descriptive keywords
   - Group related data in tables
   - Test templates incrementally

3. **Performance Considerations**:
   - Large Excel files may slow processing
   - Complex documents with many keywords take longer to process
   - Consider breaking very large documents into smaller sections