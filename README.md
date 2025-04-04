# README.md

# Document Keyword Parser

A comprehensive document processing system that replaces keywords in Word documents with data from Excel spreadsheets, user inputs, JSON files, and templates. This system enables dynamic document generation with data-driven content and interactive user inputs, using `!` as the primary separator within keywords.

## Core Components

### Excel Manager (`excel_manager.py`)

The `excelManager` class provides a robust interface for interacting with Excel workbooks. It handles reading and writing data, including calculated values from formulas, totals, ranges, and specific columns based on titles or references.

*(No changes were made to excel_manager.py as requested. The description here reflects its usage by the updated keyword parser.)*

### Keyword Parser (`keyword_parser.py`)

The `keywordParser` class processes templating keywords (enclosed in `{{}}`) within text or Word documents. It replaces keywords with values from Excel, interactive user input, JSON files, and external templates. It now uses `!` as the separator between keyword components.

#### Architecture

The parser recognizes keywords like `{{TYPE!param1!param2...}}`. It integrates with `excelManager` for Excel data, generates Streamlit forms for user input, reads external files for templates and JSON data, and can directly insert formatted tables into Word documents.

#### Keyword Types (Using '!' Separator)

##### Excel Data Keywords (`{{XL!...}}`)

Keywords to fetch data from an Excel file.

| Keyword Pattern                                       | Description                                                                                     | Example                                                       |
| :---------------------------------------------------- | :---------------------------------------------------------------------------------------------- | :------------------------------------------------------------ |
| `{{XL!CELL!cell_ref}}`                                | Get value from a single cell (e.g., `A1`).                                                      | `{{XL!CELL!A1}}`                                              |
| `{{XL!CELL!SheetName!cell_ref}}`                      | Get value from a cell on a specific sheet.                                                      | `{{XL!CELL!Sales!C10}}`                                       |
| `{{XL!LAST!cell_ref}}`                                | Get the last non-empty value going down from `cell_ref`.                                        | `{{XL!LAST!F5}}`                                              |
| `{{XL!LAST!SheetName!cell_ref}}`                      | Get the last non-empty value from a specific sheet.                                             | `{{XL!LAST!Summary!B2}}`                                      |
| `{{XL!LAST!sheet_name!cell_ref!Title}}`               | Find column by `Title` in `cell_ref`'s row, get last value below.                               | `{{XL!LAST!Items!A4!Total Costs}}`                            |
| `{{XL!RANGE!range_ref}}`                              | Get values from a range (e.g., `A1:C5` or `NamedRange`). Returns formatted table.               | `{{XL!RANGE!A5:D10}}`                                         |
| `{{XL!RANGE!SheetName!range_ref}}`                    | Get values from a range on a specific sheet.                                                    | `{{XL!RANGE!Expenses!B2:G10}}`                                |
| `{{XL!COLUMN!sheet_name!col_refs}}`                   | Get specified columns by cell reference (e.g., "A1,C1,E1"). Returns table.                      | `{{XL!COLUMN!Items!A4,E4,F4}}`                                |
| `{{XL!COLUMN!sheet_name!"Titles"!start_row}}`         | Get columns by `Titles` (e.g., "Revenue,Profit") found in `start_row`. Returns table.           | `{{XL!COLUMN!Data!"Category,Value"!1}}`                       |

##### User Input Keywords (`{{INPUT!...}}`)

Keywords to create interactive input fields in Streamlit applications.

| Keyword Pattern                         | Description                                                   | Example                                                |
| :-------------------------------------- | :------------------------------------------------------------ | :----------------------------------------------------- |
| `{{INPUT!text!label!default_value}}`    | Single-line text input.                                       | `{{INPUT!text!Your Name!John Doe}}`                    |
| `{{INPUT!area!label!default_value!height}}` | Multi-line text area (optional height).                       | `{{INPUT!area!Comments::200}}`                         |
| `{{INPUT!date!label!default_date!format}}` | Date picker ('today' or 'YYYY/MM/DD', optional format).       | `{{INPUT!date!Select Date!today!YYYY/MM/DD}}`          |
| `{{INPUT!select!label!opt1,opt2,...}}`  | Dropdown selection.                                           | `{{INPUT!select!Choose Color!Red,Green,Blue}}`         |
| `{{INPUT!check!label!default_state}}`   | Checkbox ('True' or 'False').                                 | `{{INPUT!check!Agree to Terms!false}}`                 |

##### Template Keywords (`{{TEMPLATE!...}}`)

Keywords to include content from other files or libraries.

| Keyword Pattern                                    | Description                                                               | Example                                                     |
| :------------------------------------------------- | :------------------------------------------------------------------------ | :---------------------------------------------------------- |
| `{{TEMPLATE!filename.docx}}`                       | Include entire external template file.                                    | `{{TEMPLATE!disclaimer.txt}}`                               |
| `{{TEMPLATE!filename.docx!section=name}}`          | Include specific section/bookmark (implementation specific).              | `{{TEMPLATE!report.docx!section=conclusion}}`               |
| `{{TEMPLATE!filename.txt!line=5}}`                 | Include specific line number from text file.                              | `{{TEMPLATE!data.txt!line=10}}`                             |
| `{{TEMPLATE!filename.docx!paragraph=3}}`           | Include specific paragraph number (based on "\n\n" separation).         | `{{TEMPLATE!letter.txt!paragraph=2}}`                       |
| `{{TEMPLATE!filename.docx!VARS(key=val,...}})`     | Template with variable substitution (values can be keywords).             | `{{TEMPLATE!invite.txt!VARS(name=Jane,event=Party)}}`       |
| `{{TEMPLATE!LIBRARY!template_name!version}}`       | Reference template from a predefined library (optional version).          | `{{TEMPLATE!LIBRARY!confidentiality_clause}}`               |

##### JSON Data Keywords (`{{JSON!...}}`)

Keywords to fetch data from JSON files using JSONPath.

| Keyword Pattern                                    | Description                                                               | Example                                                       |
| :------------------------------------------------- | :------------------------------------------------------------------------ | :------------------------------------------------------------ |
| `{{JSON!filename.json!json_path}}`                 | Access data using JSONPath (e.g., `$.key`, `$.array[0].name`).            | `{{JSON!config.json!$.settings.theme}}`                       |
| `{{JSON!filename.json!json_path!TRANSFORMATION}}`  | Apply optional transformation: `SUM`, `JOIN(delimiter)`, `BOOL(Yes/No)`. | `{{JSON!sales.json!$.quarterly!SUM}}`                         |
|                                                    |                                                                           | `{{JSON!users.json!$.names!JOIN(, )}}`                        |
|                                                    |                                                                           | `{{JSON!settings.json!$.active!BOOL(Enabled/Disabled)}}`      |


#### Word Document Integration

The parser integrates with `python-docx` for enhanced Word document processing:
- Generates well-formatted tables directly in the document from `XL!RANGE!` and `XL!COLUMN!` keywords.
- Applies styling like headers and alternating row colors to inserted tables.

#### Usage Example

```python
from keyword_parser import keywordParser
from excel_manager import excelManager

# Initialize with an Excel manager
excel_mgr = excelManager("financial_report.xlsx")
parser = keywordParser(excel_mgr)

# Basic keyword replacement using '!'
template = "The total revenue is {{XL!CELL!Summary!B5}}."
result = parser.parse(template)
print(result) # Example: "The total revenue is $127,350.00."

# Parse template with table from columns
report_template = """
# Financial Summary
{{XL!COLUMN!Items!"Activities,Total Project Costs"!4}}

Contact: {{INPUT!text!Contact Name:}}
"""
# In a Streamlit app, this would generate a form first
processed_report = parser.parse(report_template)