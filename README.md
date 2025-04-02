# Excel Manager

A Python class for managing Excel files with various read, write, and sheet management operations.

## Overview

The `excelManager` class provides a powerful yet simple interface for interacting with Excel workbooks. It supports reading calculated values from formulas, maintaining formatting information, and handling various Excel operations.

## Features

- Create, load, and save Excel workbooks
- Manage sheets (create, delete, list)
- Read individual cells and ranges with calculated values
- Write to cells and ranges
- Find totals at the end of columns 
- Extract consecutive items from columns with offset capability
- Support for A1 notation and row/column indices
- Consistent error handling and logging
- Currency and numeric formatting support

## Usage

### Initialization

```python
from excel_manager import excelManager

# Initialize with a file path
manager = excelManager("path/to/file.xlsx")

# Initialize without a file path
manager = excelManager()
```

### Workbook Operations

```python
# Create a new workbook
manager.create_workbook("path/to/new_file.xlsx")

# Load an existing workbook
manager.load_workbook("path/to/file.xlsx")

# Save changes
manager.save()

# Close the workbook
manager.close()
```

### Sheet Operations

```python
# Get the number of sheets
count = manager.count_sheets()

# Get all sheet names
names = manager.get_sheet_names()

# Create a new sheet
manager.create_sheet("Sheet Name")

# Get a sheet by name
sheet = manager.get_sheet("Sheet Name")

# Delete a sheet
manager.delete_sheet("Sheet Name")
```

### Read Operations

```python
# Read a cell value (using cell reference)
value = manager.read_cell("Sheet1", "A1")

# Read a cell value (using row and column numbers)
value = manager.read_cell("Sheet1", 1, 1)

# Read a range (using range reference)
values = manager.read_range("Sheet1", "A1:C3")

# Read a range (using start and end cell references)
values = manager.read_range("Sheet1", "A1", "C3")

# Read a range (using row and column numbers)
values = manager.read_range("Sheet1", 1, 1, 3, 3)

# Find a total value by traversing down a column
total = manager.read_total("Sheet1", "A1")  # Using cell reference
total = manager.read_total("Sheet1", 1, 1)  # Using row and column numbers

# Read consecutive items from a column until an empty cell
items = manager.read_items("Sheet1", "A1")  # Using cell reference
items = manager.read_items("Sheet1", "A1", offset=1)  # Exclude the last item (total)
items = manager.read_items("Sheet1", 1, 1)  # Using row and column numbers
```

### Write Operations

```python
# Write a cell value (using cell reference)
manager.write_cell("Sheet1", "A1", "Value")

# Write a cell value (using row and column numbers)
manager.write_cell("Sheet1", 1, 1, "Value")

# Write a range (using cell reference)
data = [["A", "B", "C"], [1, 2, 3], [4, 5, 6]]
manager.write_range("Sheet1", "A1", data)

# Write a range (using row and column numbers)
manager.write_range("Sheet1", 1, 1, data)
```

## Special Features

### Read Total Method

The `read_total` method traverses down a column starting from a given cell, looking for the last non-empty value before an empty cell. This is particularly useful for finding totals at the end of data columns in financial spreadsheets or reports.

Method signature:
```python
def read_total(self, sheet_name, row_or_cell, column=None)
```

Parameters:
- `sheet_name`: The name of the sheet to read from
- `row_or_cell`: Either a cell reference string (e.g., "A1") or a row number
- `column`: Optional column number (required if row_or_cell is a row number)

Returns:
- The value of the last non-empty cell in the column before an empty cell
- `None` if no values are found

Usage examples:
```python
# Find total in column F starting from row 25
total = manager.read_total("Sheet1", "F25")

# Find total using row and column numbers
total = manager.read_total("Sheet1", 25, 6)
```

How it works:
1. Starts from the specified cell position
2. Traverses down the column checking each cell
3. When it encounters an empty cell, it returns the last non-empty value it found
4. If it reaches the end of the sheet, it returns the last value
5. Maintains formatting (e.g., currency symbols) and rounds floating point values

### Read Items Method

The `read_items` method collects all consecutive non-empty values in a column starting from a given cell until it encounters an empty cell. It also provides an option to exclude a specified number of rows from the end, which is useful when you want to retrieve line items without including summary totals.

Method signature:
```python
def read_items(self, sheet_name, row_or_cell, column=None, offset=0)
```

Parameters:
- `sheet_name`: The name of the sheet to read from
- `row_or_cell`: Either a cell reference string (e.g., "A1") or a row number
- `column`: Optional column number (required if row_or_cell is a row number)
- `offset`: Number of rows to exclude from the end of the found range (default 0)

Returns:
- A list of values from the starting cell until an empty cell is found
- The list will exclude the specified number of items from the end if offset is provided
- An empty list if no values are found

Usage examples:
```python
# Get all items in column A starting from row 1
items = manager.read_items("Sheet1", "A1")

# Get items in column A, excluding the last row (e.g., a total)
items = manager.read_items("Sheet1", "A1", offset=1)

# Get items using row and column numbers, excluding the last 2 rows
items = manager.read_items("Sheet1", 1, 1, offset=2)
```

How it works:
1. Starts from the specified cell position
2. Traverses down the column collecting all non-empty values
3. Stops when it encounters the first empty cell
4. Applies the offset by removing the specified number of items from the end of the collected list
5. Maintains formatting (e.g., currency symbols) and rounds floating point values
6. Returns the resulting list of values

## Implementation Details

The class maintains two copies of each workbook:
- A data-only version for calculated values
- A formula version for maintaining formulas and formatting

This dual approach ensures that both calculated values and original formulas are accessible.