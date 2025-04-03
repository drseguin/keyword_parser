import re
import json
import os
import streamlit as st
from datetime import datetime
from excel_manager import excelManager
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

class keywordParser:
    """
    A parser class that processes various keywords and extracts data from Excel,
    handles user input, and processes templates and JSON data.
    """
    
    def __init__(self, excel_manager=None):
        """
        Initialize the keyword parser.
        
        Args:
            excel_manager: An instance of excelManager to use for Excel operations.
                           If None, a new instance will be created when needed.
        """
        self.excel_manager = excel_manager
        self.pattern = r'{{(.*?)}}'
        self.has_input_fields = False
        self.form_submitted = False
        self.word_document = None
        
    def set_word_document(self, doc):
        """Set the word document for direct table insertion."""
        self.word_document = doc
        
    def parse(self, input_string):
        """
        Parse input string and process any keywords found.
        
        Args:
            input_string: The string containing keywords to parse.
            
        Returns:
            Processed string with keywords replaced with their values.
        """
        if not input_string:
            return input_string
            
        # Find all keywords in the input string
        matches = re.finditer(self.pattern, input_string)
        self.has_input_fields = False
        
        # First scan to check if there are any INPUT keywords
        for match in matches:
            content = match.group(1)  # Just the content inside {{}}
            parts = content.split(":", 1)
            keyword_type = parts[0].strip().upper()
            
            if keyword_type == "INPUT":
                self.has_input_fields = True
                break
        
        # If we have input fields, use a form
        if self.has_input_fields:
            with st.form(key="input_form"):
                # First pass - create all the UI elements but collect their values
                matches = re.finditer(self.pattern, input_string)
                input_values = {}
                
                for match in matches:
                    keyword = match.group(0)  # The full keyword with {{}}
                    content = match.group(1)  # Just the content inside {{}}
                    
                    # Check if it's an INPUT keyword
                    parts = content.split(":", 1)
                    keyword_type = parts[0].strip().upper()
                    
                    if keyword_type == "INPUT":
                        # Process each input field and store its value
                        value = self._create_input_field(content)
                        input_values[keyword] = value
                
                # Add the submit button
                submit_button = st.form_submit_button(label="Submit")
                if submit_button:
                    self.form_submitted = True
            
            if not self.form_submitted:
                # If the form hasn't been submitted yet, just return a placeholder message
                return "Please fill in all fields and click Submit."
            
            # After form submission, process all keywords with collected input values
            result = input_string
            for keyword, value in input_values.items():
                # Replace the keyword with its value in the result string
                result = result.replace(keyword, str(value) if value is not None else "")
            
            # Process any non-INPUT keywords
            matches = re.finditer(self.pattern, result)
            for match in matches:
                keyword = match.group(0)  # The full keyword with {{}}
                content = match.group(1)  # Just the content inside {{}}
                
                # Skip already processed INPUT keywords
                parts = content.split(":", 1)
                keyword_type = parts[0].strip().upper()
                if keyword_type != "INPUT":
                    # Process the keyword and get its value
                    replacement = self._process_keyword(content)
                    # Replace the keyword with its value in the result string
                    result = result.replace(keyword, str(replacement) if replacement is not None else "")
            
            return result
        else:
            # No input fields, process normally
            result = input_string
            matches = re.finditer(self.pattern, input_string)
            for match in matches:
                keyword = match.group(0)  # The full keyword with {{}}
                content = match.group(1)  # Just the content inside {{}}
                
                # Process the keyword and get its value
                replacement = self._process_keyword(content)
                
                # Replace the keyword with its value in the result string
                result = result.replace(keyword, str(replacement) if replacement is not None else "")
            
            return result
    
    def _create_input_field(self, content):
        """Create an input field and return its current value."""
        if not content:
            return "[Invalid input reference]"
            
        # Split the content by colon to check for input type
        parts = content.split(":", 1)
        field_name = parts[0].strip()
        
        # Create a unique key for this input field
        cache_key = f"INPUT:{content}"
        
        # Initialize session state if not already done
        if cache_key not in st.session_state:
            # Initialize based on input type
            if len(parts) > 1 and not parts[1].startswith(("select:", "date:")):
                # For text input with default value
                st.session_state[cache_key] = parts[1]
            else:
                st.session_state[cache_key] = ""

        # Handle different input types
        if len(parts) == 1:
            # Basic text input - don't use value parameter
            value = st.text_input(f"Enter value for {field_name}:", key=cache_key)
            return value

        input_type = parts[1].split(":", 1)[0].lower() if len(parts[1].split(":", 1)) > 1 else ""
        
        if input_type == "select":
            # Dropdown selection input
            options = parts[1].split(":", 1)[1].split(",") if len(parts[1].split(":", 1)) > 1 else []
            # Use index parameter instead of key/value combination
            index = 0
            if st.session_state[cache_key] in options:
                index = options.index(st.session_state[cache_key])
            value = st.selectbox(f"Select a value for {field_name}", options, index=index, key=cache_key)
            return value

        elif input_type == "date":
            # Date input
            date_format = parts[1].split(":", 1)[1] if len(parts[1].split(":", 1)) > 1 else "YYYY-MM-DD"
            
            # Initialize with current date
            current_date = datetime.now()
            value = st.date_input(f"Select a date for {field_name}", value=current_date, key=cache_key)
            
            # Format the date according to the specified format
            return value.strftime("%Y-%m-%d")  # Default format

        else:
            # Text input with default value - don't set value parameter since it's in session state
            value = st.text_input(f"Enter value for {field_name}", key=cache_key)
            return value
            
    def _process_keyword(self, content):
        """
        Process a single keyword content and return the corresponding value.
        
        Args:
            content: The content inside the {{ }} brackets.
            
        Returns:
            The processed value of the keyword.
        """
        parts = content.split(":", 1)
        keyword_type = parts[0].strip().upper()
        
        # Process Excel data keywords
        if keyword_type == "XL":
            return self._process_excel_keyword(parts[1] if len(parts) > 1 else "")
            
        # Process user input keywords - these should already be handled
        elif keyword_type == "INPUT":
            return "[Input field processed]"
            
        # Process range processing keywords
        elif keyword_type == "SUM":
            return self._process_sum_keyword(parts[1] if len(parts) > 1 else "")
            
        elif keyword_type == "AVG":
            return self._process_avg_keyword(parts[1] if len(parts) > 1 else "")
            
        # Process template keywords
        elif keyword_type == "TEMPLATE":
            return self._process_template_keyword(parts[1] if len(parts) > 1 else "")
            
        # Process JSON keywords
        elif keyword_type == "JSON":
            return self._process_json_keyword(parts[1] if len(parts) > 1 else "")
            
        # Unknown keyword type
        else:
            return f"[Unknown keyword type: {keyword_type}]"
    
    def _process_excel_keyword(self, content):
        """Process Excel-related keywords."""
        if not content:
            return "[Invalid Excel reference]"
            
        if not self.excel_manager:
            return "[Excel manager not initialized]"
        
        # Get available sheet names for case-insensitive comparison
        available_sheets = self.excel_manager.get_sheet_names()
        sheet_name_map = {sheet.lower(): sheet for sheet in available_sheets}
            
        # Check if the content starts with ":" which indicates using read_total method
        if content.startswith(":"):
            # This is the new syntax for read_total() {{XL::A1}}
            cell_ref = content[1:]  # Remove the leading colon
            
            try:
                # Handle sheet references like 'Sheet With Spaces'!A1
                if "!" in cell_ref:
                    parts = cell_ref.split("!")
                    sheet_name = parts[0].strip("'")  # Remove single quotes
                    cell_ref = parts[1]
                    
                    # Case-insensitive sheet name lookup
                    if sheet_name.lower() in sheet_name_map:
                        actual_sheet_name = sheet_name_map[sheet_name.lower()]
                        return self.excel_manager.read_total(actual_sheet_name, cell_ref)
                    else:
                        return f"[Sheet not found: {sheet_name}]"
                else:
                    # Use the active sheet
                    sheet_name = available_sheets[0]
                    return self.excel_manager.read_total(sheet_name, cell_ref)
            except Exception as e:
                return f"[Error reading total: {str(e)}]"
            
        # Handle standard syntax for cell and range references
        # Check if the content contains a range (e.g., A1:C3)
        elif ":" in content and "!" not in content.split(":")[1]:
            try:
                # Handle sheet references like 'Sheet With Spaces'!A1:C3
                if "!" in content:
                    parts = content.split("!")
                    sheet_name = parts[0].strip("'")  # Remove single quotes
                    cell_range = parts[1]
                    
                    # Case-insensitive sheet name lookup
                    if sheet_name.lower() in sheet_name_map:
                        actual_sheet_name = sheet_name_map[sheet_name.lower()]
                        data = self.excel_manager.read_range(actual_sheet_name, cell_range)
                    else:
                        return f"[Sheet not found: {sheet_name}]"
                else:
                    # Use the active sheet
                    sheet_name = available_sheets[0]
                    data = self.excel_manager.read_range(sheet_name, content)
                
                # Try to create a table if we have a document reference
                if self.word_document:
                    try:
                        return self._create_word_table(data)
                    except Exception as e:
                        # If table creation fails, fall back to text formatting
                        return self._format_table(data)
                else:
                    # No document reference, just format as text
                    return self._format_table(data)
                
            except Exception as e:
                return f"[Error reading range: {str(e)}]"
        else:
            try:
                # Handle sheet references like 'Sheet With Spaces'!A1
                if "!" in content:
                    parts = content.split("!")
                    sheet_name = parts[0].strip("'")  # Remove single quotes
                    cell_ref = parts[1]
                    
                    # Case-insensitive sheet name lookup
                    if sheet_name.lower() in sheet_name_map:
                        actual_sheet_name = sheet_name_map[sheet_name.lower()]
                        return self.excel_manager.read_cell(actual_sheet_name, cell_ref)
                    else:
                        return f"[Sheet not found: {sheet_name}]"
                else:
                    # Use the active sheet
                    sheet_name = available_sheets[0]
                    return self.excel_manager.read_cell(sheet_name, content)
            except Exception as e:
                return f"[Error reading cell: {str(e)}]"
    
    def _format_table(self, data):
        """
        Format the data as a formatted table for Word.
        
        Args:
            data: A 2D list of data from Excel.
            
        Returns:
            A formatted string representing the table.
        """
        if self.word_document:
            # If we have a Word document reference, create a table directly
            return self._create_word_table(data)
        
        # Otherwise create a text-based table
        if not data or not isinstance(data, list):
            return "No data"
            
        # Calculate column widths
        col_widths = []
        for row in data:
            for i, cell in enumerate(row):
                cell_str = str(cell)
                if i >= len(col_widths):
                    col_widths.append(len(cell_str))
                else:
                    col_widths[i] = max(col_widths[i], len(cell_str))
        
        # Create the table as a string
        result = []
        for row_index, row in enumerate(data):
            row_str = []
            for i, cell in enumerate(row):
                cell_str = str(cell)
                # Right-align numbers, left-align text
                if isinstance(cell, (int, float)) or (isinstance(cell, str) and cell.replace('.', '', 1).isdigit()):
                    formatted = cell_str.rjust(col_widths[i])
                else:
                    formatted = cell_str.ljust(col_widths[i])
                row_str.append(formatted)
            result.append(" | ".join(row_str))
            
            # Add a separator after the header row
            if row_index == 0:
                separator = []
                for i, width in enumerate(col_widths):
                    separator.append("-" * width)
                result.append("-+-".join(separator))
                
        return "\n".join(result)
        
    def _create_word_table(self, data):
        """
        Create a table directly in the Word document with manually added borders.
        
        Args:
            data: A 2D list of data from Excel.
            
        Returns:
            A placeholder text to replace the keyword.
        """
        if not data or not isinstance(data, list):
            return "No data"
            
        # Create a new table in the Word document
        num_rows = len(data)
        num_cols = max(len(row) for row in data)
        
        # Create the table without specifying any style
        table = self.word_document.add_table(rows=num_rows, cols=num_cols)
        
        # Add borders using direct XML modification
        try:
            from docx.oxml import parse_xml
            from docx.oxml.ns import nsdecls
            
            # Set border size (in twips)
            border_size = 4
            
            # Get the table XML element
            tbl = table._tbl
            
            # Find or create tblPr (table properties)
            tblPr = tbl.xpath('w:tblPr')
            if not tblPr:
                tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}></w:tblPr>')
                tbl.insert(0, tblPr)
            else:
                tblPr = tblPr[0]
            
            # Create borders element
            tblBorders = parse_xml(f'<w:tblBorders {nsdecls("w")}>' +
                f'<w:top w:val="single" w:sz="{border_size}" w:space="0" w:color="auto"/>' +
                f'<w:left w:val="single" w:sz="{border_size}" w:space="0" w:color="auto"/>' +
                f'<w:bottom w:val="single" w:sz="{border_size}" w:space="0" w:color="auto"/>' +
                f'<w:right w:val="single" w:sz="{border_size}" w:space="0" w:color="auto"/>' +
                f'<w:insideH w:val="single" w:sz="{border_size}" w:space="0" w:color="auto"/>' +
                f'<w:insideV w:val="single" w:sz="{border_size}" w:space="0" w:color="auto"/>' +
                '</w:tblBorders>')
            
            # Add or replace borders
            existing_borders = tblPr.xpath('w:tblBorders')
            if existing_borders:
                tblPr.remove(existing_borders[0])
            tblPr.append(tblBorders)
        except Exception as e:
            # If borders can't be added, continue without them
            pass
        
        # Fill the table with data
        for i, row in enumerate(data):
            for j, cell_value in enumerate(row):
                if j < len(row):  # Make sure we don't go out of bounds
                    # Format the cell value
                    if isinstance(cell_value, (int, float)):
                        cell_text = f"{cell_value:,}"
                    else:
                        cell_text = str(cell_value)
                    
                    cell = table.cell(i, j)
                    cell.text = cell_text
                    
                    # Format header row (first row)
                    if i == 0:
                        try:
                            for paragraph in cell.paragraphs:
                                # Try to make header bold
                                for run in paragraph.runs:
                                    run.bold = True
                        except:
                            pass  # Skip if formatting fails
                    
                    # Right-align numbers
                    try:
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        if isinstance(cell_value, (int, float)) or (isinstance(cell_value, str) and cell_value.replace('.', '', 1).isdigit()):
                            for paragraph in cell.paragraphs:
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    except:
                        pass  # Skip if alignment fails
        
        # Add a paragraph after the table for better spacing
        self.word_document.add_paragraph()
        
        # Return a placeholder - the actual table has been added to the document
        return "[TABLE_INSERTED]"
    
    def _process_sum_keyword(self, content):
        """Process SUM range processing keywords."""
        if not content:
            return "[Invalid SUM range]"
            
        try:
            # Handle Excel ranges
            if content.startswith("XL:"):
                range_ref = content[3:]
                
                # Get available sheet names for case-insensitive comparison
                available_sheets = self.excel_manager.get_sheet_names()
                sheet_name_map = {sheet.lower(): sheet for sheet in available_sheets}
                
                # Handle sheet references
                if "!" in range_ref:
                    parts = range_ref.split("!")
                    sheet_name = parts[0].strip("'")  # Remove single quotes
                    cell_range = parts[1]
                    
                    # Case-insensitive sheet name lookup
                    if sheet_name.lower() in sheet_name_map:
                        actual_sheet_name = sheet_name_map[sheet_name.lower()]
                    else:
                        return f"[Sheet not found: {sheet_name}]"
                else:
                    actual_sheet_name = available_sheets[0]
                    cell_range = range_ref
                
                # Read the range
                values = self.excel_manager.read_range(actual_sheet_name, cell_range)
                
                # Flatten the list and sum numeric values
                flat_values = [item for sublist in values for item in sublist]
                numeric_values = []
                
                for value in flat_values:
                    try:
                        # Handle currency formatting
                        if isinstance(value, str) and '$' in value:
                            value = float(value.replace('$', '').replace(',', ''))
                        numeric_values.append(float(value))
                    except (ValueError, TypeError):
                        pass
                        
                return sum(numeric_values)
            
            return "[Unsupported SUM reference]"
            
        except Exception as e:
            return f"[Error in SUM: {str(e)}]"
    
    def _process_avg_keyword(self, content):
        """Process AVG range processing keywords."""
        if not content:
            return "[Invalid AVG range]"
            
        try:
            # Handle Excel ranges
            if content.startswith("XL:"):
                range_ref = content[3:]
                
                # Get available sheet names for case-insensitive comparison
                available_sheets = self.excel_manager.get_sheet_names()
                sheet_name_map = {sheet.lower(): sheet for sheet in available_sheets}
                
                # Handle sheet references
                if "!" in range_ref:
                    parts = range_ref.split("!")
                    sheet_name = parts[0].strip("'")  # Remove single quotes
                    cell_range = parts[1]
                    
                    # Case-insensitive sheet name lookup
                    if sheet_name.lower() in sheet_name_map:
                        actual_sheet_name = sheet_name_map[sheet_name.lower()]
                    else:
                        return f"[Sheet not found: {sheet_name}]"
                else:
                    actual_sheet_name = available_sheets[0]
                    cell_range = range_ref
                
                # Read the range
                values = self.excel_manager.read_range(actual_sheet_name, cell_range)
                
                # Flatten the list and calculate average of numeric values
                flat_values = [item for sublist in values for item in sublist]
                numeric_values = []
                
                for value in flat_values:
                    try:
                        # Handle currency formatting
                        if isinstance(value, str) and '$' in value:
                            value = float(value.replace('$', '').replace(',', ''))
                        numeric_values.append(float(value))
                    except (ValueError, TypeError):
                        pass
                        
                if not numeric_values:
                    return 0
                    
                return sum(numeric_values) / len(numeric_values)
            
            return "[Unsupported AVG reference]"
            
        except Exception as e:
            return f"[Error in AVG: {str(e)}]"
    
    def _process_template_keyword(self, content):
        """Process template keywords."""
        if not content:
            return "[Invalid TEMPLATE reference]"
            
        try:
            # Split into filename and optional parameters
            parts = content.split(":", 1)
            filename = parts[0].strip()
            
            # Handle library templates
            if filename.upper() == "LIBRARY":
                if len(parts) > 1:
                    library_parts = parts[1].split(":", 1)
                    template_name = library_parts[0].strip()
                    template_version = library_parts[1].strip() if len(library_parts) > 1 else "DEFAULT"
                    
                    # This is where you would implement a template library lookup
                    return f"[Template Library: {template_name} (Version: {template_version})]"
                return "[Invalid library template reference]"
            
            # Handle file-based templates
            if not os.path.exists(filename):
                return f"[Template file not found: {filename}]"
                
            # Read the file
            with open(filename, 'r') as file:
                content = file.read()
                
            # Check for additional parameters
            if len(parts) > 1:
                param_part = parts[1]
                
                # Handle section/bookmark
                if "section=" in param_part:
                    section_name = param_part.split("section=")[1].split(",")[0].strip()
                    # This is where you would extract specific sections
                    return f"[Section {section_name} from {filename}]"
                
                # Handle specific line
                elif "line=" in param_part:
                    line_number = int(param_part.split("line=")[1].split(",")[0].strip())
                    lines = content.splitlines()
                    if 0 <= line_number < len(lines):
                        return lines[line_number]
                    return f"[Line {line_number} not found in {filename}]"
                
                # Handle specific paragraph
                elif "paragraph=" in param_part:
                    para_number = int(param_part.split("paragraph=")[1].split(",")[0].strip())
                    paragraphs = content.split("\n\n")
                    if 0 <= para_number < len(paragraphs):
                        return paragraphs[para_number]
                    return f"[Paragraph {para_number} not found in {filename}]"
                
                # Handle variable substitution
                elif "VARS(" in param_part:
                    vars_text = param_part.split("VARS(")[1].split(")")[0]
                    var_pairs = vars_text.split(",")
                    
                    # Create a dictionary of variables
                    variables = {}
                    for pair in var_pairs:
                        if "=" in pair:
                            key, value = pair.split("=", 1)
                            variables[key.strip()] = value.strip()
                    
                    # Replace variables in the template
                    result = content
                    for key, value in variables.items():
                        result = result.replace(f"{{{key}}}", value)
                        
                    return result
            
            # Return the entire file content
            return content
            
        except Exception as e:
            return f"[Error in TEMPLATE: {str(e)}]"
    
    def _process_json_keyword(self, content):
        """Process JSON keywords."""
        if not content or ":" not in content:
            return "[Invalid JSON reference]"
            
        try:
            # Split into filename and path
            parts = content.split(":", 1)
            filename = parts[0].strip()
            
            # Check if filename is from another reference
            if filename.startswith("{{") and filename.endswith("}}"):
                # Recursively parse the reference
                filename = self.parse(filename)
            
            # Check if file exists
            if not os.path.exists(filename):
                return f"[JSON file not found: {filename}]"
                
            # Read the JSON file
            with open(filename, 'r') as file:
                json_data = json.load(file)
            
            if len(parts) == 1:
                # Return the entire JSON content
                return json_data
            
            # Parse the JSONPath
            json_path = parts[1].split(":", 1)[0].strip()
            
            # Simplistic JSONPath implementation
            # Only handles basic paths like $.key.subkey or $.array[0].key
            if json_path.startswith("$."):
                path_parts = json_path[2:].split(".")
                current = json_data
                
                for part in path_parts:
                    # Handle array indexing
                    if "[" in part and part.endswith("]"):
                        key = part.split("[")[0]
                        index = int(part.split("[")[1][:-1])
                        
                        if key:
                            if key not in current:
                                return f"[JSON key not found: {key}]"
                            current = current[key][index]
                        else:
                            current = current[index]
                    else:
                        # Handle dynamic property names
                        if part.startswith("{{") and part.endswith("}}"):
                            part = self.parse(part)
                            
                        if part not in current:
                            return f"[JSON key not found: {part}]"
                        current = current[part]
                
                # Check for transformations
                if len(parts[1].split(":")) > 1:
                    transform_type = parts[1].split(":")[1].upper()
                    
                    if transform_type == "SUM" and isinstance(current, list):
                        try:
                            return sum(float(x) for x in current)
                        except (ValueError, TypeError):
                            return f"[Cannot sum non-numeric values]"
                    
                    elif transform_type.startswith("JOIN(") and transform_type.endswith(")"):
                        delimiter = transform_type[5:-1]
                        if isinstance(current, list):
                            return delimiter.join(str(x) for x in current)
                        return str(current)
                    
                    elif transform_type.startswith("BOOL(") and transform_type.endswith(")"):
                        yes_no = transform_type[5:-1].split("/")
                        yes_text = yes_no[0] if len(yes_no) > 0 else "Yes"
                        no_text = yes_no[1] if len(yes_no) > 1 else "No"
                        
                        return yes_text if current else no_text
                
                return current
            
            return f"[Invalid JSONPath: {json_path}]"
            
        except Exception as e:
            return f"[Error in JSON: {str(e)}]"
    
    def reset_form_state(self):
        """Reset the form submission state."""
        self.form_submitted = False
    
    def clear_input_cache(self):
        """Clear the cached user inputs."""
        # Find all INPUT keys in session state and clear them
        keys_to_clear = [key for key in st.session_state.keys() if key.startswith("INPUT:")]
        for key in keys_to_clear:
            st.session_state[key] = ""
        self.form_submitted = False

    def get_keyword_help(self):
        """
        Get help text explaining how to use keywords.
        
        Returns:
            A string with help information about available keywords.
        """
        help_text = """
        ## Keyword System Help

        ### Excel Data Keywords
        ```
        {{XL:A1}}             # Basic cell reference
        {{XL:Sheet2!B5}}      # Reference with sheet name
        {{XL::A1}}            # Find total value starting at A1 (traverses down to find last non-empty cell)
        {{XL::Sheet2!B5}}     # Find total value starting at B5 in Sheet2 (traverses down to find last non-empty cell)
        {{XL:Sales!C3:C7}}    # Range of cells (returns formatted table)
        {{XL:named_range}}    # Named range in Excel
        ```

        ### User Input Keywords
        ```
        {{INPUT:field_name}}          # Basic text input
        {{INPUT:field_name:default}}  # With default value
        {{INPUT:date:YYYY-MM-DD}}     # Date input with format
        {{INPUT:select:option1,option2,option3}}  # Dropdown selection
        ```

        ### Range Processing
        ```
        {{SUM:XL:A1:A10}}             # Sum of range
        {{AVG:XL:Sheet1!B1:B10}}      # Average of range
        ```

        ### Template Keywords
        ```
        {{TEMPLATE:filename.docx}}                 # Include entire external template
        {{TEMPLATE:filename.docx:section_name}}    # Include specific section/bookmark
        {{TEMPLATE:filename.txt:line=5}}           # Include specific line from text file
        {{TEMPLATE:filename.docx:paragraph=3}}     # Include specific paragraph
        {{TEMPLATE:filename.docx:VARS(name=John,date=2025-04-01)}}  # Template with variable substitution
        {{TEMPLATE:LIBRARY:legal_disclaimer}}      # Reference template from a predefined library
        ```

        ### JSON Data Keywords
        ```
        {{JSON:filename.json:$.key}}                  # Access top-level property
        {{JSON:config.json:$.settings.theme}}         # Access nested property using JSONPath
        {{JSON:data.json:$.users[0].name}}            # Access array element
        {{JSON:products.json:$.items[*].price}}       # Access all prices (returns array)
        {{JSON:data.json:$.values:SUM}}               # Sum numeric values in array
        {{JSON:data.json:$.names:JOIN(,)}}            # Join array items with delimiter
        {{JSON:config.json:$.enabled:BOOL(Yes/No)}}   # Transform boolean to text
        ```
        """
        return help_text