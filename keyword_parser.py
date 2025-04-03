import re
import json
import os
import streamlit as st
from datetime import datetime
from excel_manager import excelManager
from docx.shared import Pt, Cm
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
        self.input_values = {}  # Store input values
        
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
        matches = list(re.finditer(self.pattern, input_string))
        
        # First handle all INPUT keywords
        input_keywords = []
        for match in matches:
            content = match.group(1)  # Content inside {{}}
            keyword = match.group(0)  # The full {{keyword}}
            parts = content.split(":", 1)
            keyword_type = parts[0].strip().upper()
            
            if keyword_type == "INPUT":
                input_keywords.append((keyword, content))
        
        # If we have input fields, process them first
        if input_keywords:
            with st.form(key=f"input_form_{id(input_string)}"):
                st.subheader("Please provide input values:")
                
                # Create input fields and store their values
                for keyword, content in input_keywords:
                    value = self._create_input_field(content)
                    self.input_values[keyword] = value
                
                # Add submit button
                submit = st.form_submit_button("Submit")
                if not submit:
                    return "Please fill in all fields and click Submit."
        
        # After processing inputs or if no inputs, process all keywords
        result = input_string
        for match in matches:
            keyword = match.group(0)  # Full keyword with {{}}
            content = match.group(1)  # Content inside {{}}
            
            # If this is an INPUT keyword we've already processed
            if keyword in self.input_values:
                replacement = self.input_values[keyword]
            else:
                replacement = self._process_keyword(content)
            
            # Replace the keyword with its value
            result = result.replace(keyword, str(replacement) if replacement is not None else "")
        
        return result
    
    def _create_input_field(self, content):
        """
        Create an appropriate input field based on the INPUT keyword.
        
        Args:
            content: The content inside the {{ }} brackets.
            
        Returns:
            The value from the input field.
        """
        if not content:
            return "[Invalid input reference]"
        
        # Split the content into tokens
        tokens = content.split(":")
        if len(tokens) < 2:
            return "[Invalid INPUT format]"
        
        # Get the keyword type (INPUT) and input type (text, area, date, select, check)
        keyword_type = tokens[0].strip().upper()
        input_type = tokens[1].strip().lower() if len(tokens) > 1 else ""
        
        # Check for valid INPUT keyword
        if keyword_type != "INPUT":
            return "[Invalid INPUT keyword]"
        
        # Handle text input - {{INPUT:text:label:value}}
        if input_type == "text":
            label = tokens[2] if len(tokens) > 2 else ""
            default_value = tokens[3] if len(tokens) > 3 else ""
            return st.text_input(
                label=label, 
                value=default_value, 
                label_visibility="visible"
            )
        
        # Handle text area - {{INPUT:area:label:value:height}}
        elif input_type == "area":
            label = tokens[2] if len(tokens) > 2 else ""
            default_value = tokens[3] if len(tokens) > 3 else ""
            height_px = tokens[4] if len(tokens) > 4 else None
            
            # Convert height to integer if provided
            height = None
            if height_px:
                try:
                    height = int(height_px)
                except ValueError:
                    # If height is not a valid integer, ignore it
                    pass
            
            # Set height if provided, otherwise use default
            if height:
                return st.text_area(
                    label=label, 
                    value=default_value, 
                    height=height,
                    label_visibility="visible"
                )
            else:
                return st.text_area(
                    label=label, 
                    value=default_value, 
                    label_visibility="visible"
                )
        
        # Handle date input - {{INPUT:date:label:value:format}}
        elif input_type == "date":
            label = tokens[2] if len(tokens) > 2 else ""
            default_value = tokens[3] if len(tokens) > 3 else "today"
            date_format = tokens[4] if len(tokens) > 4 else "YYYY/MM/DD"
            
            # Handle "today" default value
            import datetime
            if default_value.lower() == "today":
                default_date = datetime.date.today()
            else:
                try:
                    # Try to parse the date based on the format
                    if date_format == "YYYY/MM/DD":
                        default_date = datetime.datetime.strptime(default_value, "%Y/%m/%d").date()
                    elif date_format == "DD/MM/YYYY":
                        default_date = datetime.datetime.strptime(default_value, "%d/%m/%Y").date()
                    elif date_format == "MM/DD/YYYY":
                        default_date = datetime.datetime.strptime(default_value, "%m/%d/%Y").date()
                    else:
                        # Default to ISO format if format is not recognized
                        default_date = datetime.datetime.strptime(default_value, "%Y-%m-%d").date()
                except ValueError:
                    default_date = datetime.date.today()
            
            date_value = st.date_input(
                label=label,
                value=default_date,
                label_visibility="visible"
            )
            
            # Return the date in the requested format
            if date_format == "YYYY/MM/DD":
                return date_value.strftime("%Y/%m/%d")
            elif date_format == "DD/MM/YYYY":
                return date_value.strftime("%d/%m/%Y")
            elif date_format == "MM/DD/YYYY":
                return date_value.strftime("%m/%d/%Y")
            else:
                return date_value.strftime("%Y/%m/%d")  # Default format
        
        # Handle select box - {{INPUT:select:label:options}}
        elif input_type == "select":
            label = tokens[2] if len(tokens) > 2 else ""
            options_str = tokens[3] if len(tokens) > 3 else ""
            
            # Parse options (comma-separated)
            options = [opt.strip() for opt in options_str.split(",")] if options_str else []
            
            if not options:
                return "[No options provided]"
            
            return st.selectbox(
                label=label,
                options=options,
                label_visibility="visible"
            )
        
        # Handle checkbox - {{INPUT:check:label:value}}
        elif input_type == "check":
            label = tokens[2] if len(tokens) > 2 else ""
            default_value_str = tokens[3].lower() if len(tokens) > 3 else "false"
            
            # Convert string value to boolean
            default_value = default_value_str == "true"
            
            return st.checkbox(
                label=label,
                value=default_value,
                label_visibility="visible"
            )
        
        # Default for unrecognized input types
        else:
            return f"[Unsupported input type: {input_type}]"
    
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
            
        # Process user input keywords - these should already be handled in parse()
        elif keyword_type == "INPUT":
            params = parts[1] if len(parts) > 1 else ""
            return self._process_input_keyword(params)
            
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
    
    def _process_input_keyword(self, params):
        """Process INPUT keywords directly if needed."""
        # This is a fallback method in case an INPUT keyword wasn't processed in the form
        input_parts = params.split(":")
        input_type = input_parts[0].lower() if input_parts else ""
        
        if input_type == "text" or input_type == "area":
            label = input_parts[1] if len(input_parts) > 1 else ""
            default_value = input_parts[2] if len(input_parts) > 2 else ""
            return default_value
        
        elif input_type == "date":
            import datetime
            today = datetime.date.today()
            return today.strftime("%Y/%m/%d")
        
        elif input_type == "select":
            options_str = input_parts[2] if len(input_parts) > 2 else ""
            options = [opt.strip() for opt in options_str.split(",")] if options_str else []
            return options[0] if options else ""
        
        elif input_type == "check":
            default_value_str = input_parts[2].lower() if len(input_parts) > 2 else "false"
            return default_value_str == "true"
        
        else:
            return params if params else "[Input value]"
    
    # The rest of the methods remain unchanged
    def _process_excel_keyword(self, content):
        """Process Excel-related keywords with improved table handling."""
        if not content:
            return "[Invalid Excel reference]"
                
        if not self.excel_manager:
            return "[Excel manager not initialized]"
        
        # Get available sheet names for case-insensitive comparison
        available_sheets = self.excel_manager.get_sheet_names()
        sheet_name_map = {sheet.lower(): sheet for sheet in available_sheets}
                
        # Check if the content starts with ":" which indicates using read_total method
        if content.startswith(":"):
            # This is the syntax for read_total() {{XL::A1}}
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
            # This is a range reference
            try:
                # Handle sheet references like 'Sheet With Spaces'!A1:C3
                sheet_name = available_sheets[0]  # Default to first sheet
                cell_range = content
                
                if "!" in content:
                    parts = content.split("!")
                    sheet_ref = parts[0].strip("'")  # Remove single quotes
                    cell_range = parts[1]
                    
                    # Case-insensitive sheet name lookup
                    if sheet_ref.lower() in sheet_name_map:
                        sheet_name = sheet_name_map[sheet_ref.lower()]
                    else:
                        return f"[Sheet not found: {sheet_ref}]"
                
                # Read the range data
                data = self.excel_manager.read_range(sheet_name, cell_range)
                
                # Try to create a table if we have a document reference
                if self.word_document and data:
                    try:
                        # Always attempt to create a proper Word table
                        return self._create_word_table(data)
                    except Exception as e:
                        # Log the error and fall back to text formatting
                        print(f"Error creating Word table: {str(e)}")
                        return self._format_table(data)
                else:
                    # No document reference or no data, format as text
                    return self._format_table(data)
                    
            except Exception as e:
                return f"[Error reading range: {str(e)}]"
        else:
            # This is a single cell reference
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
        Create a visually appealing table directly in the Word document with proper styling.
        
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
        
        # Create the table with the 'Table Grid' style for consistent borders
        table = self.word_document.add_table(rows=num_rows, cols=num_cols)
        table.style = 'Table Grid'
        
        # Set overall table properties for better appearance
        try:
            from docx.shared import Pt, Cm
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            from docx.oxml import parse_xml
            from docx.oxml.ns import nsdecls
            
            # Set the table to auto-fit contents
            table.autofit = True
            
            # Set first row as header with distinct formatting
            header_row = True
            
            # Fill the table with data and apply formatting
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
                        
                        # Apply padding to all cells
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(10)  # Consistent font size
                            
                            # Add spacing within cells
                            paragraph.paragraph_format.space_before = Pt(3)
                            paragraph.paragraph_format.space_after = Pt(3)
                        
                        # Format header row (first row)
                        if i == 0 and header_row:
                            # Make header bold with light gray background
                            cell.paragraphs[0].runs[0].font.bold = True
                            
                            # Add light gray shading to header row
                            tcPr = cell._tc.get_or_add_tcPr()
                            shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9D9D9"/>')
                            tcPr.append(shading_elm)
                            
                            # Center align header text
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # Right-align numbers for better readability
                        elif isinstance(cell_value, (int, float)) or (
                                isinstance(cell_value, str) and
                                cell_value.replace(',', '').replace('.', '', 1).replace('$', '').isdigit()):
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        else:
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            # Further table improvements
            # Add subtle alternating row colors for better readability
            for i in range(1, num_rows):
                if i % 2 == 1:  # Odd rows (excluding header)
                    for j in range(num_cols):
                        cell = table.cell(i, j)
                        tcPr = cell._tc.get_or_add_tcPr()
                        shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F5F5F5"/>')  # Very light gray
                        tcPr.append(shading_elm)
        
        except Exception as e:
            # If enhanced formatting fails, continue with basic table
            print(f"Warning: Some table formatting could not be applied: {str(e)}")
            pass
        
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
        self.input_values = {}  # Clear input values
    
    def clear_input_cache(self):
        """Clear the cached user inputs."""
        # Find all INPUT keys in session state and clear them
        keys_to_clear = [key for key in st.session_state.keys() if key.startswith("INPUT:")]
        for key in keys_to_clear:
            st.session_state[key] = ""
        self.form_submitted = False
        self.input_values = {}  # Clear input values


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
        {{INPUT:text:label:value}}              # Text input with label and default value
        {{INPUT:area:label:value:height}}       # Text area (multi-line) with label, default value, and optional height in pixels
        {{INPUT:date:label:value:format}}       # Date input with label, default date and format
                                               # Format can be "YYYY/MM/DD", "DD/MM/YYYY", or "MM/DD/YYYY"
                                               # Default value can be "today" or a date matching the format
        {{INPUT:select:label:option1,option2}}  # Dropdown selection with label and comma-separated options
        {{INPUT:check:label:True}}              # Checkbox with label and default state (True/False)
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