import re
import json
import os
import streamlit as st
from datetime import datetime
from excel_manager import excelManager

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
        self.input_cache = {}  # Cache for user inputs to avoid repeated prompts
        
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
        result = input_string
        
        # Process each keyword
        for match in matches:
            keyword = match.group(0)  # The full keyword with {{}}
            content = match.group(1)  # Just the content inside {{}}
            
            # Process the keyword and get its value
            replacement = self._process_keyword(content)
            
            # Replace the keyword with its value in the result string
            result = result.replace(keyword, str(replacement) if replacement is not None else "")
            
        return result
    
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
            
        # Process user input keywords
        elif keyword_type == "INPUT":
            return self._process_input_keyword(parts[1] if len(parts) > 1 else "")
            
        # Process formatting keywords
        elif keyword_type == "FORMAT":
            return self._process_format_keyword(parts[1] if len(parts) > 1 else "")
            
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
            
        # Check if the content contains a range (e.g., A1:C3)
        if ":" in content and "!" not in content.split(":")[1]:
            try:
                # Handle sheet references like 'Sheet With Spaces'!A1:C3
                if "!" in content:
                    parts = content.split("!")
                    sheet_name = parts[0].strip("'")  # Remove single quotes
                    cell_range = parts[1]
                    
                    # Case-insensitive sheet name lookup
                    if sheet_name.lower() in sheet_name_map:
                        actual_sheet_name = sheet_name_map[sheet_name.lower()]
                        return self.excel_manager.read_range(actual_sheet_name, cell_range)
                    else:
                        return f"[Sheet not found: {sheet_name}]"
                else:
                    # Use the active sheet
                    sheet_name = available_sheets[0]
                    return self.excel_manager.read_range(sheet_name, content)
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
    
    def _process_input_keyword(self, content):
        """Process user input keywords."""
        if not content:
            return "[Invalid input reference]"
            
        # Split the content by colon to check for input type
        parts = content.split(":", 1)
        field_name = parts[0].strip()
        
        # Check if we already have this input cached
        cache_key = f"INPUT:{content}"
        if cache_key in self.input_cache:
            return self.input_cache[cache_key]
            
        # Initialize session state for the input if not already done
        if cache_key not in st.session_state:
            st.session_state[cache_key] = ""

        # Handle different input types
        if len(parts) == 1:
            # Basic text input
            user_input = st.text_area(f"Enter value for {field_name}:", value=st.session_state[cache_key])
            if st.button('INSERT'):
                st.session_state[cache_key] = user_input
            return st.session_state[cache_key]

        input_type = parts[1].split(":", 1)[0].lower() if len(parts[1].split(":", 1)) > 1 else ""
        
        if input_type == "select":
            # Dropdown selection input
            options = parts[1].split(":", 1)[1].split(",") if len(parts[1].split(":", 1)) > 1 else []
            selected = st.selectbox(f"Select a value for {field_name}", options, index=options.index(st.session_state[cache_key]) if st.session_state[cache_key] in options else 0)
            if st.button('INSERT'):
                st.session_state[cache_key] = selected
            return st.session_state[cache_key]

        elif input_type == "date":
            # Date input
            date_format = parts[1].split(":", 1)[1] if len(parts[1].split(":", 1)) > 1 else "YYYY-MM-DD"
            selected_date = st.date_input(f"Select a date for {field_name}", value=datetime.strptime(st.session_state[cache_key], "%Y-%m-%d") if st.session_state[cache_key] else datetime.now())
            
            # Format the date according to the specified format
            formatted_date = selected_date.strftime("%Y-%m-%d")  # Default format
            # More format mappings could be added here
            
            if st.button('INSERT'):
                st.session_state[cache_key] = formatted_date
            return st.session_state[cache_key]

        else:
            # Text input with default value
            default_value = parts[1]
            user_input = st.text_area(f"Enter value for {field_name}", value=st.session_state[cache_key] or default_value)
            if st.button('INSERT'):
                st.session_state[cache_key] = user_input
            return st.session_state[cache_key]
    
    def _process_format_keyword(self, content):
        """Process formatting keywords."""
        if not content or ":" not in content:
            return "[Invalid FORMAT statement]"
            
        try:
            # Split into value and format type
            parts = content.split(":", 1)
            value_part, format_part = parts
            
            # Handle Excel value
            if value_part.startswith("XL:"):
                value = self._process_excel_keyword(value_part[3:])
            else:
                value = value_part
                
            # Apply formatting
            format_parts = format_part.split(":", 1)
            format_type = format_parts[0].lower()
            
            if format_type == "currency":
                try:
                    return f"${float(value):.2f}"
                except (ValueError, TypeError):
                    return f"${value}"
            
            elif format_type == "date":
                try:
                    if isinstance(value, str):
                        # Try to parse the date string
                        date_obj = datetime.strptime(value, "%Y-%m-%d")
                    else:
                        # Assume it's already a date object
                        date_obj = value
                        
                    # Apply the specified date format
                    if len(format_parts) > 1:
                        format_str = format_parts[1]
                        # Map the format string to strftime format
                        format_map = {
                            "MM/DD/YY": "%m/%d/%y",
                            "MM/DD/YYYY": "%m/%d/%Y",
                            "DD/MM/YYYY": "%d/%m/%Y",
                            "YYYY-MM-DD": "%Y-%m-%d"
                        }
                        date_format = format_map.get(format_str, "%Y-%m-%d")
                        return date_obj.strftime(date_format)
                    else:
                        return date_obj.strftime("%Y-%m-%d")
                except Exception:
                    return value
            
            else:
                return f"[Unknown format type: {format_type}]"
                
        except Exception as e:
            return f"[Error in FORMAT: {str(e)}]"
    
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
    
    def clear_input_cache(self):
        """Clear the cached user inputs."""
        self.input_cache = {}

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
        {{XL:Sales!C3:C7}}    # Range of cells (returns concatenated values or list)
        {{XL:named_range}}    # Named range in Excel
        ```

        ### User Input Keywords
        ```
        {{INPUT:field_name}}          # Basic text input
        {{INPUT:field_name:default}}  # With default value
        {{INPUT:date:YYYY-MM-DD}}     # Date input with format
        {{INPUT:select:option1,option2,option3}}  # Dropdown selection
        ```

        ### Formatting Keywords
        ```
        {{FORMAT:XL:B2:currency}}     # Format as currency
        {{FORMAT:XL:C3:date:MM/DD/YY}}  # Format as date
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