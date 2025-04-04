import streamlit as st
import os
import re
import docx
import tempfile
import time
from excel_manager import excelManager
from keyword_parser import keywordParser
from collections import Counter

def preprocess_word_doc(doc_path):
    """
    Analyze a Word document to determine what keywords it contains.
    
    Args:
        doc_path: Path to the Word document
        
    Returns:
        Dictionary with keyword counts and whether Excel file is needed
    """
    # Load the document
    doc = docx.Document(doc_path)
    
    # Compile regex pattern for keywords
    pattern = r'{{(.*?)}}'
    
    # Track keyword types and counts
    keywords = {
        "excel": [],
        "input": {
            "text": [],
            "area": [],
            "date": [],
            "select": [],
            "check": []
        },
        "template": [],
        "json": [],
        "other": []
    }
    
    # Track if Excel file is needed
    needs_excel = False
    
    # Total keyword count
    total_keywords = 0
    
    # Scan paragraphs for keywords
    for para_idx, paragraph in enumerate(doc.paragraphs):
        matches = list(re.finditer(pattern, paragraph.text))
        total_keywords += len(matches)
        
        for match in matches:
            content = match.group(1)  # Content inside {{}}
            parts = content.split(":", 1)
            keyword_type = parts[0].strip().upper()
            
            if keyword_type == "XL":
                needs_excel = True
                keywords["excel"].append(content)
            elif keyword_type == "SUM" or keyword_type == "AVG":
                if len(parts) > 1 and parts[1].strip().upper().startswith("XL:"):
                    needs_excel = True
                keywords["excel"].append(content)
            elif keyword_type == "INPUT":
                if len(parts) > 1:
                    input_parts = parts[1].split(":")
                    input_type = input_parts[0].lower() if input_parts else "text"
                    
                    if input_type in keywords["input"]:
                        keywords["input"][input_type].append(content)
                    else:
                        keywords["input"]["text"].append(content)
            elif keyword_type == "TEMPLATE":
                keywords["template"].append(content)
            elif keyword_type == "JSON":
                keywords["json"].append(content)
            else:
                keywords["other"].append(content)
    
    # Scan tables for keywords
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    matches = list(re.finditer(pattern, paragraph.text))
                    total_keywords += len(matches)
                    
                    for match in matches:
                        content = match.group(1)  # Content inside {{}}
                        parts = content.split(":", 1)
                        keyword_type = parts[0].strip().upper()
                        
                        if keyword_type == "XL":
                            needs_excel = True
                            keywords["excel"].append(content)
                        elif keyword_type == "SUM" or keyword_type == "AVG":
                            if len(parts) > 1 and parts[1].strip().upper().startswith("XL:"):
                                needs_excel = True
                            keywords["excel"].append(content)
                        elif keyword_type == "INPUT":
                            if len(parts) > 1:
                                input_parts = parts[1].split(":")
                                input_type = input_parts[0].lower() if input_parts else "text"
                                
                                if input_type in keywords["input"]:
                                    keywords["input"][input_type].append(content)
                                else:
                                    keywords["input"]["text"].append(content)
                        elif keyword_type == "TEMPLATE":
                            keywords["template"].append(content)
                        elif keyword_type == "JSON":
                            keywords["json"].append(content)
                        else:
                            keywords["other"].append(content)
    
    # Prepare summary
    summary = {
        "total_keywords": total_keywords,
        "excel_count": len(keywords["excel"]),
        "input_counts": {k: len(v) for k, v in keywords["input"].items()},
        "template_count": len(keywords["template"]),
        "json_count": len(keywords["json"]),
        "other_count": len(keywords["other"]),
        "needs_excel": needs_excel,
        "keywords": keywords
    }
    
    return summary

def process_word_doc(doc_path, excel_path=None):
    """
    Process a Word document, replacing keywords with values from Excel spreadsheet.
    
    Args:
        doc_path: Path to the Word document
        excel_path: Path to the Excel spreadsheet (optional)
        
    Returns:
        Processed document object and a count of replaced keywords
        
    Note:
        If form is waiting for submission, returns (None, 0)
    """
    # Load the document
    doc = docx.Document(doc_path)
    
    # Initialize Excel manager if needed
    excel_mgr = None
    if excel_path:
        excel_mgr = excelManager(excel_path)
    
    # Initialize keyword parser with the Excel manager and pass the document reference
    parser = keywordParser(excel_mgr)
    parser.set_word_document(doc)
    
    # Compile regex pattern for keywords
    pattern = r'{{(.*?)}}'
    
    # Count total keywords for progress tracking
    total_keywords = 0
    
    # Count keywords in paragraphs
    for paragraph in doc.paragraphs:
        total_keywords += len(re.findall(pattern, paragraph.text))
    
    # Count keywords in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    total_keywords += len(re.findall(pattern, paragraph.text))
    
    if total_keywords == 0:
        st.warning("No keywords found in the document.")
        return doc, 0
    
    # Initialize progress bar
    progress_bar = st.progress(0)
    progress_text = st.empty()
    
    # Counter for processed keywords
    processed_count = 0
    
    # First handle all INPUT keywords with a single form if needed
    input_keywords = []
    input_locations = []
    
    # Create a unique identifier for each keyword occurrence
    keyword_positions = {}
    
    # Collect all INPUT keywords from paragraphs with unique position identifiers
    for para_idx, paragraph in enumerate(doc.paragraphs):
        matches = list(re.finditer(pattern, paragraph.text))
        for match_idx, match in enumerate(matches):
            keyword = match.group(0)
            content = match.group(1)
            if content.split(":", 1)[0].strip().upper() == "INPUT":
                # Assign a completely unique position ID
                position_id = f"p{para_idx}_m{match_idx}"
                keyword_positions[position_id] = keyword
                input_keywords.append((position_id, keyword, content))
                input_locations.append(("paragraph", paragraph, match.start(), match.end()))
    
    # Collect all INPUT keywords from tables with unique position identifiers
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for para_idx, paragraph in enumerate(cell.paragraphs):
                    matches = list(re.finditer(pattern, paragraph.text))
                    for match_idx, match in enumerate(matches):
                        keyword = match.group(0)
                        content = match.group(1)
                        if content.split(":", 1)[0].strip().upper() == "INPUT":
                            # Assign a completely unique position ID for table cells
                            position_id = f"t{table_idx}_r{row_idx}_c{cell_idx}_p{para_idx}_m{match_idx}"
                            keyword_positions[position_id] = keyword
                            input_keywords.append((position_id, keyword, content))
                            input_locations.append(("table_cell", paragraph, match.start(), match.end()))
    
    # Store if we have any input keywords
    has_inputs = len(input_keywords) > 0
    
    # If there are input keywords, process them all at once
    input_values = {}
    
    # Create a form state storage key
    form_key = f"form_processed_{doc_path}"
    if form_key not in st.session_state:
        st.session_state[form_key] = False
        
    if has_inputs and not st.session_state[form_key]:
        with st.form(key="document_input_form"):
            st.subheader("Please provide values for input fields:")
            
            for i, (position_id, keyword, content) in enumerate(input_keywords):
                # Extract the parts from the content
                parts = content.split(":")
                
                # Create a truly unique key for each input field based on position
                unique_key = f"input_{position_id}"
                
                # Create appropriate input field based on the content
                if len(parts) > 1:
                    input_type = parts[1].strip().lower()
                    label = parts[2] if len(parts) > 2 else f"Input {i+1}"
                    default = parts[3] if len(parts) > 3 else ""
                    
                    if input_type == "text":
                        value = st.text_input(label, value=default, key=unique_key)
                    elif input_type == "area":
                        value = st.text_area(label, value=default, key=unique_key)
                    elif input_type == "date":
                        import datetime
                        if default.lower() == "today":
                            default_date = datetime.date.today()
                        else:
                            try:
                                default_date = datetime.datetime.strptime(default, "%Y/%m/%d").date()
                            except ValueError:
                                default_date = datetime.date.today()
                        
                        date_value = st.date_input(label, value=default_date, key=unique_key)
                        value = date_value.strftime("%Y/%m/%d")
                    elif input_type == "select":
                        options_str = default
                        options = [opt.strip() for opt in options_str.split(",")] if options_str else ["Option 1"]
                        value = st.selectbox(label, options, key=unique_key)
                    elif input_type == "check":
                        default_bool = default.lower() == "true"
                        value = st.checkbox(label, value=default_bool, key=unique_key)
                    else:
                        value = st.text_input(label, value=default, key=unique_key)
                else:
                    value = st.text_input(f"Input {i+1}", key=unique_key)
                
                # Store the value with the position ID
                input_values[position_id] = value
            
            submit = st.form_submit_button("Submit")
            if submit:
                # Mark the form as submitted in session state
                st.session_state[form_key] = True
                # Store the input values in session state for persistence
                st.session_state[f"input_values_{doc_path}"] = input_values
                # Force a rerun to continue processing after form submission
                st.rerun()
            else:
                st.info("Please fill in all fields and click Submit to continue.")
                return None, 0  # Return None to indicate form is not submitted yet
    
    # If we have inputs but form wasn't submitted yet, stop here
    if has_inputs and not st.session_state[form_key]:
        return None, 0
    
    # Retrieve stored input values if they exist
    if has_inputs and f"input_values_{doc_path}" in st.session_state:
        input_values = st.session_state[f"input_values_{doc_path}"]
    
    # We're now ready to process the document
    progress_text.text("Processing keywords...")
    
    # Track paragraphs with Excel range keywords
    range_paragraphs = []
    
    # First scan to find paragraphs containing Excel range keywords
    for paragraph in doc.paragraphs:
        matches = list(re.finditer(pattern, paragraph.text))
        for match in matches:
            content = match.group(1)
            parts = content.split(":", 1)
            if parts[0].strip().upper() == "XL" and len(parts) > 1:
                # Check if it's a range reference like A1:C3
                if ":" in parts[1] and "!" not in parts[1].split(":", 1)[1]:
                    # Store the paragraph for special handling
                    if paragraph not in range_paragraphs:
                        range_paragraphs.append(paragraph)
    
    # Create a reverse mapping from original keywords to position IDs
    keyword_to_position = {}
    for position_id, keyword in keyword_positions.items():
        if keyword not in keyword_to_position:
            keyword_to_position[keyword] = []
        keyword_to_position[keyword].append(position_id)
    
    # Keep track of which occurrence of each keyword we're currently processing
    keyword_counters = {}
    
    # Now process all keywords in paragraphs, with special handling for range paragraphs
    for para_idx, paragraph in enumerate(doc.paragraphs):
        # If this is a range paragraph that needs special handling for table insertion
        if paragraph in range_paragraphs:
            # Get all keywords in this paragraph
            matches = list(re.finditer(pattern, paragraph.text))
            
            for match_idx, match in enumerate(matches):
                keyword = match.group(0)
                content = match.group(1)
                parts = content.split(":", 1)
                
                # Process Excel range keywords
                if parts[0].strip().upper() == "XL" and len(parts) > 1:
                    # If it's an Excel range reference
                    if ":" in parts[1] and "!" not in parts[1].split(":", 1)[1]:
                        try:
                            # Get paragraph text and position
                            orig_text = paragraph.text
                            start_pos = match.start()
                            end_pos = match.end()
                            
                            # Parse the keyword to potentially create a table
                            parser.set_word_document(doc)
                            replacement = parser.parse(keyword)
                            
                            # If a table was inserted
                            if replacement == "[TABLE_INSERTED]":
                                # Update the paragraph text to remove the keyword
                                if start_pos == 0 and end_pos == len(orig_text):
                                    # If keyword is the entire paragraph, empty it
                                    paragraph.text = ""
                                else:
                                    # Otherwise remove just the keyword
                                    paragraph.text = orig_text[:start_pos] + orig_text[end_pos:]
                            else:
                                # If table insertion failed, replace with text normally
                                paragraph.text = orig_text[:start_pos] + str(replacement) + orig_text[end_pos:]
                        except Exception as e:
                            # Handle errors
                            replacement = f"[Error: {str(e)}]"
                            paragraph.text = paragraph.text[:match.start()] + replacement + paragraph.text[match.end():]
                        
                        # Update progress
                        processed_count += 1
                        progress_bar.progress(processed_count / total_keywords)
                        progress_text.text(f"Processing keywords: {processed_count}/{total_keywords}")
                        continue
                
                # Process non-range keywords
                # If it's an INPUT keyword we need to find its unique position ID
                if content.split(":", 1)[0].strip().upper() == "INPUT":
                    # Get position identifier for this specific occurrence
                    position_id = f"p{para_idx}_m{match_idx}"
                    
                    if position_id in input_values:
                        replacement = input_values[position_id]
                    else:
                        # Fallback to parser
                        parser.set_word_document(None)
                        replacement = parser.parse(keyword)
                else:
                    # Otherwise parse the keyword normally
                    parser.set_word_document(None)  # Don't use direct table insertion
                    replacement = parser.parse(keyword)
                
                # Replace the keyword in the paragraph text
                paragraph.text = paragraph.text[:match.start()] + str(replacement) + paragraph.text[match.end():]
                
                # Update progress
                processed_count += 1
                progress_bar.progress(processed_count / total_keywords)
                progress_text.text(f"Processing keywords: {processed_count}/{total_keywords}")
        else:
            # This is a regular paragraph, process normally
            matches = list(re.finditer(pattern, paragraph.text))
            for match_idx, match in enumerate(reversed(matches)):  # Process in reverse to avoid index issues
                keyword = match.group(0)  # Full keyword with {{}}
                content = match.group(1)  # Content inside {{}}
                
                # Process based on the keyword type
                if content.split(":", 1)[0].strip().upper() == "INPUT":
                    # Calculate the correct match index for the original order (since we're iterating in reverse)
                    orig_match_idx = len(matches) - 1 - match_idx
                    position_id = f"p{para_idx}_m{orig_match_idx}"
                    
                    if position_id in input_values:
                        replacement = input_values[position_id]
                    else:
                        # Fallback to parser
                        parser.set_word_document(None)
                        replacement = parser.parse(keyword)
                else:
                    # Otherwise parse the keyword normally
                    parser.set_word_document(None)  # Don't use direct table insertion
                    replacement = parser.parse(keyword)
                
                # Replace the keyword in the paragraph text
                paragraph.text = paragraph.text[:match.start()] + str(replacement) + paragraph.text[match.end():]
                
                # Update progress
                processed_count += 1
                progress_bar.progress(processed_count / total_keywords)
                progress_text.text(f"Processing keywords: {processed_count}/{total_keywords}")
    
    # Process tables with position tracking for input keywords
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for para_idx, paragraph in enumerate(cell.paragraphs):
                    # Find all keywords in the paragraph
                    matches = list(re.finditer(pattern, paragraph.text))
                    
                    # Process each keyword
                    for match_idx, match in enumerate(reversed(matches)):  # Process in reverse to avoid index issues
                        keyword = match.group(0)  # Full keyword with {{}}
                        content = match.group(1)  # Content inside {{}}
                        
                        # Calculate the correct match index for the original order (since we're iterating in reverse)
                        orig_match_idx = len(matches) - 1 - match_idx
                        
                        # Process based on the keyword type
                        if content.split(":", 1)[0].strip().upper() == "INPUT":
                            position_id = f"t{table_idx}_r{row_idx}_c{cell_idx}_p{para_idx}_m{orig_match_idx}"
                            
                            if position_id in input_values:
                                replacement = input_values[position_id]
                            else:
                                # Fallback to parser
                                parser.set_word_document(None)
                                replacement = parser.parse(keyword)
                        else:
                            # For tables, don't use direct table insertion
                            parser.set_word_document(None)
                            replacement = parser.parse(keyword)
                        
                        # Replace the keyword in the paragraph text
                        paragraph.text = paragraph.text[:match.start()] + str(replacement) + paragraph.text[match.end():]
                        
                        # Update progress
                        processed_count += 1
                        progress_bar.progress(processed_count / total_keywords)
                        progress_text.text(f"Processing keywords: {processed_count}/{total_keywords}")
    
    # Ensure progress is complete
    progress_bar.progress(1.0)
    progress_text.text(f"Processed {processed_count} keywords.")
    
    # Close Excel manager if it was created
    if excel_mgr:
        excel_mgr.close()
    
    return doc, processed_count

def display_keyword_summary(summary):
    """
    Display a summary of the keywords found in the document.
    
    Args:
        summary: Dictionary with keyword counts and information
    """
    # Display total keywords outside the expander
    st.write(f"Total keywords found: **{summary['total_keywords']}**")
    
    # Show detailed breakdown in an expander
    with st.expander("Document Analysis Summary"):
        # Create 4 columns for the different keyword types
        col1, col2, col3, col4 = st.columns(4)
        
        # COLUMN 1: Excel Keywords
        with col1:
            st.markdown("**Excel Keywords**")
            if summary["excel_count"] > 0:
                st.write(f"Total: {summary['excel_count']}")
                st.write("*Excel spreadsheet required*")
            else:
                st.write("Total: 0")
        
        # COLUMN 2: User Input Keywords
        with col2:
            st.markdown("**User Input Keywords**")
            total_inputs = sum(summary["input_counts"].values())
            st.write(f"Total: {total_inputs}")
            
            # List all input types, even if zero
            input_types = ["text", "area", "date", "select", "check"]
            for input_type in input_types:
                count = summary["input_counts"].get(input_type, 0)
                st.write(f"{input_type}: {count}")
        
        # COLUMN 3: Template Keywords
        with col3:
            st.markdown("**Template Keywords**")
            st.write(f"Total: {summary['template_count']}")
        
        # COLUMN 4: JSON Keywords
        with col4:
            st.markdown("**JSON/Other Keywords**")
            st.write(f"JSON: {summary['json_count']}")
            st.write(f"Other: {summary['other_count']}")

def main():
    st.title("Document Keyword Parser")
    st.write("Upload a Word document and process keywords based on document analysis.")
    
    # Add a more detailed description
    st.markdown("""
    This tool analyzes your Word document for keywords and processes them accordingly:
    * Automatically detects what types of keywords are present
    * Only asks for Excel upload when needed
    * Handles user input fields, template references, and more
    """)
    
    # Show keyword help
    with st.expander("Keyword Reference Guide"):
        parser = keywordParser()
        st.markdown(parser.get_keyword_help())
    
    # File upload section - initially just for Word document
    doc_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"])
    
    # Initialize session state for tracking processing status
    if 'preprocessing_complete' not in st.session_state:
        st.session_state.preprocessing_complete = False
    if 'preprocessing_results' not in st.session_state:
        st.session_state.preprocessing_results = None
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'output_path' not in st.session_state:
        st.session_state.output_path = None
    if 'processed_count' not in st.session_state:
        st.session_state.processed_count = 0
    if 'doc_path' not in st.session_state:
        st.session_state.doc_path = None
    
    # Process the Word document for analysis when uploaded
    if doc_file and not st.session_state.preprocessing_complete:
        st.subheader("Analyzing Document")
        
        # Save uploaded file to temporary location
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_doc:
            tmp_doc.write(doc_file.getvalue())
            doc_path = tmp_doc.name
            st.session_state.doc_path = doc_path
        
        try:
            # Show progress bar for preprocessing
            preprocessing_progress = st.progress(0)
            preprocessing_text = st.empty()
            preprocessing_text.text("Analyzing document content...")
            
            # Simulate progress (visual feedback for the user)
            for i in range(10):
                preprocessing_progress.progress(i/10)
                time.sleep(0.1)
            
            # Preprocess the document
            summary = preprocess_word_doc(doc_path)
            
            # Complete the progress bar
            preprocessing_progress.progress(1.0)
            preprocessing_text.text("Analysis complete!")
            
            # Store the results
            st.session_state.preprocessing_complete = True
            st.session_state.preprocessing_results = summary
            
            # Force refresh to show results
            st.rerun()
            
        except Exception as e:
            st.error(f"An error occurred during analysis: {str(e)}")
    
    # If preprocessing is complete, show results and continue
    if st.session_state.preprocessing_complete and st.session_state.preprocessing_results:
        # Use a container for the summary to keep it together visually
        with st.container():
            # Display summary of keywords found with a small divider above
            st.markdown("---")
            display_keyword_summary(st.session_state.preprocessing_results)
            st.markdown("---")
        
        # Reset button - show only if not in the middle of processing
        if not st.session_state.get('process_clicked') or st.session_state.get('processing_complete'):
            if st.button("Reset Analysis"):
                # Clear session state
                for key in list(st.session_state.keys()):
                    if key != "keyword_help":  # Keep the help text
                        del st.session_state[key]
                st.rerun()
        
        # Continue with processing
        st.subheader("Process Document")
        
        # Check if Excel file is needed based on preprocessing
        needs_excel = st.session_state.preprocessing_results["needs_excel"]
        
        # If Excel needed, show Excel upload
        excel_file = None
        if needs_excel:
            excel_file = st.file_uploader("Upload Excel Spreadsheet (.xlsx)", type=["xlsx"])
            if not excel_file:
                st.warning("An Excel spreadsheet is required based on the keywords found in your document.")
                return
        
        # Initialize or retrieve the processing state
        if 'process_clicked' not in st.session_state:
            st.session_state.process_clicked = False
        
        # Process button
        if st.button("Process Document") or st.session_state.process_clicked:
            st.session_state.process_clicked = True
            
            # Check if we can continue
            if needs_excel and not excel_file:
                st.error("Please upload an Excel spreadsheet to continue.")
                return
            
            # Process the document
            st.subheader("Processing Document")
            
            # Save Excel file if needed
            excel_path = None
            if excel_file:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
                    tmp_excel.write(excel_file.getvalue())
                    excel_path = tmp_excel.name
            
            try:
                # Process the document
                processed_doc, count = process_word_doc(st.session_state.doc_path, excel_path)
                
                # If processed_doc is None, it means the form is waiting for submission
                if processed_doc is None:
                    st.info("Please complete the form and click Submit to continue.")
                    return
                
                if count > 0:
                    # Save the processed document
                    tmp_folder = "tmp"
                    if not os.path.exists(tmp_folder):
                        os.makedirs(tmp_folder)
                    output_path = os.path.join(tmp_folder, "processed_document.docx")
                    processed_doc.save(output_path)
                    
                    # Store the results in session state
                    st.session_state.processing_complete = True
                    st.session_state.output_path = output_path
                    st.session_state.processed_count = count
                    
                    # Show success message and download button immediately
                    st.success(f"Successfully processed {count} keywords!")
                    with open(output_path, "rb") as file:
                        st.download_button(
                            label="Download Processed Document",
                            data=file,
                            file_name="processed_document.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                else:
                    st.info("No keywords were processed. The document remains unchanged.")
                    
            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")
            
            finally:
                # Clean up temporary Excel file
                if 'excel_path' in locals() and excel_path and excel_path is not None:
                    try:
                        os.unlink(excel_path)
                    except Exception as e:
                        st.error(f"Error cleaning up temporary file: {str(e)}")
        
        # This section is now handled inside the processing logic to ensure 
        # the download button appears immediately after form submission
    
    # Additional information
    st.markdown("---")
    st.write("This app processes keywords in Word documents and replaces them with values from various sources.")
    
    # Clean up temporary files when the app is closed
    def cleanup():
        if 'doc_path' in st.session_state and st.session_state.doc_path:
            try:
                os.unlink(st.session_state.doc_path)
            except:
                pass
        if 'output_path' in st.session_state and st.session_state.output_path:
            try:
                os.unlink(st.session_state.output_path)
            except:
                pass
    
    # Register the cleanup function to be called when the script exits
    import atexit
    atexit.register(cleanup)

if __name__ == "__main__":
    main()