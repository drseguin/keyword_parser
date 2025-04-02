import streamlit as st
import os
import pandas as pd
import tempfile
from excel_manager import excelManager

st.title("Excel Manager App")

# Initialize session state
if 'excel_manager' not in st.session_state:
    st.session_state.excel_manager = None
if 'file_path' not in st.session_state:
    st.session_state.file_path = None
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = tempfile.mkdtemp()

# Function to reset the app
def reset_app():
    st.session_state.excel_manager = None
    st.session_state.file_path = None

# Sidebar for file operations
st.sidebar.header("File Operations")

# File upload
uploaded_file = st.sidebar.file_uploader("Upload Excel file", type=["xlsx", "xls"])
if uploaded_file is not None:
    # Save uploaded file to temp directory
    file_path = os.path.join(st.session_state.temp_dir, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # Initialize ExcelManager with the uploaded file
    st.session_state.excel_manager = excelManager(file_path)
    st.session_state.file_path = file_path
    st.sidebar.success(f"Loaded: {uploaded_file.name}")

# Create new file
new_file_name = st.sidebar.text_input("Or create a new file (name.xlsx):")
if st.sidebar.button("Create New File") and new_file_name:
    if not new_file_name.endswith(('.xlsx', '.xls')):
        new_file_name += '.xlsx'
    
    file_path = os.path.join(st.session_state.temp_dir, new_file_name)
    st.session_state.excel_manager = excelManager()
    st.session_state.excel_manager.create_workbook(file_path)
    st.session_state.file_path = file_path
    st.sidebar.success(f"Created: {new_file_name}")

# Reset app
if st.sidebar.button("Reset"):
    reset_app()
    st.sidebar.success("Reset complete")

# Main content
if st.session_state.excel_manager is not None:
    st.subheader("Excel File Management")
    
    # Tabs for different operations
    tab1, tab2, tab3, tab4 = st.tabs(["Sheets", "Read", "Write", "Delete"])
    
    with tab1:
        st.subheader("Sheet Operations")
        
        # Count sheets
        if st.button("Count Sheets"):
            count = st.session_state.excel_manager.count_sheets()
            st.info(f"Number of sheets: {count}")
        
        # Get sheet names
        if st.button("Get Sheet Names"):
            names = st.session_state.excel_manager.get_sheet_names()
            st.info(f"Sheet names: {', '.join(names)}")
        
        # Create new sheet
        new_sheet_name = st.text_input("New sheet name:")
        if st.button("Create Sheet") and new_sheet_name:
            st.session_state.excel_manager.create_sheet(new_sheet_name)
            st.success(f"Created sheet: {new_sheet_name}")
            st.session_state.excel_manager.save()
    
    with tab2:
        st.subheader("Read Operations")
        
        # Select sheet
        if st.session_state.excel_manager:
            sheet_names = st.session_state.excel_manager.get_sheet_names()
            selected_sheet = st.selectbox("Select sheet", sheet_names)
            
            # Read cell (using cell reference)
            st.subheader("Read Cell")
            cell_reference = st.text_input("Cell Reference (e.g. A1, B5):", "A1")
            
            if st.button("Read Cell"):
                try:
                    value = st.session_state.excel_manager.read_cell(selected_sheet, cell_reference)
                    st.info(f"Cell value: {value}")
                except Exception as e:
                    st.error(f"Error reading cell: {str(e)}")
            
            # Read range
            st.subheader("Read Range")
            range_reference = st.text_input("Range Reference (e.g. A1:C5):", "A1:B5")
            
            if st.button("Read Range"):
                try:
                    values = st.session_state.excel_manager.read_range(selected_sheet, range_reference)
                    # Convert to pandas DataFrame for better display
                    df = pd.DataFrame(values)
                    st.dataframe(df)
                except Exception as e:
                    st.error(f"Error reading range: {str(e)}")
            
            # Read total (new functionality)
            st.subheader("Read Total")
            total_start_reference = st.text_input("Starting Cell (e.g. A1, F25):", "A1", key="total_start_ref")
            
            if st.button("Find Total"):
                try:
                    total_value = st.session_state.excel_manager.read_total(selected_sheet, total_start_reference)
                    if total_value is not None:
                        st.info(f"Total value: {total_value}")
                    else:
                        st.warning("No total value found in this column.")
                except Exception as e:
                    st.error(f"Error finding total: {str(e)}")
            
            # Read items (new functionality)
            st.subheader("Read Items")
            items_start_reference = st.text_input("Starting Cell (e.g. A1, F25):", "A1", key="items_start_ref")
            offset_value = st.number_input("Offset (rows to exclude from end):", min_value=0, value=0, key="offset_value")
            
            if st.button("Find Items"):
                try:
                    items = st.session_state.excel_manager.read_items(selected_sheet, items_start_reference, offset=offset_value)
                    if items:
                        st.info(f"Found {len(items)} items:")
                        # Display items as a dataframe for better formatting
                        df = pd.DataFrame({"Items": items})
                        st.dataframe(df)
                    else:
                        st.warning("No items found starting from this cell.")
                except Exception as e:
                    st.error(f"Error finding items: {str(e)}")
    
    with tab3:
        st.subheader("Write Operations")
        
        # Select sheet
        if st.session_state.excel_manager:
            sheet_names = st.session_state.excel_manager.get_sheet_names()
            selected_sheet = st.selectbox("Select sheet", sheet_names, key="write_sheet")
            
            # Write cell (using cell reference)
            st.subheader("Write Cell")
            cell_reference = st.text_input("Cell Reference (e.g. A1, B5):", "A1", key="write_cell_ref")
            write_value = st.text_input("Value:", key="write_value")
            
            if st.button("Write Cell"):
                try:
                    st.session_state.excel_manager.write_cell(selected_sheet, cell_reference, write_value)
                    st.success(f"Wrote '{write_value}' to cell {cell_reference}")
                    st.session_state.excel_manager.save()
                except Exception as e:
                    st.error(f"Error writing cell: {str(e)}")
            
            # Write range (using CSV input)
            st.subheader("Write Range")
            start_cell = st.text_input("Start Cell (e.g. A1):", "A1", key="range_start_cell")
            
            csv_data = st.text_area(
                "Enter CSV data (comma-separated values, one row per line):",
                "1,2,3\n4,5,6\n7,8,9"
            )
            
            if st.button("Write Range"):
                try:
                    # Parse CSV data
                    rows = []
                    for line in csv_data.strip().split("\n"):
                        values = line.split(",")
                        rows.append(values)
                    
                    st.session_state.excel_manager.write_range(selected_sheet, start_cell, rows)
                    st.success(f"Wrote data to range starting at {start_cell}")
                    st.session_state.excel_manager.save()
                except Exception as e:
                    st.error(f"Error writing range: {str(e)}")
    
    with tab4:
        st.subheader("Delete Operations")
        
        # Delete sheet
        if st.session_state.excel_manager:
            sheet_names = st.session_state.excel_manager.get_sheet_names()
            sheet_to_delete = st.selectbox("Select sheet to delete", sheet_names)
            
            if st.button("Delete Sheet") and len(sheet_names) > 1:
                st.session_state.excel_manager.delete_sheet(sheet_to_delete)
                st.success(f"Deleted sheet: {sheet_to_delete}")
                st.session_state.excel_manager.save()
            elif len(sheet_names) <= 1:
                st.error("Cannot delete the only sheet in the workbook.")
    
    # Download the file
    if st.session_state.file_path:
        with open(st.session_state.file_path, "rb") as file:
            file_name = os.path.basename(st.session_state.file_path)
            st.download_button(
                label="Download Excel file",
                data=file,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Please upload an Excel file or create a new one to start.")