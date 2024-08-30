import pandas as pd
import os
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import streamlit as st
import tempfile

def extract_and_merge_columns_basic(old_file, new_file, old_keys, new_keys, old_cols, new_cols):
    """Perform a basic merge without additional columns."""
    # Concatenate primary keys to create unique identifiers
    old_file['UniqueKey'] = old_file[old_keys].astype(str).agg('_'.join, axis=1)
    new_file['UniqueKey'] = new_file[new_keys].astype(str).agg('_'.join, axis=1)

    # Extract specified columns from old and new files
    old_data = old_file[['UniqueKey'] + old_cols]
    new_data = new_file[['UniqueKey'] + new_cols]

    # Perform a VLOOKUP-like merge based on the unique identifiers
    merged_data = pd.merge(old_data, new_data, on='UniqueKey', how='inner')

    # Drop duplicate rows based on the UniqueKey column
    merged_data = merged_data.drop_duplicates(subset='UniqueKey')

    # Create a temporary file to save the merged data
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        merged_file_path = temp_file.name

    with pd.ExcelWriter(merged_file_path, engine='openpyxl') as writer:
        merged_data.to_excel(writer, sheet_name='Merged', index=False)

    return merged_file_path

def calculate_delay_days(etb_atb_col, proforma_col):
    """Calculate the Delay Days column based on the ETB / ATB and Proforma Berth columns."""
    # Convert ETB / ATB and Proforma Berth to datetime, assuming time format
    etb_atb = pd.to_datetime(etb_atb_col, format='%H:%M:%S', errors='coerce')
    proforma = pd.to_datetime(proforma_col, format='%H:%M:%S', errors='coerce')

    # Calculate delay days as the difference in days
    delay_days = (etb_atb - proforma).dt.total_seconds() / 86400  # Difference in days

    # Round up the delay days to the nearest whole number
    delay_days = np.ceil(delay_days)
    return delay_days

def calculate_delay_status(delay_days):
    """Calculate the Delay Status based on Delay Days."""
    if pd.isna(delay_days):
        return np.nan
    if -14 <= delay_days <= -1:
        return 'Advance'
    elif 1 <= delay_days <= 14:
        return 'Delay'
    else:
        return np.nan

def extract_and_merge_columns_with_delay(old_file, new_file, old_keys, new_keys, old_cols, new_cols):
    """Perform merge and add Delay Days and Delay Status columns."""
    # Concatenate primary keys to create unique identifiers
    old_file['UniqueKey'] = old_file[old_keys].astype(str).agg('_'.join, axis=1)
    new_file['UniqueKey'] = new_file[new_keys].astype(str).agg('_'.join, axis=1)

    # Extract specified columns from old and new files
    old_data = old_file[['UniqueKey'] + old_cols]
    new_data = new_file[['UniqueKey'] + new_cols]

    # Perform a VLOOKUP-like merge based on the unique identifiers
    merged_data = pd.merge(old_data, new_data, on='UniqueKey', how='inner')

    # Drop duplicate rows based on the UniqueKey column
    merged_data = merged_data.drop_duplicates(subset='UniqueKey')

    # Add Delay Days and Delay Status columns
    merged_data['Delay Days'] = calculate_delay_days(
        merged_data['ETB / ATB'], 
        merged_data['Proforma Berth']
    )
    merged_data['Delay Status'] = merged_data['Delay Days'].apply(calculate_delay_status)

    # Create a temporary file to save the merged data
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        merged_file_path = temp_file.name

    with pd.ExcelWriter(merged_file_path, engine='openpyxl') as writer:
        merged_data.to_excel(writer, sheet_name='Merged', index=False)

    return merged_file_path

def main():
    st.title('Vessel Schedule Report Compile')

    # Instructions for users
    st.markdown("""
    ---- Upload Files ----
    **Step 1**: Upload C-Report in the Excel A box  
    **Step 2**: Upload Berthing Report in the Excel B box
    
    ---- Define Primary Keys for Merging ----
    **Step 3**: Enter `BKH - Vessel Name,BKH - Voyage Ref` in the Primary Key A box  
    **Step 4**: Enter `Vessel Name,Voyage Ref` in the Primary Key B box  
    **Step 5**: Enter `BKH - Vessel Name,BKH - Voyage Ref` in the Output Columns A box  
    **Step 6**: Enter `ETB / ATB,Proforma Berth` in the Output Columns B box
    
    ---- Download merged C-Report & Berthing Report ----
    **Step 7**: Click "Basic Merge" and download the file
    
    ---- Upload the files 2nd batch for merging ----
    **Step 8**: Upload the first merged report file in the Excel A box  
    **Step 9**: Upload Terminal Report in the Excel B box  
    **Step 10**: Enter `BKH - Vessel Name` in the Primary Key A box  
    **Step 11**: Enter `Vessel Name` in the Primary Key B box  
    **Step 12**: Enter `BKH - Vessel Name,BKH - Voyage Ref,ETB / ATB,Proforma Berth` in the Output Columns A box  
    **Step 13**: Enter `SCN,ETA,ATA,Yard Closing` in the Output Columns B box  
    **Step 14**: Click "Merge with Delay Formula" and download the final report named `Vessel_Notification_Report.xlsx`
    """)

    # Upload widgets for the old and new Excel files for column extraction
    old_file_upload = st.file_uploader("Upload Excel A for merging (both files must have the same primary keys)", type=['xlsx'], key='old')
    new_file_upload = st.file_uploader("Upload Excel B for merging (both files must have the same primary keys)", type=['xlsx'], key='new')

    # Specify the primary keys and columns to extract
    old_keys = st.text_input("Unique Key - Enter Primary Key Column names for Excel A (separated by commas & no space)", "PrimaryKeyA1,PrimaryKeyA2")
    new_keys = st.text_input("Unique Key - Enter Primary Key Column names for Excel B (separated by commas & no space)", "PrimaryKeyB1,PrimaryKeyB2")
    old_columns = st.text_input("Enter Column names (include primary keys) to extract from Excel A (separated by commas & no space)", "Column1,Column2..etc")
    new_columns = st.text_input("Enter Column names to extract from Excel B (separated by commas & no space)", "Column1,Column2..etc")

    st.subheader("1. Basic Merge")
    if st.button("Extract and Merge Columns (Basic)"):
        if old_file_upload and new_file_upload:
            try:
                old_file = pd.read_excel(old_file_upload)
                new_file = pd.read_excel(new_file_upload)

                # Convert input columns to lists
                old_keys_list = [key.strip() for key in old_keys.split(',')]
                new_keys_list = [key.strip() for key in new_keys.split(',')]
                old_cols = [col.strip() for col in old_columns.split(',')]
                new_cols = [col.strip() for col in new_columns.split(',')]

                # Check if keys and columns are in the files
                if not all(key in old_file.columns for key in old_keys_list) or not all(col in old_file.columns for col in old_cols):
                    st.error("One or more keys/columns not found in Excel A.")
                    return
                if not all(key in new_file.columns for key in new_keys_list) or not all(col in new_file.columns for col in new_cols):
                    st.error("One or more keys/columns not found in Excel B.")
                    return

                merged_file_path = extract_and_merge_columns_basic(old_file, new_file, old_keys_list, new_keys_list, old_cols, new_cols)

                st.write("Basic merge completed successfully.")
                with open(merged_file_path, "rb") as f:
                    st.download_button("Download Merged File (Basic)", f, file_name="merged_output_basic.xlsx")

                # Clean up temporary file
                os.remove(merged_file_path)

            except Exception as e:
                st.error(f"An error occurred: {e}")
        else:
            st.error("Please upload both Excel files.")

    st.subheader("2. Merge with Delay Columns")
    if st.button("Extract and Merge Columns (With Delay Info)"):
        if old_file_upload and new_file_upload:
            try:
                old_file = pd.read_excel(old_file_upload)
                new_file = pd.read_excel(new_file_upload)

                # Convert input columns to lists
                old_keys_list = [key.strip() for key in old_keys.split(',')]
                new_keys_list = [key.strip() for key in new_keys.split(',')]
                old_cols = [col.strip() for col in old_columns.split(',')]
                new_cols = [col.strip() for col in new_columns.split(',')]

                # Check if keys and columns are in the files
                if not all(key in old_file.columns for key in old_keys_list) or not all(col in old_file.columns for col in old_cols):
                    st.error("One or more keys/columns not found in Excel A.")
                    return
                if not all(key in new_file.columns for key in new_keys_list) or not all(col in new_file.columns for col in new_cols):
                    st.error("One or more keys/columns not found in Excel B.")
                    return

                merged_file_path = extract_and_merge_columns_with_delay(old_file, new_file, old_keys_list, new_keys_list, old_cols, new_cols)

                st.write("Merge with delay columns completed successfully.")
                with open(merged_file_path, "rb") as f:
                    st.download_button("Download Merged File (With Delay Info)", f, file_name="merged_output_with_delay.xlsx")

                # Clean up temporary file
                os.remove(merged_file_path)

            except Exception as e:
                st.error(f"An error occurred: {e}")
        else:
            st.error("Please upload both Excel files.")

if __name__ == "__main__":
    main()
