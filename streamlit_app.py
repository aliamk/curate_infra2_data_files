import streamlit as st
import pandas as pd
from datetime import datetime
import os
import tempfile

# Function to read the source file
def read_source_file(source_path):
    xl = pd.ExcelFile(source_path)
    df1 = xl.parse('Sheet1')
    df2 = xl.parse('Sheet2')
    return df1, df2

# Function to create the destination file
def create_destination_file(source_path):
    df1, df2 = read_source_file(source_path)
    
    columns_mapping = {
        'Transaction ID': 'Realfin INFRA Transaction Upload ID',
        'Transaction Name': 'Transaction Name',
        'Asset Class': 'Infrastructure',  # Column C has a fixed value
        'Transaction Status': 'Transaction Stage',
        'Finance Type': 'Finance Type',
        'Transaction Type': 'Transaction Type',
        'BlankG': None,  # Column G is blank
        'BlankH': None,  # Column H is blank
        'Transaction Local Currency': 'Transaction Currency',
        'Transaction Value (Local Currency)': 'Transaction Value (Local Currency m)',
        'Transaction Debt (Local Currency)': 'Transaction Debt (Local Currency m)',
        'Transaction Equity (Local Currency)': 'Transaction Equity (Local Currency m)',
        'Debt/Equity Ratio': 'Debt/Equity Ratio',
        'BlankN': None,  # Column N is blank
        'Region - Country': 'Transaction Country/Region',
        'BlankP': None,  # Column P is blank
        'BlankQ': None,  # Column Q is blank
        'Any Level Sectors': ['Transaction Sector', 'Transaction Sub-sector'],
        'PPP': 'PPP',
        'Concession Period': 'Concession Period',
        'Contract': 'Contract',
        'SPV': None,  # Column V will be filled later
        'Active': 'True',  # Column W has a fixed value 'True'
        'BlankX': None,  # Column X is blank
        'BlankY': None,  # Column Y is blank
        'BlankZ': None   # Column Z is blank
    }

    transaction_data = {}
    for dest_col, source_col in columns_mapping.items():
        if source_col is None:
            transaction_data[dest_col] = [None] * len(df1)
        elif source_col == 'Infrastructure':
            transaction_data[dest_col] = ['Infrastructure'] * len(df1)
        elif source_col == 'True':
            transaction_data[dest_col] = ['True'] * len(df1)
        elif isinstance(source_col, list):
            transaction_data[dest_col] = df1[source_col[0]].astype(str) + ', ' + df1[source_col[1]].astype(str)
        else:
            transaction_data[dest_col] = df1[source_col] if source_col in df1.columns else [None] * len(df1)

    transaction_df = pd.DataFrame(transaction_data)
    
    spv_mapping = df2.set_index('Realfin INFRA Transaction Upload ID')['SPV'].dropna().to_dict()
    transaction_df['SPV'] = transaction_df['Transaction ID'].map(spv_mapping)

    column_names = list(transaction_df.columns)
    for idx in range(len(column_names)):
        if column_names[idx].startswith('Blank'):
            column_names[idx] = ''
    transaction_df.columns = column_names
    
    base, ext = os.path.splitext(source_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destination_filename = f"{base}_Destination_{timestamp}.xlsx"
    
    with pd.ExcelWriter(destination_filename, engine='openpyxl') as writer:
        transaction_df.to_excel(writer, sheet_name='Transaction', index=False)
        
        # Create empty tabs with specified headers
        pd.DataFrame(columns=['Transaction Upload ID', 'Asset Upload ID']).to_excel(writer, sheet_name='Underlying_Asset', index=False)
        pd.DataFrame(columns=['Transaction Upload ID', 'Event Date', 'Event Type', 'Event Title']).to_excel(writer, sheet_name='Events', index=False)
        pd.DataFrame(columns=['Transaction Upload ID', 'Role Type', 'Role Subtype', 'Company', 'Fund', 'Bidder Status', 'Client Counterparty', 'Client Company Name', 'Fund Name']).to_excel(writer, sheet_name='Bidders_Any', index=False)
        pd.DataFrame(columns=['Transaction Upload ID', 'Tranche Upload ID', 'Tranche Primary Type', 'Tranche Secondary Type', 'Tranche Tertiary Type', 'Value', 'Maturity Start Date', 'Maturity End Date', 'Tenor', 'Tranche ESG Type']).to_excel(writer, sheet_name='Tranches', index=False)
        pd.DataFrame(columns=['Transaction Upload ID', 'Tranche Benchmark', 'Basis Point From', 'Basis Point To', 'Period From', 'Period To', 'Period Duration', 'Comment']).to_excel(writer, sheet_name='Tranche_Pricings', index=False)
        pd.DataFrame(columns=['Transaction Upload ID', 'Tranche Upload ID', 'Tranche Role Type', 'Company', 'Fund', 'Value', 'Percentage', 'Comment']).to_excel(writer, sheet_name='Tranche_Roles_Any', index=False)
    
    return destination_filename

# Streamlit app
st.title('Curating INFRA2 data files')

uploaded_file = st.file_uploader("Choose a source file", type=["xlsx"])

if uploaded_file is not None:
    # Save the uploaded file to a temporary directory
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
        temp_file.write(uploaded_file.getbuffer())
        temp_file_path = temp_file.name
    
    try:
        with st.spinner("Processing the file..."):
            destination_path = create_destination_file(temp_file_path)
        st.success("File processed successfully!")
        
        # Provide a download link for the processed file
        with open(destination_path, "rb") as file:
            st.download_button(
                label="Download Processed File",
                data=file,
                file_name=os.path.basename(destination_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"An error occurred: {e}")
    finally:
        # Clean up temporary files
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)
        if os.path.exists(destination_path):
            os.remove(destination_path)
else:
    st.info("Please upload an Excel file to start processing.")
