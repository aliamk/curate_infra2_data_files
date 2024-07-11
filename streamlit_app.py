import streamlit as st
import pandas as pd
from datetime import datetime
import os
import tempfile
import re
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# Replacement dictionary for specific replacements in 'Any Level Sectors'
replacement_dict_any_level_sectors = {
    'Renewables': 'Renewable Energy',
    'Social & Defence': 'Social Infrastructure',
    'Telecoms': 'Digital Infrastructure',
    'Airports': 'Airport',
    'Base Metals': 'Metal',
    'Bridges': 'Bridge',
    'Car Parks': 'Car Park',
    'Co Generation': 'Cogeneration Power',
    'Data Centres': 'Data Centre',
    'Transmission & Distribution': 'Transmission',
    'Distribution': 'Water Distribution',
    'District Heating': 'Heat Network',
    'Gas-Fired': 'Gas-Fired Power',
    'Manufacturing': 'Processing',
    'Maritime Transport': 'Waterway',
    'Minerals': 'Mineral',
    'Mobile': 'Tower',
    'Municipal': 'Municipal Building',
    'Nuclear': 'Nuclear Power',
    'Offshore Wind - Fixed': 'Wind (Offshore)',
    'Offshore Wind - Floating': 'Wind (Offshore)',
    'Onshore Wind': 'Wind (Onshore)',
    'Other Renewables': 'Renewable Energy',
    'Other Telecoms': 'Digital Infrastructure',
    'Other Transport': 'Transport',
    'Ports': 'Port',
    'Precious Metals': 'Metal',
    'Roads': 'Road',
    'Small Hydro': 'Hydro',
    'Smart Meters': '',
    'Solar PV - Floating': 'Solar (Floating PV)',
    'Solar - Floating': 'Solar (Floating PV)',
    'Solar PV': 'Solar (Land-Based PV)',
    'Solar Thermal': 'Solar (Thermal)',
    'Terrestrial': 'Digital Infrastructure',
    'Transit': 'Light Transport',
    'Treatment': 'Water Treatment',
    'Tunnels': 'Tunnel',
    'Waste-to-Energy': 'Waste to Energy',
    'Other Oil & Gas': 'Oil & Gas',
    'IWPP': '',
    'Other Water': 'Water',
    'Other Mining': 'Mining',
    'Other Social & Defence': 'Social Infrastructure',
    'Oil-fired': 'Oil-Fired Power',
    'Fire & Rescue': 'Social Infrastructure',
    'Street Lighting': 'Social Infrastructure',
    'Other Digital Infrastructure': 'Digital Infrastructure',
    'Agriculture': '',
    'Software': '',
    'Technology Processing': '',
    'Other Beyond Infrastructure': '',
    'Beyond Infra': ''
}

# Function to read the source file
def read_source_file(source_path):
    xl = pd.ExcelFile(source_path)
    df1 = xl.parse('Sheet1')
    df2 = xl.parse('Sheet2')
    return df1, df2

# Function to create the transaction DataFrame
def create_transaction_df(df1, df2):
    columns_mapping = {
        'Transaction Upload ID': 'Realfin INFRA Transaction Upload ID',
        'Transaction Name': 'Transaction Name',
        'Transaction Asset Class': 'Infrastructure',  # Column C has a fixed value
        'Transaction Status': 'Transaction Stage',
        'Finance Type': 'Finance Type',
        'Transaction Type': 'Transaction Type',
        'Unknown Asset': None,  # Column G is blank
        'Underlying Asset Configuration': None,  # Column H is blank
        'Transaction Local Currency': 'Transaction Currency',
        'Transaction Value (Local Currency)': 'Transaction Value (Local Currency m)',
        'Transaction Debt (Local Currency)': 'Transaction Debt (Local Currency m)',
        'Transaction Equity (Local Currency)': 'Transaction Equity (Local Currency m)',
        'Debt/Equity Ratio': 'Debt/Equity Ratio',
        'Underlying Number of Assets': None,  # Column N is blank
        'Region - Country': 'Transaction Country/Region',
        'Region - State': None,  # Column P is blank
        'Region - City': None,  # Column Q is blank
        'Any Level Sectors': ['Transaction Sector', 'Transaction Sub-sector'],
        'PPP': 'PPP',
        'Concession Period': 'Concession Period',
        'Contract': 'Contract',
        'SPV': None,  # Column V will be filled later
        'Active': 'True',  # Column W has a fixed value 'True'
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
    transaction_df['SPV'] = transaction_df['Transaction Upload ID'].map(spv_mapping)

    return transaction_df

# Function to clean up the Transaction Name column
def clean_transaction_name(transaction_df):
    transaction_df['Transaction Name'] = transaction_df['Transaction Name'].str.strip()  # Remove leading/trailing spaces
    transaction_df['Transaction Name'] = transaction_df['Transaction Name'].apply(lambda x: re.sub(r'\s+', ' ', x))  # Replace multiple spaces with single space
    return transaction_df

# Function to apply replacements based on a dictionary
def apply_replacements(df, column, replacements):
    def replace_value(cell_value):
        if isinstance(cell_value, str):
            for old, new in replacements.items():
                cell_value = cell_value.replace(old, new)
        return cell_value

    df[column] = df[column].apply(replace_value)

# Function to apply replacements based on a dictionary in a specific order
def apply_replacements_in_order(df, column, replacements):
    def replace_value(cell_value, replacements):
        if isinstance(cell_value, str):
            for old, new in replacements.items():
                cell_value = cell_value.replace(old, new)
        return cell_value

    # Step 1: Replace 'Other Beyond Infrastructure' first
    df[column] = df[column].apply(lambda x: replace_value(x, {'Other Beyond Infrastructure': replacements['Other Beyond Infrastructure']}))

    # Step 2: Replace the rest, including 'Beyond Infra'
    replacements.pop('Other Beyond Infrastructure')  # Remove 'Other Beyond Infrastructure' from the dictionary
    df[column] = df[column].apply(lambda x: replace_value(x, replacements))

# Apply specific replacements in 'Any Level Sectors' column with order consideration
def apply_specific_replacements(df, column, replacements):
    apply_replacements_in_order(df, column, replacements)

# Function to format date columns
def format_date_columns(df, date_columns):
    for col in date_columns:
        if (col in df.columns) and (df[col].dtype == 'object'):
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
    return df

# Function to autofit columns
def autofit_columns(writer):
    for sheetname in writer.sheets:
        worksheet = writer.sheets[sheetname]
        for col in worksheet.columns:
            max_length = 0
            column = get_column_letter(col[0].column)  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column].width = adjusted_width

# Function to create the destination file
def create_destination_file(source_path):
    df1, df2 = read_source_file(source_path)
    
    # Create transaction DataFrame
    transaction_df = create_transaction_df(df1, df2)
    
    # Clean up the Transaction Name column
    transaction_df = clean_transaction_name(transaction_df)

    # Function to replace words in Any Level Sectors columns
    def replace_words(cell_value):
        if isinstance(cell_value, str):
            # Step 1: Replace 'Coal-fired' with 'Xoal-Fired'
            cell_value = cell_value.replace('Coal-fired', 'Xoal-Fired')
            # Step 2: Replace 'Coal' with 'Mineral'
            cell_value = cell_value.replace('Coal', 'Mineral')
            # Step 3: Replace 'Other Power' with 'OtherConventionalEnergy' as a temporary placeholder
            cell_value = cell_value.replace('Other Power', 'OtherConventionalEnergy')
            # Step 4: Replace 'Power' with 'Conventional Energy' only if it's not part of 'Coal-Fired Power'
            cell_value = re.sub(r'\bPower\b', 'Conventional Energy', cell_value)
            # Step 5: Replace the temporary placeholder 'OtherConventionalEnergy' back to 'Conventional Energy'
            cell_value = cell_value.replace('OtherConventionalEnergy', 'Conventional Energy')
            # Step 6: Replace 'Xoal-Fired' with 'Coal-Fired Power'
            cell_value = cell_value.replace('Xoal-Fired', 'Coal-Fired Power')
            # Step 7: Replace 'Biofuels' with 'Biofuels/Biomass'
            cell_value = cell_value.replace('Biofuels', 'Biofuels/Biomass')
            # Step 8: Replace 'Biomass' with 'Biofuels/Biomass' only if 'Biofuels/Biomass' is not already in the cell
            if 'Biofuels/Biomass' not in cell_value:
                cell_value = cell_value.replace('Biomass', 'Biofuels/Biomass')

        return cell_value
    
    # Apply word replacement to 'Any Level Sectors' column in the transaction_df
    transaction_df['Any Level Sectors'] = transaction_df['Any Level Sectors'].apply(replace_words)
    
    # Apply specific replacements to 'Any Level Sectors' column in the transaction_df
    apply_specific_replacements(transaction_df, 'Any Level Sectors', replacement_dict_any_level_sectors)
    
    # Format date columns in transaction_df
    date_columns_transaction = ['Latest Transaction Event Date', 'Financial Close Date']
    transaction_df = format_date_columns(transaction_df, date_columns_transaction)
    
    # Generate the destination filename with a timestamp
    base, ext = os.path.splitext(source_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    destination_filename = f"curated_INFRA2_{timestamp}.xlsx"
    
    with pd.ExcelWriter(destination_filename, engine='openpyxl') as writer:
        # Write transaction data to the 'Transaction' sheet
        transaction_df.to_excel(writer, sheet_name='Transaction', index=False)
        
        # Create empty tabs with specified headers
        pd.DataFrame(columns=['Transaction Upload ID', 'Asset Upload ID']).to_excel(writer, sheet_name='Underlying_Asset', index=False)
        
        # Populate the Events tab with data from Source file (Sheet1)
        events_data = {
            'Transaction Upload ID': df1['Realfin INFRA Transaction Upload ID'],
            'Event Date': df1['Latest Transaction Event Date'],
            'Event Type': df1['Latest Transaction Event'],
            'Event Title': [None] * len(df1)  # Column D remains empty
        }
        events_df = pd.DataFrame(events_data)
        
        # Append the additional rows for Financial Close Date (Sheet1)
        additional_events_data = {
            'Transaction Upload ID': df1['Realfin INFRA Transaction Upload ID'],
            'Event Date': df1['Financial Close Date'],
            'Event Type': ['Financial Close'] * len(df1),  # Column C with 'Financial Close'
            'Event Title': [None] * len(df1)  # Column D remains empty
        }
        additional_events_df = pd.DataFrame(additional_events_data)
        
        # Append the additional rows for Transaction Announced Date (Sheet2)
        announced_events_data = {
            'Transaction Upload ID': df2['Realfin INFRA Transaction Upload ID'],
            'Event Date': df2['Transaction Announced Date'].replace('N/A', pd.NA),
            'Event Type': ['Announced'] * len(df2),  # Column C with 'Announced'
            'Event Title': [None] * len(df2)  # Column D remains empty
        }
        announced_events_df = pd.DataFrame(announced_events_data)
        
        # Append the additional rows for Transaction Request for Proposals Date (Sheet2)
        proposals_events_data = {
            'Transaction Upload ID': df2['Realfin INFRA Transaction Upload ID'],
            'Event Date': df2['Transaction Request For Proposals Date'].replace('N/A', pd.NA),
            'Event Type': ['Request for Proposals'] * len(df2),  # Column C with 'Request for Proposals'
            'Event Title': [None] * len(df2)  # Column D remains empty
        }
        proposals_events_df = pd.DataFrame(proposals_events_data)
        
        # Append the additional rows for Transaction Tender Launch Date (Sheet2)
        tender_events_data = {
            'Transaction Upload ID': df2['Realfin INFRA Transaction Upload ID'],
            'Event Date': df2['Transaction Tender Launch Date'].replace('N/A', pd.NA),
            'Event Type': ['Tender'] * len(df2),  # Column C with 'Tender'
            'Event Title': [None] * len(df2)  # Column D remains empty
        }
        tender_events_df = pd.DataFrame(tender_events_data)
        
        # Append the additional rows for Transaction Preferred Bidder Date (Sheet2)
        bidder_events_data = {
            'Transaction Upload ID': df2['Realfin INFRA Transaction Upload ID'],
            'Event Date': df2['Transaction Preferred Bidder Date'].replace('N/A', pd.NA),
            'Event Type': ['Preferred Bidder'] * len(df2),  # Column C with 'Preferred Bidder'
            'Event Title': [None] * len(df2)  # Column D remains empty
        }
        bidder_events_df = pd.DataFrame(bidder_events_data)
        
        # Concatenate all data
        full_events_df = pd.concat([
            events_df,
            additional_events_df,
            announced_events_df,
            proposals_events_df,
            tender_events_df,
            bidder_events_df
        ], ignore_index=True)
        
        # Format date columns in events_df
        date_columns_events = ['Event Date']
        full_events_df = format_date_columns(full_events_df, date_columns_events)
        
        # Remove rows where 'Event Date' is blank or 'N/A'
        full_events_df = full_events_df.dropna(subset=['Event Date'])
        full_events_df = full_events_df[full_events_df['Event Date'] != 'N/A']

        # Apply replacements to 'Event Type'
        replacements_event_type = {
            'Best And Final Offer': 'Best and Final Offer',
            'Next Milestone': '',
            'Undisclosed Financial Close': '',
            'Financial Close Transaction': 'Financial Close',
            'General Announcement': '',
            'Risk Alert': '',
            'Adviser Mandate Won': 'Adviser Appointed',
            'Tender Launch': 'Tender',
            'Request for Qualification': 'Request for Qualifications',
            'Bank Market Approach': 'Financing Sought',
            'Transaction Announced': 'Announced',
            'Bank Mandate Won': 'Lenders Appointed',
            'EoI (Expression of Interest)': 'Expression of Interest',
            'Offtake Agreement Signed': 'Offtake Agreement',
            'Concession Signed': 'Concession Agreement',
            'Financing Signed': 'Financing Agreement',
            'RoI (Request for Information)': 'Request for Information',
            'Sponsor withdrawal': ''
        }
        apply_replacements(full_events_df, 'Event Type', replacements_event_type)

        # Remove rows where 'Event Type' is blank
        full_events_df = full_events_df[full_events_df['Event Type'] != '']

        # Remove duplicate rows
        full_events_df = full_events_df.drop_duplicates()

        full_events_df.to_excel(writer, sheet_name='Events', index=False)

        # Populate the Bidders Any tab
        role_bidders_data = {
            'Transaction Upload ID': df2['Realfin INFRA Transaction Upload ID'],
            'Role Type': df2['Transaction Role'].replace('N/A', pd.NA),
            'Role Subtype': None,  # Column C remains empty
            'Company': df2['Company Name'].replace('N/A', pd.NA),
            'Fund': None,  # Column E remains empty
            'Bidder Status': 'Successful',  # Column F with 'Successful'
            'Client Counterparty': df2['Advise To'].replace('N/A', pd.NA),
            'Client Company Name': df2['Company Advised (Client Company)'].replace('N/A', pd.NA),
            'Fund Name': None  # Column I remains empty
        }
        bidders_any_df = pd.DataFrame(role_bidders_data)
        
        # Apply replacements to 'Role Type'
        replacements_role_type = {
            'O&M': 'Operations & Maintenance'
        }
        apply_replacements(bidders_any_df, 'Role Type', replacements_role_type)

        # Apply replacements to 'Client Counterparty'
        replacements_client_counterparty = {
            'AwardingAuthority': 'Awarding Authority'
        }
        apply_replacements(bidders_any_df, 'Client Counterparty', replacements_client_counterparty)
        
        # Remove rows where 'Role Type' is blank, 'N/A', or 'Other'
        bidders_any_df = bidders_any_df.dropna(subset=['Role Type'])
        bidders_any_df = bidders_any_df[~bidders_any_df['Role Type'].str.contains('N/A|^$|Other')]
        
        # Arrange columns to match the required output for Bidders_Any tab
        bidders_any_columns = ['Transaction Upload ID', 'Role Type', 'Role Subtype', 'Company', 'Fund', 'Bidder Status', 'Client Counterparty', 'Client Company Name', 'Fund Name']
        bidders_any_df = bidders_any_df.reindex(columns=bidders_any_columns)
        
        bidders_any_df.to_excel(writer, sheet_name='Bidders Any', index=False)

        # Populate the Tranches tab
        tranches_data = {
            'Transaction Upload ID': df2.get('Realfin INFRA Transaction Upload ID'),
            'Tranche Upload ID': df2.get('Realfin INFRA Tranche Upload ID'),
            'Tranche Primary Type': df2.get('Tranche Instrument Primary Type'),
            'Tranche Secondary Type': df2.get('Tranche Instrument Secondary Type'),
            'Tranche Tertiary Type': df2.get('Tranche Instrument Tertiary Type'),
            'Helper_Tranche Name': df2.get('Tranche Name'),
            'Helper_Tranche Value $': df2.get('Tranche Value ($m)'),
            'Helper_Transaction Value (USD m)': df2.get('Transaction Value (USD m)'),
            'Helper_Transaction Value (LC m)': df2.get('Transaction Value (Local Currency m)'),
            'Maturity Start Date': df2.get('Tranche Maturity Start Date'),
            'Maturity End Date': df2.get('Tranche Maturity End Date'),
            'Tenor': df2.get('Tranche Maturity Duration (Years)')
        }
        tranches_df = pd.DataFrame(tranches_data)
        
        # Apply replacements to 'Tranche Secondary Type'
        replacements_tranche_secondary_type = {
            'Loans': 'Loan',
            'IFI Government Support': 'Non-Commercial Instrument',
            'Bonds': 'Bond'
        }
        apply_replacements(tranches_df, 'Tranche Secondary Type', replacements_tranche_secondary_type)

        # Apply replacements to 'Tranche Tertiary Type'
        replacements_tranche_tertiary_type = {
            'Cash Equity': 'Equity',
            'Revolver': 'Revolving Credit Facility',
            'Credit Facility': '',
            'Bridge Facility': 'Bridge',
            'Green Bond': '',
            'Green Loan': '',
            'Sustainability-linked Loan': '',
            'Working Capital': 'Working Capital Facility',
            'Government Loan': 'State Loan',
            'Sustainability-linked Bond': '',
            'Mezzanine Debt': 'Mezzanine',
            'Islamic Loan': '',
            'Islamic Bond': ''
        }
        apply_replacements(tranches_df, 'Tranche Tertiary Type', replacements_tranche_tertiary_type)
        
        # Populate 'Tranche ESG Type' based on 'Helper_Tranche Name'
        esg_mapping_name = {
            'Islamic': 'Sharia-Compliant',
            'sharia': 'Sharia-Compliant',
            'sukuk': 'Sharia-Compliant',
            'green': 'Green',
            'sustainab': 'Sustainability-Linked',
            'social': 'Social',
            'blue': 'Blue'
        }
        for keyword, esg_type in esg_mapping_name.items():
            tranches_df.loc[tranches_df['Helper_Tranche Name'].str.contains(keyword, case=False, na=False), 'Tranche ESG Type'] = esg_type
        
        # Populate 'Tranche ESG Type' based on 'Tranche Tertiary Type'
        esg_mapping_tertiary = {
            'Sustainability-linked Loan': 'Sustainability-Linked',
            'Sustainability-linked Bond': 'Sustainability-Linked',
            'Green Loan': 'Green',
            'Green Bond': 'Green',
            'Islamic Loan': 'Sharia-Compliant',
            'Islamic Bond': 'Sharia-Compliant'
        }
        for keyword, esg_type in esg_mapping_tertiary.items():
            tranches_df.loc[tranches_df['Tranche Tertiary Type'].str.contains(keyword, case=False, na=False), 'Tranche ESG Type'] = esg_type
        
        # Format date columns in tranches_df
        date_columns_tranches = ['Maturity Start Date', 'Maturity End Date']
        tranches_df = format_date_columns(tranches_df, date_columns_tranches)
        
        # Calculate 'Helper_Tranche Value $ as % of Transaction Value USD m'
        tranches_df['Helper_Tranche Value $ as % of Transaction Value USD m'] = tranches_df['Helper_Tranche Value $'] / tranches_df['Helper_Transaction Value (USD m)']

        # Populate 'Value' column based on calculated percentage
        tranches_df['Value'] = tranches_df['Helper_Tranche Value $ as % of Transaction Value USD m'] * tranches_df['Helper_Transaction Value (LC m)']

        # Arrange columns to match the required output for Tranches tab
        tranches_columns = [
            'Transaction Upload ID', 'Tranche Upload ID', 'Tranche Primary Type', 'Tranche Secondary Type', 'Tranche Tertiary Type', 
            'Value', 'Maturity Start Date', 'Maturity End Date', 'Tenor', 'Tranche ESG Type', 
            'Helper_Tranche Name', 'Helper_Tranche Value $', 'Helper_Transaction Value (USD m)', 
            'Helper_Transaction Value (LC m)', 'Helper_Tranche Value $ as % of Transaction Value USD m'
        ]
        tranches_df = tranches_df.reindex(columns=tranches_columns)
        
        tranches_df.to_excel(writer, sheet_name='Tranches', index=False)

        # Populate the Tranche_Pricings tab
        tranche_pricings_data = {
            'Tranche Upload ID': df2.get('Realfin INFRA Tranche Upload ID'),
            'Tranche Benchmark': df2.get('Tranche Loan Reference Rate'),
            'Basis Point From': df2.get('Range From'),
            'Basis Point To': df2.get('Range To'),
            'Period From': None,  # Column E remains empty
            'Period To': None,  # Column F remains empty
            'Period Duration': None,  # Column G remains empty
            'Comment': None  # Column H remains empty
        }
        tranche_pricings_df = pd.DataFrame(tranche_pricings_data)
        
        # Arrange columns to match the required output for Tranche_Pricings tab
        tranche_pricings_columns = ['Tranche Upload ID', 'Tranche Benchmark', 'Basis Point From', 'Basis Point To', 'Period From', 'Period To', 'Period Duration', 'Comment']
        tranche_pricings_df = tranche_pricings_df.reindex(columns=tranche_pricings_columns)
        
        tranche_pricings_df.to_excel(writer, sheet_name='Tranche_Pricings', index=False)

        # Populate the Tranche_Roles_Any tab
        tranche_roles_any_data = {
            'Transaction Upload ID': df2['Realfin INFRA Transaction Upload ID'],
            'Tranche Upload ID': df2['Realfin INFRA Tranche Upload ID'],
            'Tranche Role Type': df2['Tranche Role'],
            'Company': df2['Company Name'],
            'Fund': None,  # Column E remains empty
            'Value': None,  # Column F remains empty
            'Percentage': None,  # Column G remains empty
            'Comment': None,  # Column H remains empty,
            'Helper_Tranche Primary Type': df2['Tranche Instrument Primary Type'],
            'Helper_Tranche Value $': df2['Tranche Value ($m)'],
            'Helper_Transaction Value (USD m)': df2['Transaction Value (USD m)'],
            'Helper_LT Accredited Value ($m)': df2['LT Accredited Value ($m)'],
            'Helper_Sponsor Equity USD m': df2['Sponsor Equity (USDm)']
        }
        
        tranche_roles_any_df = pd.DataFrame(tranche_roles_any_data)
        
        # Apply replacements and updates to 'Tranche Role Type' based on 'Helper_Tranche Primary Type'
        tranche_roles_any_df['Tranche Role Type'] = tranche_roles_any_df.apply(
            lambda row: 'Sponsor' if row['Helper_Tranche Primary Type'] == 'Equity' and row['Tranche Role Type'] in ['Fund', 'Multilateral', 'Export Credit Agency', 'State Lender', 'Public Finance Institution', 'Institutional Investor', 'International Finance Institution'] else 
                        'Debt Provider' if row['Helper_Tranche Primary Type'] == 'Debt' and row['Tranche Role Type'] in ['Fund', 'Multilateral', 'Export Credit Agency', 'State Lender', 'Public Finance Institution', 'Institutional Investor', 'International Finance Institution', 'Development Equity'] else 
                        row['Tranche Role Type'], axis=1
        )

        replacements_tranche_role_type = {
            'MLA': 'Mandated Lead Arranger',
            'Participant': 'Debt Provider'
        }
        apply_replacements(tranche_roles_any_df, 'Tranche Role Type', replacements_tranche_role_type)
        
        # Copy 'Value' from 'Tranches' tab to 'Helper_Tranche_Value LC' in 'Tranche_Roles_Any' tab
        tranches_df = tranches_df.set_index('Tranche Upload ID')
        tranche_roles_any_df = tranche_roles_any_df.set_index('Tranche Upload ID')
        tranche_roles_any_df['Helper_Tranche_Value LC'] = tranches_df['Value']
        tranche_roles_any_df = tranche_roles_any_df.reset_index()

        # Ensure 'Helper_LT Accredited Value ($m)' column exists and is properly referenced
        if 'Helper_LT Accredited Value ($m)' in tranche_roles_any_df.columns:
            tranche_roles_any_df['Helper_LT Accredited Value ($m) as % of Helper_Tranche Value $'] = tranche_roles_any_df['Helper_LT Accredited Value ($m)'] / tranche_roles_any_df['Helper_Tranche Value $']
            tranche_roles_any_df['Helper_Debt Provider Underwriting Value LC'] = tranche_roles_any_df['Helper_LT Accredited Value ($m) as % of Helper_Tranche Value $'] * tranche_roles_any_df['Helper_Tranche_Value LC']
        
        # Create new columns and populate them with calculated values
        tranche_roles_any_df['Helper_Sponsor Equity $ as % of Helper_Tranche Value $'] = tranche_roles_any_df['Helper_Sponsor Equity USD m'] / tranche_roles_any_df['Helper_Tranche Value $']
        tranche_roles_any_df['Helper_Sponsor Equity LC'] = tranche_roles_any_df['Helper_Sponsor Equity USD m'] * tranche_roles_any_df['Helper_Tranche Value $']

        # Populate the 'Value' column based on conditions
        tranche_roles_any_df['Value'] = tranche_roles_any_df.apply(
            lambda row: row['Helper_Sponsor Equity LC'] if row['Helper_Tranche Primary Type'] == 'Equity' else
                        (row['Helper_Debt Provider Underwriting Value LC'] if row['Helper_Tranche Primary Type'] == 'Debt' else None),
            axis=1
        )
        
        # Arrange columns to match the required output for Tranche_Roles_Any tab
        tranche_roles_any_columns = [
            'Transaction Upload ID', 'Tranche Upload ID', 'Tranche Role Type', 'Company', 'Fund', 'Value', 'Percentage', 'Comment',
            'Helper_Tranche Primary Type', 'Helper_Tranche Value $', 'Helper_Transaction Value (USD m)', 'Helper_LT Accredited Value ($m)', 'Helper_Sponsor Equity USD m'
        ]
        tranche_roles_any_df = tranche_roles_any_df.reindex(columns=tranche_roles_any_columns)
        
        tranche_roles_any_df.to_excel(writer, sheet_name='Tranche_Roles_Any', index=False)

        # Apply word replacements to specified columns in the 'Transaction' tab
        replacements_transaction_name = {
            'Additional Facility': 'Additional Financing',
            'Bond Facility': 'Bond',
            ' and ': ' & ',
            ' Cancelled': '',
            'Acquisition of a Minority Stake in ': '',
            'Acquisition of a Majority Stake in': '',
            'Acquisition of a ': '',
            'Acquisition of ': '',
            'Acquisiition of ': '',
            'Acquisiton of ': '',
            'Acquisiion of ': '',
            'Acquisistion of ': '',
            'Acqusition of ': ''
        }
        apply_replacements(transaction_df, 'Transaction Name', replacements_transaction_name)
        
        replacements_transaction_status = {
            'Financial close': 'Financial Close',
            'Pre-financing': 'Preparation'
        }
        apply_replacements(transaction_df, 'Transaction Status', replacements_transaction_status)
        
        replacements_finance_type = {
            'Corporate Finance': 'Corporate',
            'Non-Commercial Finance': 'Non-Commercial',
            'Project Finance': 'Limited-Recourse',
            'Design-Build': 'Corporate',
            'Public Sector Finance': 'Non-Commercial'
        }
        apply_replacements(transaction_df, 'Finance Type', replacements_finance_type)
        
        replacements_transaction_type = {
            'Asset acquisition': 'Asset Acquisition',
            'Company acquisition': 'Corporate Acquisition',
            'Additional Facility': 'Additional Financing'
        }
        apply_replacements(transaction_df, 'Transaction Type', replacements_transaction_type)
        
        replacements_region_country = {
            'China - Chinese Taipei': 'Taiwan',
            'China - Hong Kong (SAR)': 'Hong Kong',
            'China - Mainland': 'China',
            'Cook Islands': '',
            'Fiji Islands': '',
            'Marshall Islands': '',
            'Myanmar (Burma)': 'Myanmar',
            'Timor-Leste (East Timor)': 'Timor-Leste',
            'Tonga': '',
            'Virgin Islands (US)': 'US Virgin Islands',
            'Hong Kong (SAR)': 'Hong Kong',
            'Mainland': 'China',
            'Chinese Taipei': 'Taiwan',
            'Macau (SAR)': 'Macau'
        }
        apply_replacements(transaction_df, 'Region - Country', replacements_region_country)
        
        replacements_contract = {
            'Unknown': ''
        }
        apply_replacements(transaction_df, 'Contract', replacements_contract)
        
        # Write the updated transaction_df again to reflect changes in specified columns
        transaction_df.to_excel(writer, sheet_name='Transaction', index=False)
        
        # Autofit columns for all sheets
        autofit_columns(writer)
    
    return destination_filename

# Streamlit app
st.title('Curating INFRA 2 data files')

uploaded_file = st.file_uploader("Choose a source file", type=["xlsx"])

# Close the temporary file before further processing ( temp_file_path = temp_file.name )
if uploaded_file is not None:
    # Save the uploaded file to a temporary directory
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
        temp_file.write(uploaded_file.getbuffer())
        temp_file_path = temp_file.name
    
    destination_path = None
    
    try:
        with st.spinner("Processing the file..."):
            destination_path = create_destination_file(temp_file_path)
        st.success("File processed successfully!")

        # Provide a download button for the processed file
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
        if temp_file_path:
            temp_file.close()
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
        if destination_path and os.path.exists(destination_path):
            os.remove(destination_path)

else:
    st.info("Please upload an Excel file to start processing.")
