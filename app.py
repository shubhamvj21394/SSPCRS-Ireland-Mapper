import streamlit as st
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
import datetime
import io

def precise_round(num):
    try:
        return float(Decimal(str(num)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))
    except:
        return 0.0

st.set_page_config(page_title="SSPCRS Transformer", layout="wide")

st.sidebar.title("SSPCRS Data Processor")
source_file = st.sidebar.file_input("Upload Source Data (CSV/XLSX)")
ref_file = st.sidebar.file_input("Upload Reference Data (CSV/XLSX)")

if source_file and ref_file:
    if st.sidebar.button("Process Data"):
        # Loading
        if source_file.name.endswith('.csv'):
            src_df = pd.read_csv(source_file)
        else:
            src_df = pd.read_excel(source_file)
            
        if ref_file.name.endswith('.csv'):
            ref_df = pd.read_csv(ref_file)
        else:
            ref_df = pd.read_excel(ref_file)
            
        # Logic
        ref_df['PRICE_ID'] = ref_df['PRICE_ID'].astype(str).str.strip()
        src_df['Code'] = src_df['Code'].astype(str).str.strip()
        
        merged = pd.merge(src_df, ref_df, left_on='Code', right_on='PRICE_ID', how='left')
        
        # Calculations
        effective_date = datetime.date.today().replace(day=1).strftime("%Y-%m-%d")
        
        processed = pd.DataFrame()
        processed['PRICE_ID'] = merged['Code']
        processed['Wholesale Price'] = merged['Reimbursement Price'].fillna(0).astype(float)
        processed['Item Type'] = merged['Item Type'].fillna('Non-Fridge')
        processed['VAT'] = merged['VAT'].fillna(0).astype(float)
        
        processed['Manufacturer Price'] = processed.apply(
            lambda x: x['Wholesale Price'] / 1.12 if x['Item Type'] == 'Fridge' else x['Wholesale Price'] / 1.08, 
            axis=1
        ).apply(precise_round)
        
        processed['Retail Price without VAT'] = (processed['Wholesale Price'] + 4.84).apply(precise_round)
        processed['Retail Price'] = processed.apply(
            lambda x: x['Retail Price without VAT'] * 1.23 if x['VAT'] == 23 else x['Retail Price without VAT'],
            axis=1
        ).apply(precise_round)
        
        processed['Multiplication Factor'] = "1"
        processed['Country'] = "IRELAND"
        processed['Currency'] = "EUR"
        processed['Effective Price Date'] = effective_date
        processed['Source Name'] = "HSE"
        
        st.success("Processing complete!")
        st.dataframe(processed)
        
        # Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            processed.to_excel(writer, index=False)
        st.download_button("Download Output as Excel", data=output.getvalue(), file_name="SSPCRS_Processed.xlsx")
