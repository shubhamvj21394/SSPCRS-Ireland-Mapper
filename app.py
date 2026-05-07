import streamlit as st
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
import datetime
import io

def precise_round(num):
    try:
        if pd.isna(num): return 0.0
        return float(Decimal(str(num)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))
    except:
        return 0.0

st.set_page_config(page_title="SSPCRS Ireland Mapper", layout="wide")

st.markdown("""
<style>
    .main {
        background-color: #f8fafc;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #10b981;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.title("🇮🇪 SSPCRS Data Transformation Tool")
st.info("Professional tool for mapping HSE reimbursement prices to Manufacturer and Retail values.")

st.sidebar.header("Data Inputs")
# FIX: Use file_uploader instead of file_input
source_file = st.sidebar.file_uploader("Upload Source Data", type=['csv', 'xlsx', 'xls'], help="Contains Code and Reimbursement Price")
ref_file = st.sidebar.file_uploader("Upload Reference Data", type=['csv', 'xlsx', 'xls'], help="Contains PRICE_ID, Item Type, and VAT")

if source_file and ref_file:
    if st.sidebar.button("Run Transformation Pipeline"):
        try:
            # Loading Source
            if source_file.name.endswith('.csv'):
                src_df = pd.read_csv(source_file)
            else:
                src_df = pd.read_excel(source_file)
            
            # Loading Reference
            if ref_file.name.endswith('.csv'):
                ref_df = pd.read_csv(ref_file)
            else:
                ref_df = pd.read_excel(ref_file)
            
            # Data Cleaning
            ref_df['PRICE_ID'] = ref_df['PRICE_ID'].astype(str).str.strip()
            src_df['Code'] = src_df['Code'].astype(str).str.strip()
            
            # Merge
            merged = pd.merge(src_df, ref_df, left_on='Code', right_on='PRICE_ID', how='left')
            
            # Calculations
            effective_date = datetime.date.today().replace(day=1).strftime("%Y-%m-%d")
            
            processed = pd.DataFrame()
            processed['PRICE_ID'] = merged['Code']
            processed['Brand Name'] = merged['Product Name'].fillna(merged['Code'])
            
            # Wholesale Price Extraction
            processed['Wholesale Price'] = merged['Reimbursement Price'].fillna(0).astype(float)
            
            # Metadata
            item_types = merged['Item Type'].fillna('Non-Fridge').astype(str)
            vats = merged['VAT'].fillna(0).astype(float)
            
            # Manufacturer Price
            processed['Manufacturer Price'] = merged.apply(
                lambda x: (float(x['Reimbursement Price']) / 1.12) if str(x['Item Type']) == 'Fridge' else (float(x['Reimbursement Price']) / 1.08),
                axis=1
            ).fillna(0).apply(precise_round)
            
            # Retail Calculations
            processed['Retail Price without VAT'] = (processed['Wholesale Price'] + 4.84).apply(precise_round)
            
            processed['Retail Price'] = processed.apply(
                lambda x: (x['Retail Price without VAT'] * 1.23) if merged.iloc[x.name]['VAT'] == 23 else x['Retail Price without VAT'],
                axis=1
            ).apply(precise_round)
            
            # Template Fillers
            processed['Multiplication Factor'] = "1"
            processed['Country'] = "IRELAND"
            processed['Currency'] = "EUR"
            processed['VAT'] = vats
            processed['Item Type'] = item_types
            processed['Effective Price Date'] = effective_date
            processed['Source Name'] = "HSE"
            
            st.success(f"Successfully processed {len(processed)} records.")
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Records", len(processed))
            col2.metric("Avg Retail", f"€{processed['Retail Price'].mean():.2f}")
            col3.metric(" fridge Items", len(processed[processed['Item Type'] == 'Fridge']))
            
            st.dataframe(processed, use_container_width=True)
            
            # Excel Export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                processed.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Download Result (Excel)",
                data=output.getvalue(),
                file_name=f"SSPCRS_Ireland_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Processing Error: {str(e)}")
else:
    st.warning("Please upload both source and reference files to start.")
