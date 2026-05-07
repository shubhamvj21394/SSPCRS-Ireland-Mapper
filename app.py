import streamlit as st
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
import datetime
import io
import re

# Precise rounding to 2 decimal places (Round Half Up)
def precise_round(num):
    try:
        if num is None or pd.isna(num):
            return 0.0
        return float(Decimal(str(num)).quantize(Decimal('1.00'), rounding=ROUND_HALF_UP))
    except (ValueError, TypeError, ArithmeticError):
        return 0.0

def clean_price(val):
    if val is None or pd.isna(val):
        return 0.0
    s = str(val)
    s = re.sub(r'[^0-9.]', '', s)
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0

st.set_page_config(page_title="SSPCRS Ireland Mapper", layout="wide", page_icon="🇮🇪")

# Styling
st.markdown("""
<style>
    [data-testid="stMetricValue"] { font-size: 2rem !important; font-weight: 700; color: #0F172A; }
    .main-header { font-size: 2rem; font-weight: 300; color: #1e293b; border-bottom: 2px solid #e2e8f0; padding-bottom: 1rem; margin-bottom: 2rem; }
    .main-header b { font-weight: 700; }
    .stButton>button { background-color: #10b981; color: white; border: none; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">Ireland HSE <b>SSPCRS Transformer</b></div>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.title("Configuration")
    source_file = st.file_uploader("1. Source Data (Ireland_Source_Data)", type=['csv', 'xlsx', 'xls'])
    ref_file = st.file_uploader("2. Reference Data (Reference_Data)", type=['csv', 'xlsx', 'xls'])
    st.divider()
    process_btn = st.button("🚀 PROCESS DATABASE", use_container_width=True)

if source_file and ref_file:
    if process_btn:
        try:
            # Data Loading
            src_df = pd.read_csv(source_file) if source_file.name.endswith('.csv') else pd.read_excel(source_file)
            ref_df = pd.read_csv(ref_file) if ref_file.name.endswith('.csv') else pd.read_excel(ref_file)
            
            # Standardize Identifiers
            src_df['Code_Standard'] = src_df['Code'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            ref_df['PRICE_ID_Standard'] = ref_df['PRICE_ID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            
            # Pivot Merge
            merged = pd.merge(src_df, ref_df, left_on='Code_Standard', right_on='PRICE_ID_Standard', how='left')
            effective_date = datetime.date.today().replace(day=1).strftime("%Y-%m-%d")
            
            # Mapping logic
            results = []
            for _, row in merged.iterrows():
                wholesale = clean_price(row.get('Reimbursement Price', 0))
                vat = clean_price(row.get('VAT', 0))
                item_type = str(row.get('Item Type', 'Non-Fridge'))
                
                mfg_factor = 1.12 if item_type == 'Fridge' else 1.08
                mfg_price = precise_round(wholesale / mfg_factor)
                net_retail = precise_round(wholesale + 4.84)
                retail_price = precise_round(net_retail * 1.23) if vat == 23 else net_retail
                
                results.append({
                    "PRICE_ID": row['Code_Standard'],
                    "Wholesale Price": precise_round(wholesale),
                    "Manufacturer Price": mfg_price,
                    "Retail Price without VAT": net_retail,
                    "Retail Price": retail_price,
                    "VAT": int(vat),
                    "Item Type": item_type,
                    "Brand Name": row.get('Name', ''),
                    "Active Ingredient": row.get('INN', ''),
                    "Effective Price Date": effective_date,
                    "Country": "IRELAND",
                    "Currency": "EUR",
                    "Source Name": "HSE"
                })
            
            final_df = pd.DataFrame(results)
            
            # Analytics
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Records", len(final_df))
            c2.metric("Coverage", f"{(final_df['Item Type'].count()/len(final_df)*100):.1f}%")
            c3.metric("Avg Retail", f"€{final_df['Retail Price'].mean():.2f}")
            c4.metric("Fridge Items", len(final_df[final_df['Item Type']=='Fridge']))
            
            st.subheader("📊 Manufacturer Price Distribution")
            bins = [0, 10, 50, 100, 500, 1000]
            labels = ['€0-10', '€10-50', '€50-100', '€100-500', '€500+']
            final_df['Range'] = pd.cut(final_df['Manufacturer Price'], bins=bins + [float('inf')], labels=labels)
            st.bar_chart(final_df['Range'].value_counts().reindex(labels))
            
            st.dataframe(final_df, use_container_width=True)
            
            # Export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False)
            st.sidebar.download_button("📥 DOWNLOAD XLSX", output.getvalue(), "SSPCRS_Ireland_Master.xlsx", use_container_width=True)
            
        except Exception as e:
            st.error(f"Arithmetic Error: {str(e)}")
