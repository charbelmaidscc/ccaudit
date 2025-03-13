import streamlit as st
import pandas as pd
import io
import xlsxwriter  # Ensure xlsxwriter is explicitly imported
from datetime import datetime

def preprocess_cca(df):
    df['Start Of Contract'] = pd.to_datetime(df['Start Of Contract'], errors='coerce').dt.strftime('%m/%d/%Y')
    df['Contract'] = df['Contract'].astype(str).str.split('.').str[0]  # Ensure contract is clean without decimals
    return df

def preprocess_hp(df):
    df['Contract Name'] = df['Contract Name'].astype(str).str.replace("Contr-", "", regex=False).str.strip()
    return df

def preprocess_pt(df):
    df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce').dt.strftime('%m/%d/%Y')
    df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce').dt.strftime('%m/%d/%Y')
    return df

def preprocess_ec(df):
    df['Cont #'] = df['Cont #'].astype(str).str.split('.').str[0]  # Ensure consistency in contract numbers
    return df

def add_columns(cca, hp, ec, pt, month_start_date):
    if cca.empty or hp.empty or ec.empty or pt.empty:
        return pd.DataFrame()  # Return an empty DataFrame if any input is empty

    hp_filtered = hp[(hp['Status'] == 'WITH_CLIENT') & (hp['Type Of maid'] == 'CC')].copy()
    hp_filtered['Contract Name'] = hp_filtered['Contract Name'].astype(str).str.strip()
    cca['Contract'] = cca['Contract'].astype(str).str.strip()

    # Ensure valid To Check column
    cca['To Check'] = cca['Contract'].apply(lambda x: 'Yes' if x in hp_filtered['Contract Name'].tolist() else 'No')
    
    # If all rows have "No" in To Check, something is wrong
    if (cca['To Check'] == 'No').all():
        st.warning("Warning: No contracts matched in 'To Check' column. Verify contract formats.")
    
    cca['Exceptional Case'] = cca['Contract'].apply(lambda x: 'Yes' if x in ec['Cont #'].tolist() else 'No')
    
    def check_paying_now(row):
        if row['To Check'] == 'No':
            return ''
        if row['Exceptional Case'] == 'Yes':
            ec_value = ec.loc[ec['Cont #'] == row['Contract'], 'Monthly Payment'].values
            if ec_value.size > 0:
                try:
                    ec_amount = pd.to_numeric(ec_value[0], errors='coerce')
                    if pd.isna(ec_amount):
                        return 'Yes'
                    return 'Yes' if row['Amount Of Payment'] >= ec_amount else 'No'
                except Exception:
                    return 'No'
        pt_value = pt[(pt['Nationality'] == row['Maid Nationality']) & (pt['Contract Type'] == row['Contract Type'])]
        if not pt_value.empty:
            latest_price = pd.to_numeric(pt_value.loc[pt_value['End Date'].idxmax(), 'Minimum monthly payment + VAT'], errors='coerce')
            return 'Yes' if row['Amount Of Payment'] >= latest_price else 'No'
        return ''
    
    cca['Paying Correctly on Price of Now'] = cca.apply(check_paying_now, axis=1)
    
    def check_paying_contract_start(row):
        if row['To Check'] == 'No' or row['Exceptional Case'] == 'Yes':
            return ''
        if row['Paying Correctly on Price of Now'] == 'No':
            pt_value = pt[(pt['Nationality'] == row['Maid Nationality']) & (pt['Contract Type'] == row['Contract Type']) & ((pt['Start Date'] <= row['Start Of Contract']) & (pt['End Date'] >= row['Start Of Contract']))]
            if not pt_value.empty:
                latest_price = pt_value['Minimum monthly payment + VAT'].max()
                return 'Yes' if row['Amount Of Payment'] >= latest_price else 'No'
        return ''
    
    cca['Paying Correctly on Price of Contract Start Date'] = cca.apply(check_paying_contract_start, axis=1)
    
    return cca

def main():
    st.title("Client’s Contract Audit Processing")
    
    month_start_date = st.date_input("Month Start Date", value=datetime.today()).strftime('%m/%d/%Y')
    
    hp_file = st.file_uploader("Upload Housemaid Payroll", type=["xls", "xlsx"], key="hp")
    cca_file = st.file_uploader("Upload Client’s Contract Audit", type=["xls", "xlsx"], key="cca")
    ec_file = st.file_uploader("Upload Exceptional Cases", type=["xls", "xlsx"], key="ec")
    pt_file = st.file_uploader("Upload Price Trends", type=["xls", "xlsx"], key="pt")
    
    if st.button("Generate"):
        if hp_file and cca_file and ec_file and pt_file:
            hp = preprocess_hp(pd.read_excel(io.BytesIO(hp_file.getvalue()), engine='openpyxl'))
            cca = preprocess_cca(pd.read_excel(io.BytesIO(cca_file.getvalue()), engine='openpyxl'))
            ec = preprocess_ec(pd.read_excel(io.BytesIO(ec_file.getvalue()), engine='openpyxl'))
            pt = preprocess_pt(pd.read_excel(io.BytesIO(pt_file.getvalue()), engine='openpyxl'))
            
            labeled_cca = add_columns(cca, hp, ec, pt, month_start_date)
            
            # Debugging check
            if labeled_cca is None or labeled_cca.empty:
                st.error("Error: Processed DataFrame is empty. Please check the input data.")
                return
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                labeled_cca.to_excel(writer, index=False, sheet_name="Labeled CCA")
            output.seek(0)
            
            st.download_button("Download Labeled Client’s Contract Audit", data=output, file_name="Labeled_CCA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("Please upload all required files before generating the output.")

if __name__ == "__main__":
    main()
