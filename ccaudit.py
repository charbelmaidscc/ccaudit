import streamlit as st
import pandas as pd
import io
from datetime import datetime

def preprocess_cca(df):
    df['Start Of Contract'] = df['Start Of Contract'].astype(str).str.rstrip()
    return df

def preprocess_hp(df):
    df['Contract Name'] = df['Contract Name'].astype(str).str[6:]
    return df

def add_columns(cca, hp, ec, pt, month_start_date):
    cca['To Check'] = cca['Contract'].apply(lambda x: 'Yes' if x in hp[(hp['Status'] == 'WITH_CLIENT') & (hp['Type Of maid'] == 'CC')]['Contract Name'].tolist() else 'No')
    cca['Exceptional Case'] = cca['Contract'].apply(lambda x: 'Yes' if x in ec['Cont #'].tolist() else 'No')
    
    def check_paying_now(row):
        if row['To Check'] == 'No':
            return ''
        if row['Exceptional Case'] == 'Yes':
            ec_value = ec.loc[ec['Cont #'] == row['Contract'], 'Monthly Payment'].values
            return 'Yes' if (ec_value.size > 0 and (ec_value[0] in ['N/A', '-'] or row['Amount Of Payment'] >= ec_value[0])) else 'No'
        pt_value = pt[(pt['Nationality'] == row['Maid Nationality']) & (pt['Contract Type'] == row['Contract Type'])]['Minimum monthly payment + VAT'].max()
        return 'Yes' if row['Amount Of Payment'] >= pt_value else 'No'
    
    cca['Paying Correctly on Price of Now'] = cca.apply(check_paying_now, axis=1)
    
    def check_paying_contract_start(row):
        if row['To Check'] == 'No' or row['Exceptional Case'] == 'Yes':
            return ''
        if row['Paying Correctly on Price of Now'] == 'No':
            pt_value = pt[(pt['Nationality'] == row['Maid Nationality']) & (pt['Contract Type'] == row['Contract Type']) & (pt['Start Date'] <= row['Start Of Contract']) & (pt['End Date'] >= row['Start Of Contract'])]['Minimum monthly payment + VAT'].max()
            return 'Yes' if row['Amount Of Payment'] >= pt_value else 'No'
        return ''
    
    cca['Paying Correctly on Price of Contract Start Date'] = cca.apply(check_paying_contract_start, axis=1)
    
    def check_upgrading_nationality(row):
        if row['To Check'] == 'No' or row['Exceptional Case'] == 'Yes':
            return ''
        if row['Paying Correctly on Price of Now'] == 'No' and row['Paying Correctly on Price of Contract Start Date'] == 'No':
            if pd.isna(row['Upgrading Nationality Payment Amount']):
                return 'No'
            pt_value = pt[(pt['Nationality'] == row['Maid Nationality']) & (pt['Contract Type'] == row['Contract Type'])]['Minimum monthly payment + VAT'].max()
            return 'Yes' if (row['Amount Of Payment'] + row['Upgrading Nationality Payment Amount']) >= pt_value else 'No'
        return ''
    
    cca['Paying Correctly if Upgrading Nationality'] = cca.apply(check_upgrading_nationality, axis=1)
    
    def check_pro_rated(row):
        if row['To Check'] == 'No' or row['Exceptional Case'] == 'Yes':
            return ''
        if row['Paying Correctly on Price of Now'] == 'No' and row['Paying Correctly on Price of Contract Start Date'] == 'No' and row['Paying Correctly if Upgrading Nationality'] == 'No':
            if row['Start Of Contract'] < month_start_date:
                return 'No'
            return 'Yes' if row['Amount Of Payment'] >= row['Pro-Rated'] else 'No'
        return ''
    
    cca['Paying Correctly if Pro-Rated Value'] = cca.apply(check_pro_rated, axis=1)
    
    return cca

def main():
    st.title("Client’s Contract Audit Processing")
    
    month_start_date = st.date_input("Month Start Date", value=datetime.today())
    
    hp_file = st.file_uploader("Upload Housemaid Payroll", type=["xls", "xlsx"], key="hp")
    cca_file = st.file_uploader("Upload Client’s Contract Audit", type=["xls", "xlsx"], key="cca")
    ec_file = st.file_uploader("Upload Exceptional Cases", type=["xls", "xlsx"], key="ec")
    pt_file = st.file_uploader("Upload Price Trends", type=["xls", "xlsx"], key="pt")
    
    if st.button("Generate"):
        if hp_file and cca_file and ec_file and pt_file:
            hp = preprocess_hp(pd.read_excel(hp_file))
            cca = preprocess_cca(pd.read_excel(cca_file))
            ec = pd.read_excel(ec_file)
            pt = pd.read_excel(pt_file)
            
            labeled_cca = add_columns(cca, hp, ec, pt, pd.to_datetime(month_start_date))
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                labeled_cca.to_excel(writer, index=False, sheet_name="Labeled CCA")
            output.seek(0)
            
            st.download_button("Download Labeled Client’s Contract Audit", data=output, file_name="Labeled_CCA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("Please upload all required files before generating the output.")

if __name__ == "__main__":
    main()