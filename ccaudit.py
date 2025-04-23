import streamlit as st
import pandas as pd
import io
import xlsxwriter
from datetime import datetime

def preprocess_cca(df):
    df['Start Of Contract'] = pd.to_datetime(df['Start Of Contract'], errors='coerce')
    df['Contract'] = df['Contract'].astype(str).str.split('.').str[0]
    return df

def preprocess_hp(df):
    df['Contract Name'] = df['Contract Name'].astype(str).str.replace("Contr-", "", regex=False).str.strip()
    return df

def preprocess_pt(df):
    df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce')
    df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce')
    return df

def preprocess_ec(df):
    df['Cont #'] = df['Cont #'].astype(str).str.split('.').str[0]
    return df

def add_columns(cca, hp, ec, pt, month_start_date):
    if cca.empty or hp.empty or ec.empty or pt.empty:
        return pd.DataFrame()

    def map_nationality(nat):
        if pd.isna(nat):
            return 'Other'
        nat_str = str(nat).strip()
        return nat_str if nat_str in ['Filipina', 'Ethiopian'] else 'Other'

    cca['Mapped Nationality'] = cca['Maid Nationality During Payroll Month'].apply(map_nationality)

    hp_filtered = hp[(hp['Status'] == 'WITH_CLIENT') & (hp['Type Of maid'] == 'CC')].copy()
    hp_filtered['Contract Name'] = hp_filtered['Contract Name'].astype(str).str.strip()
    hp_contract_list = hp_filtered['Contract Name'].tolist()

    cca['Contract'] = cca['Contract'].astype(str).str.strip()
    cca['To Check'] = cca['Contract'].apply(lambda x: 'Yes' if x in hp_contract_list else 'No')

    ec_list = ec['Cont #'].tolist()

    def determine_exceptional_case(row):
        if row['To Check'] == 'No':
            return ''
        return 'Yes' if row['Contract'] in ec_list else 'No'

    cca['Exceptional Case'] = cca.apply(determine_exceptional_case, axis=1)

    pt_latest = pt.loc[pt.groupby(['Nationality', 'Contract Type'])['End Date'].idxmax()]
    pt_latest = pt_latest[['Nationality', 'Contract Type', 'Minimum monthly payment + VAT']]

    def get_latest_price(mapped_nat, contract_type):
        match = pt_latest[(pt_latest['Nationality'] == mapped_nat) & (pt_latest['Contract Type'] == contract_type)]
        if not match.empty:
            return pd.to_numeric(match['Minimum monthly payment + VAT'].values[0], errors='coerce')
        return None

    def get_price_at_contract_start(mapped_nat, contract_type, contract_date):
        if pd.isna(contract_date):
            return None
        match = pt[
            (pt['Nationality'] == mapped_nat) &
            (pt['Contract Type'] == contract_type) &
            (pt['Start Date'].dt.date <= contract_date.date()) &
            (pt['End Date'].dt.date >= contract_date.date())
        ]
        if not match.empty:
            return pd.to_numeric(match.iloc[0]['Minimum monthly payment + VAT'], errors='coerce')
        return None

    def check_paying_now(row):
        if row['To Check'] == 'No':
            return ''
        if row['Exceptional Case'] == 'Yes':
            ec_value = ec.loc[ec['Cont #'] == row['Contract'], 'Monthly Payment'].values
            if ec_value.size > 0:
                val = str(ec_value[0]).strip()
                if val in ['N/A', '-']:
                    return 'Yes'
                try:
                    val_num = float(val)
                    return 'Yes' if row['Amount Of Payment'] >= val_num else 'No'
                except:
                    return 'No'
            return 'No'
        else:
            latest_price = get_latest_price(row['Mapped Nationality'], row['Contract Type'])
            if latest_price is not None:
                return 'Yes' if row['Amount Of Payment'] >= latest_price else 'No'
            return 'No'

    cca['Paying Correctly on Price of Now'] = cca.apply(check_paying_now, axis=1)

    def check_paying_contract_start(row):
        if row['To Check'] == 'No' or row['Exceptional Case'] == 'Yes':
            return ''
        contract_price = get_price_at_contract_start(row['Mapped Nationality'], row['Contract Type'], row['Start Of Contract'])
        if contract_price is not None:
            return 'Yes' if row['Amount Of Payment'] >= contract_price else 'No'
        return 'No'

    cca['Paying Correctly on Price of Contract Start Date'] = cca.apply(check_paying_contract_start, axis=1)

    def check_upgrading(row):
        if row['To Check'] == 'No' or row['Exceptional Case'] == 'Yes':
            return ''
        if pd.isna(row['Upgrading Nationality Payment Amount']) or row['Upgrading Nationality Payment Amount'] == '':
            return 'No'
        latest_price = get_latest_price(row['Mapped Nationality'], row['Contract Type'])
        if latest_price is not None:
            total_paid = row['Amount Of Payment'] + row['Upgrading Nationality Payment Amount']
            return 'Yes' if total_paid >= latest_price else 'No'
        return 'No'

    cca['Paying Correctly if Upgrading Nationality'] = cca.apply(check_upgrading, axis=1)

    def check_pro_rated(row):
        if row['To Check'] == 'No' or row['Exceptional Case'] == 'Yes':
            return ''
        if pd.isna(row['Start Of Contract']):
            return 'No'
        if row['Start Of Contract'] < month_start_date:
            return 'No'
        return 'Yes' if row['Amount Of Payment'] >= row['Pro-Rated'] else 'No'

    cca['Paying Correctly if Pro-Rated Value'] = cca.apply(check_pro_rated, axis=1)

    def check_old_price(row):
        if row['To Check'] == 'No' or row['Exceptional Case'] == 'Yes':
            return ''
        filtered_pt = pt[
            (pt['Nationality'] == row['Mapped Nationality']) &
            (pt['Contract Type'] == row['Contract Type'])
        ]
        for price in filtered_pt['Minimum monthly payment + VAT']:
            try:
                price = float(price)
                if abs(row['Amount Of Payment'] - price) <= 5:
                    return 'Yes'
            except:
                continue
        return 'No'

    cca['Paying Correctly on Old Price'] = cca.apply(check_old_price, axis=1)

    cca_export = cca.copy()
    cca_export['Start Of Contract'] = cca_export['Start Of Contract'].dt.strftime('%m/%d/%Y')

    return cca_export

def main():
    st.title("Client’s Contract Audit Processing")

    month_start_date_input = st.date_input("Month Start Date", value=datetime.today())
    month_start_date = datetime.combine(month_start_date_input, datetime.min.time())

    hp_file = st.file_uploader("Upload Housemaid Payroll", type=["xls", "xlsx"], key="hp")
    cca_file = st.file_uploader("Upload Client’s Contract Audit", type=["xls", "xlsx"], key="cca")
    ec_file = st.file_uploader("Upload Exceptional Cases", type=["xls", "xlsx"], key="ec")
    pt_file = st.file_uploader("Upload Price Trends", type=["xls", "xlsx"], key="pt")

    if st.button("Generate"):
        if hp_file and cca_file and ec_file and pt_file:
            hp = preprocess_hp(pd.read_excel(hp_file, engine='openpyxl'))
            cca = preprocess_cca(pd.read_excel(cca_file, engine='openpyxl'))
            ec = preprocess_ec(pd.read_excel(ec_file, engine='openpyxl'))
            pt = preprocess_pt(pd.read_excel(pt_file, engine='openpyxl'))

            labeled_cca = add_columns(cca, hp, ec, pt, month_start_date)

            if labeled_cca.empty:
                st.error("Error: Processed DataFrame is empty.")
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
