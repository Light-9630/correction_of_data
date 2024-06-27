import streamlit as st
import pandas as pd
from io import BytesIO

# Load reference sheets at the backend
@st.cache_data
def load_reference_sheets():
    ref_sheets = ["trade", "state", "district", "type", "response"]
    refs = {sheet: pd.read_excel("ref.xlsx", sheet_name=sheet) for sheet in ref_sheets}
    return refs

refs = load_reference_sheets()

# Function to remove extra spaces
def clean_string(s):
    if pd.isna(s):
        return ''
    return ' '.join(str(s).strip().split())

# Helper function to correct values
def correct_value(value, correction_dict):
    if pd.isna(value) or str(value).strip() == '':
        return ''
    value_clean = clean_string(value)
    return correction_dict.get(value_clean.lower(), "#N/A")

# Main Streamlit app
def main():

    st.markdown("<h1 style='text-align: center;'>Data Correction App</h1>", unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload your main data Excel file", type=["xlsx"])

    if uploaded_file:
        main_data = pd.read_excel(uploaded_file)

        # Clean extra spaces from incorrect values in reference data and create dictionaries for correct values
        correct_trade = dict(zip(refs["trade"]["incorrect trade"].apply(clean_string), refs["trade"]["correct trade"]))
        correct_state = dict(zip(refs["state"]["incorrect state"].apply(clean_string), refs["state"]["correct state"]))
        correct_district = dict(zip(refs["district"]["incorrect district"].apply(clean_string), refs["district"]["correct district"]))
        correct_type = dict(zip(refs["type"]["incorrect type"].apply(clean_string), refs["type"]["correct type"]))
        correct_response = dict(zip(refs["response"]["incorrect response"].apply(clean_string), refs["response"]["correct response"]))
        correct_tr_cert = correct_response
        correct_ar_cert = correct_response

        # Columns to correct and their respective dictionaries
        columns_to_correct = {
            "trade": correct_trade,
            "state": correct_state,
            "district": correct_district,
            "type": correct_type,
            "response": correct_response,
            "tr certificate approved on sip": correct_tr_cert,
            "ar certificate approved on sip": correct_ar_cert
        }

        # Correct each column and insert it just after the original column
        for col, correction_dict in columns_to_correct.items():
            if col in main_data.columns:
                corrected_values = [correct_value(value, correction_dict) for value in main_data[col]]
                main_data.insert(main_data.columns.get_loc(col) + 1, f"correct {col}", corrected_values)

        # Show corrected data
        st.write("Corrected Data")
        st.dataframe(main_data)

        # Provide download link for corrected data
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        main_data.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.close()
        processed_data = output.getvalue()

        st.download_button(
            label="Download Corrected Data",
            data=processed_data,
            file_name="cleaned_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
