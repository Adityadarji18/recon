import pandas as pd
import streamlit as st
import base64
from io import BytesIO
import os

def process_excel(input_file):
    # Reading input Excel file
    our = pd.read_excel(input_file, sheet_name='Our')
    gov = pd.read_excel(input_file, sheet_name='Gov')

    # 'GSTIN No' in our but not in gov
    AccessOurData = our[~our['GSTIN No'].isin(gov['GSTIN of supplier'])]

    # 'GSTIN of supplier' in gov but not in our
    AccessGovData = gov[~gov['GSTIN of supplier'].isin(our['GSTIN No'])]

    # Grouping and summing the 'TOTAL' and 'Invoice Value' based on 'GSTIN'
    our_grouped = our.groupby('GSTIN No')['TOTAL'].sum().reset_index()
    gov_grouped = gov.groupby('GSTIN of supplier')['Invoice Value'].sum().reset_index()

    # Merging the grouped data based on 'GSTIN' to compare 'TOTAL' and 'Invoice Value'
    merged_data = pd.merge(our_grouped, gov_grouped, left_on='GSTIN No', right_on='GSTIN of supplier', suffixes=('_our', '_gov'))

    # Matched and Mismatched dataframes
    Matched = merged_data[merged_data['TOTAL'] == merged_data['Invoice Value']]
    Mismatched = merged_data[merged_data['TOTAL'] != merged_data['Invoice Value']]

    # Filter 'our' and 'gov' data based on 'GSTIN' from merged data
    Matched_our = our[our['GSTIN No'].isin(Matched['GSTIN No'])]
    Matched_gov = gov[gov['GSTIN of supplier'].isin(Matched['GSTIN No'])]

    Mismatched_our = our[our['GSTIN No'].isin(Mismatched['GSTIN No'])]
    Mismatched_gov = gov[gov['GSTIN of supplier'].isin(Mismatched['GSTIN No'])]

        # Create a BytesIO buffer to store the processed Excel file in memory
    processed_file_buffer = BytesIO()

    # Create an Excel writer with the buffer
    with pd.ExcelWriter(processed_file_buffer, engine='xlsxwriter') as writer:
        # Write 'Our' dataframe
        our.to_excel(writer, sheet_name='Our', startrow=0, startcol=0, index=False)

        # Write 'Gov' dataframe
        gov.to_excel(writer, sheet_name='Gov', startrow=0, startcol=0 ,index=False)

        # Write 'AccessOurData' dataframe
        AccessOurData.to_excel(writer, sheet_name='AccessOurData', startrow=0, startcol=0 ,index=False)

        # Write 'AccessGovData' dataframe
        AccessGovData.to_excel(writer, sheet_name='AccessGovData', startrow=0, startcol=0, index=False)

        # Write 'Matched' dataframes
        Matched_our.to_excel(writer, sheet_name='Matched', startrow=0, startcol=0,index=False)
        Matched_gov.to_excel(writer, sheet_name='Matched', startrow=0, startcol=Matched_our.shape[1] + 2, index=False)

        # Write 'Mismatched' dataframes
        Mismatched_our.to_excel(writer, sheet_name='Mismatched', startrow=0, startcol=0,index=False)
        Mismatched_gov.to_excel(writer, sheet_name='Mismatched', startrow=0, startcol=Mismatched_our.shape[1] + 2, index=False)

    # Provide download link without saving to local folder
    st.write('Download the processed file:')
    file_bytes = processed_file_buffer.getvalue()
    file_base64 = base64.b64encode(file_bytes).decode()
    st.markdown(f"[Download file](data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{file_base64})", unsafe_allow_html=True)

# Streamlit app layout
st.title('GST Reconciliation App')
st.write('Upload your Excel file')

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Create a BytesIO buffer to store the processed Excel file in memory
    processed_file_buffer = BytesIO()
    process_excel(uploaded_file)
