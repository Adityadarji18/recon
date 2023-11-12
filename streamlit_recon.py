import streamlit as st
import pandas as pd
import base64
from github import Github

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

    # Create a Pandas Excel writer using xlsxwriter as the engine
    with pd.ExcelWriter('gstr_recon_output.xlsx', engine='xlsxwriter') as writer:

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

# Streamlit app layout
st.title('GST Reconciliation App')
st.write('Upload your Excel file')

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
if uploaded_file is not None:
    process_excel(uploaded_file)
    st.write('Download the processed file:')
    file_path = 'gstr_recon_output.xlsx'  # Local path to the processed file

    # Save the file to GitHub
    access_token = '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'  # Your GitHub access token
    repo_owner = 'Adityadarji18'  # Repository owner username
    repo_name = 'recon'  # Repository name

    # Initialize PyGithub
    g = Github(access_token)

    # Get the repository
    repo = g.get_user(repo_owner).get_repo(repo_name)

    # Read the file content
    file_content = open(file_path, 'rb').read()

    # Create or update the file in the repository
    repo.create_file('path/in/repo/gstr_recon_output.xlsx', 'Commit message', file_content, branch='main')

    # Display download link in Streamlit
    with open(file_path, 'rb') as f:
        file_bytes = f.read()
        file_base64 = base64.b64encode(file_bytes).decode()
        st.markdown(f"[Download file](data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{file_base64})", unsafe_allow_html=True)
