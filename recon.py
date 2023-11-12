import pandas as pd

# Reading our data 
our = pd.read_excel('C:/Users/Aditya/Dropbox/PC/Downloads/gstr_recon.xlsx',sheet_name='Our')

gov = pd.read_excel('C:/Users/Aditya/Dropbox/PC/Downloads/gstr_recon.xlsx',sheet_name='Gov')

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
with pd.ExcelWriter('gstr_recon.xlsx', engine='xlsxwriter') as writer:

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

'''
from flask import Flask, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

@app.route('/download_excel')
def download_excel():
    output = BytesIO()
    with pd.ExcelWriter('gstr_recon.xlsx', engine='xlsxwriter') as writer:
        our.to_excel(writer, sheet_name='Our', startrow=0, startcol=0, index=False)
        gov.to_excel(writer, sheet_name='Gov', startrow=0, startcol=0, index=False)
        AccessOurData.to_excel(writer, sheet_name='AccessOurData', startrow=0, startcol=0, index=False)
        AccessGovData.to_excel(writer, sheet_name='AccessGovData', startrow=0, startcol=0, index=False)
        Matched_our.to_excel(writer, sheet_name='Matched', startrow=0, startcol=0, index=False)
        Matched_gov.to_excel(writer, sheet_name='Matched', startrow=0, startcol=Matched_our.shape[1] + 2, index=False)
        Mismatched_our.to_excel(writer, sheet_name='Mismatched', startrow=0, startcol=0, index=False)
        Mismatched_gov.to_excel(writer, sheet_name='Mismatched', startrow=0, startcol=Mismatched_our.shape[1] + 2, index=False)

    output.seek(0)
    return send_file(output, as_attachment=True, attachment_filename='gstr_recon.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run()
    
'''
