### API Call
# https://www.ecom.immi.gov.au/evo/ws/first-party/json?country=IDN&dateofbirth=20010108&passport=E3645392&visagrant=0579584823213

### install syntax
# !pip install --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org ipywidgets

import json
import pandas as pd
import requests
from datetime import datetime
import re, os
import numpy as np

bulk_visa_path = "C:\\Users\\abueljerick\\OneDrive - GWF\\Documents\\Reports\\10 Temporary Visa about to Expire\\Visa Checker\\03 Input Files\\Visa Bulk Import Template.xlsx"
bulk_visa_df = pd.read_excel(bulk_visa_path, dtype=str, sheet_name = "Input Visa Here" )
bulk_visa_df['Birthday'] = pd.to_datetime(bulk_visa_df['Birthday']).dt.strftime("%Y/%m/%d")
bulk_visa_dict = bulk_visa_df.to_dict('index')

list_for_concat = []
visa_cols = ['visaGrantNumber','visaSubclass','visaDescription','visaGrantDate','visaExpiryDate','visaStatus','workEntitlements','visaConditions']

def remove_html_tags(text):
    try:
        clean = re.compile('<.*?>')
        return re.sub(clean, '', text)
    except TypeError:
        return np.nan

json_url = "https://www.ecom.immi.gov.au/evo/ws/first-party/json"

class VEVO:
    def __init__(self, dob, passport, country_code, visa_number):
        self.n = visa_number
        self.p = passport
        self.c = country_code
        self.dob = dob

    def visa(self):
        visa = {}
        response = requests.get(self.prepare_url(),verify=False)
        if response.status_code != 200:
            raise Exception(f"HTTP call failed: {response.status_code}")
        return response.json()

    def prepare_url(self):
        u = f"{json_url}?"
        params = {
            "passport": self.p.upper(),
            "country": self.c.upper(),
            "dateofbirth": self.dob.strftime("%Y%m%d"),
        }
        vn = self.n.upper()
        if vn.startswith("E"):
            params["trn"] = vn
        else:
            params["visagrant"] = vn
        u += "&".join([f"{key}={value}" for key, value in params.items()])
        return u

if __name__ == "__main__":
    for visa_details in bulk_visa_dict.values(): 
        # date of birth
        dob = datetime.strptime(visa_details['Birthday'], "%Y/%m/%d" )

        # passport number
        passport = visa_details['Passport Number']
        # country code
        cc = visa_details['Country']
        # visa grant number or transaction reference number
        vgn = visa_details['Visa Grant Number']
        v = VEVO(dob, passport, cc, vgn)
        try:
            visa = v.visa()
            ### Separate dictionaries
            entitlementDetails = visa['entitlementDetails']
            visa.pop('entitlementDetails',None)
            ### merge back
            visa.update(entitlementDetails)
            visa.pop('enquiryDetails',None)
            visa.pop('visaConditionCodes',None)

            ## write to dataframe
            df = pd.DataFrame(visa,index=[0])
            df = df[visa_cols]
            list_for_concat.append(df)
        except Exception as err:
            print(err)

        ### merge result in the input df
        visa_result_df = pd.concat(list_for_concat)
        final_df = bulk_visa_df.merge(visa_result_df, how = 'left', left_on = 'Visa Grant Number', right_on = 'visaGrantNumber' )
        final_df['workEntitlements'] = final_df['workEntitlements'].apply(remove_html_tags)
        final_df['visaConditions'] = final_df['visaConditions'].apply(remove_html_tags)

date_str = datetime.today().strftime("%d%m%Y")
output_file_name = 'C:\\Users\\abueljerick\\OneDrive - GWF\\Documents\\Reports\\10 Temporary Visa about to Expire\Visa Checker\\04 Output Files\\Visa Vevo Bulk Check '+date_str+'.xlsx'

# creating an ExcelWriter object
with pd.ExcelWriter(output_file_name) as writer:
    final_df.to_excel(writer, sheet_name=date_str, index=False)