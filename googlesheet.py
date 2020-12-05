import os

import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from df2gspread import df2gspread as d2g
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/spreadsheets']

CLIENT_SECRET_FILE: str = 'Employee_details.json'



# ----------------------------------------------------------------------------------------------------------------------#

#    Establish Connection with Google API's via client_secret.json file resides in src directory
# ----------------------------------------------------------------------------------------------------------------------#
def get_google_api_connection():
    credential_path = os.path.join(CURRENT_DIR, CLIENT_SECRET_FILE)
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credential_path, SCOPES)
    google_connection = gspread.authorize(credentials)
    SHEET_ID = '1qTRyOfzg3gPht1FA4nvJCAMKMsyLnV90cAil0rzMFRE'
    SHEET_NAME = 'Employee_details'
    googlesheetdata = google_connection.open_by_key(SHEET_ID).worksheet(SHEET_NAME).get_all_values()
    sheets_df = pd.DataFrame(googlesheetdata)
    sheetdata = sheets_df.drop(sheets_df.index[0])
    lmt = sheetdata[sheetdata.columns[0:14]]
    val = lmt.sort_values([4], ascending=False)
    newval = val.groupby([9])
    for key, item in newval:
        group_df = newval.get_group(key)
        result = group_df.head(10)
        spreadsheet_key = '1qTRyOfzg3gPht1FA4nvJCAMKMsyLnV90cAil0rzMFRE'
        wks_name = 'Employees_top10_salary_deptwise'
        d2g.upload(result, spreadsheet_key, wks_name, credentials=credentials, row_names=True)


get_google_api_connection()
