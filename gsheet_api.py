from __future__ import print_function
from google.oauth2 import service_account
import googleapiclient.discovery
from path import *
import config 
from googleapiclient.errors import HttpError


class GsheetData:
    def __init__(self, spreadsheet_id, key_range, charles_data, input_range):
        self.spreadsheet_id = spreadsheet_id
        self.key_range = key_range
        self.input_range = input_range
        self.charles_data = charles_data

   
    def gsheet_auth(self):

        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    

        SERVICE_ACCOUNT_FILE = path_param_2() / "charles_json/service_account.json"

        credentials = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=SCOPES)

        gsheetadmin = googleapiclient.discovery.build('sheets', 'v4', credentials=credentials)

        gsheet = gsheetadmin.spreadsheets()

        return gsheet

    def get_gsheet_data(self):
        gsheet_ = self.gsheet_auth()

        dimension_key = gsheet_.values().get(spreadsheetId=self.spreadsheet_id,
                                        range=self.key_range).execute()
        
        return dimension_key
        
       



    def dimension_key_to_dict(self):
        dimension_key = self.get_gsheet_data()
        dimension_key_list = dimension_key.get('values')[0]
        dimension_key_dict = {}
        for k in dimension_key_list:
            k.strip()
            dimension_key_dict[k] = None
        return dimension_key_dict

       

    def match_dimension_val(self):
        
        dimension_key_dict = self.dimension_key_to_dict()
        


        #put the value from charles_data to dimension_key_dict
        
        
        for k, v in config.charles_data.items():
            for k2, v2 in dimension_key_dict.items():
                if k == k2:
                   dimension_key_dict[k2] = v

        #print(dimension_key_dict)
        return dimension_key_dict
        
    
    
    def input_charles_value(self):

        data_dict = self.match_dimension_val()
        #print(type(data_dict))

        value_list = []
        for k, v in data_dict.items():
            value_list.append(v)
       
        try:
            gsheet_ = self.gsheet_auth()

            values = [value_list]
            body = {
                'values': values
            }
            result = gsheet_.values().update(
                spreadsheetId=self.spreadsheet_id, range=self.input_range,
                valueInputOption="RAW", body=body).execute()
            

            print(f"{result.get('updatedCells')} cells updated.")
            return result
        
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error




input_charles_data = GsheetData(spreadsheet_id=config.config_dict.get("spreadsheet_id"),
                                 key_range=config.config_dict.get("key_range"), 
                                input_range=config.config_dict.get("input_range"),
                                charles_data=config.charles_data)


GsheetData.input_charles_value(input_charles_data)