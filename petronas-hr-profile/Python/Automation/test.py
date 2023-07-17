from sqlalchemy import create_engine
from azure.storage.filedatalake import DataLakeServiceClient
import urllib
import glob
import spur_position_profile
import spur_job_profile
import os
from datetime import datetime 
import config as conf
import spur_data_processor_sql
import position_data_processor_sql
from Helper.sql_helper import *
from Helper.storage_account_helper import *
from Helper.key_vault_helper import *
from Helper.web_helper import *
import create_clob_file
import create_blob_file
import json

#Initialization
process_datetime = datetime.now().strftime('%Y%m%d%H%M')
consumption_dir =  conf.main_local_dir + '\\' + process_datetime
sql_engine = create_engine("mssql+pyodbc:///?autocommit=true&odbc_connect=%s" % urllib.parse.quote_plus(conf.sql_connection_string))
service_client = DataLakeServiceClient(account_url="{}://{}.dfs.core.windows.net".format("https", conf.storage_account_name), credential=conf.storage_account_key)
mulesoft_api_header = {'client_id': conf.mulesoft_api_client_id, 'client_secret': conf.mulesoft_api_client_secret, 'Content-Type': 'application/json'}
batch_list = []

# storage_account_name = 'adlspetronashrprofiledev'
# consumption_container_name = 'consumption'
# remote_path= datetime.now().strftime('%Y/%m/%d') 
# print(remote_path)

# Batch_Consumption_Path = 'https://'+storage_account_name+'.blob.core.windows.net/'+consumption_container_name+'/'+remote_path
# print(Batch_Consumption_Path)

# #         "url": "https://adlspetronashrprofiledev.blob.core.windows.net/consumption/2023/07/07/"


position_df = pd.DataFrame(sql_read(sql_engine, f'SELECT distinct BatchConsumptionPath as "url" FROM [Batch] where MulesoftNotifiedFlag=0 and isnull(BatchConsumptionPath,\'\')!=\'\''))
list_of_dicts = position_df.to_dict('records')
print(list_of_dicts)
# for item in list_of_dicts:
#     print(item)
#for item in position_df:
        #print(item)
post(
    url = 'https://devapi.petronas.com/dev/pet/web/hrempmgmt/exp/api/v1/employees' ,
    header = {'Authorization' : 'Basic UEVUX01VTEVfSU5URUdSQVRJT046V2VsYzBtZUAxMjM=','Content-Type' : 'application/json' },

    payload=json.dumps(list_of_dicts )
)

