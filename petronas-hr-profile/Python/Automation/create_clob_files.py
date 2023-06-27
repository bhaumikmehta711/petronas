from sqlalchemy import create_engine
from azure.storage.filedatalake import DataLakeServiceClient
import urllib
import glob
import spur_job_profile
import os
from datetime import datetime 
import config as conf
import spur_data_processor_sql
from Helper.sql_helper import *
from Helper.storage_account_helper import *
from Helper.key_vault_helper import *
from openpyxl import load_workbook
import logging.config
import shutil
import re



#Initialization
process_datetime = datetime.now().strftime('%Y%m%d%H%M')
consumption_dir =  conf.main_local_dir + '\\' + process_datetime
sql_engine = create_engine("mssql+pyodbc:///?autocommit=true&odbc_connect=%s" % urllib.parse.quote_plus(conf.sql_connection_string))
service_client = DataLakeServiceClient(account_url="{}://{}.dfs.core.windows.net".format("https", conf.storage_account_name), credential=conf.storage_account_key)

#Create dir
data_dir = consumption_dir + "\\" + "data"
log_dir = consumption_dir + "\\" + "log_files"
job_dir = consumption_dir + "\\" + "Job_SPUR"
job_blob_dir = job_dir + "\\" + "BlobFiles"
job_clob_dir = job_dir + "\\" + "ClobFiles"
job_dat_dir = job_dir + "\\" + "DatFiles"
position_dir = consumption_dir + "\\" + "Position_SPUR"
position_blob_dir = position_dir + "\\" + "BlobFiles"
position_clob_dir = position_dir + "\\" + "ClobFiles"
position_dat_dir = position_dir + "\\" + "DatFiles"

if not os.path.exists(consumption_dir):
    os.makedirs(consumption_dir)
    os.makedirs(data_dir)
    # os.makedirs(data_dir + "\\" + "Competency")
    os.makedirs(data_dir + "\\" + "final_processed_data")
    os.makedirs(data_dir + "\\" + "Position_Master_Data")
    os.makedirs(data_dir + "\\" + "Simplified_Template")
    os.makedirs(data_dir + "\\" + "Extracted_data")
    os.makedirs(data_dir + "\\" + "Write_up")
    os.makedirs(data_dir + "\\" + "PD")
    os.makedirs(data_dir + "\\" + "JCP")
    os.makedirs(log_dir)
    os.makedirs(job_dir)
    os.makedirs(job_blob_dir)
    os.makedirs(job_clob_dir)
    os.makedirs(job_dat_dir)
    os.makedirs(position_dir)
    os.makedirs(position_blob_dir)
    os.makedirs(position_clob_dir)
    os.makedirs(position_dat_dir)


#Load staging data
spur_data_processor_sql.data_processor(
    consumption_dir=consumption_dir,
    process_datetime=process_datetime,
    sql_engine = sql_engine
).spur_data()
spur_df = sql_read(sql_engine, f'SELECT \
        A.SPURCode UR_CODE,\
        A.SPURName UR_NAME,\
        A.PurposeAndAccountability,\
        A.Challenge CHALLENGES,\
        A.Experience EXPERIENCE \
    FROM [dbo].[SPUR] A')


for i in range(len(spur_df)):
        # DESCRIPTION
        with open(
            job_clob_dir + "\\" + spur_df.iloc[i]["UR_CODE"] + "_DESCRIPTION.txt",
            "w",
            encoding="utf-8",
        ) as f:
            f.write(
                 spur_df.iloc[i]["PurposeAndAccountability"]
               
            )
        # RESPONSIBILITY
        with open(
            job_clob_dir + "\\" + spur_df.iloc[i]["UR_CODE"] + "_RESPONSIBILITY.txt",
            "w",
            encoding="utf-8",
        ) as f:
            f.write(spur_df.iloc[i]["PurposeAndAccountability"].strip()) #["KPI"]
        # QUALIFICATION
        with open(
            job_clob_dir + "\\" + spur_df.iloc[i]["UR_CODE"] + "_QUALIFICATION.txt",
            "w",
            encoding="utf-8",
        ) as f:
            f.write(spur_df.iloc[i]["EXPERIENCE"].strip())



   
