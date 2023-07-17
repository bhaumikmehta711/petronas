
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

#Initialization
process_datetime = datetime.now().strftime('%Y%m%d%H%M')
consumption_dir =  conf.main_local_dir + '\\' + process_datetime
sql_engine = create_engine("mssql+pyodbc:///?autocommit=true&odbc_connect=%s" % urllib.parse.quote_plus(conf.sql_connection_string))
service_client = DataLakeServiceClient(account_url="{}://{}.dfs.core.windows.net".format("https", conf.storage_account_name), credential=conf.storage_account_key)
mulesoft_api_header = {'client_id': conf.mulesoft_api_client_id, 'client_secret': conf.mulesoft_api_client_secret, 'Content-Type': 'application/json'}
batch_list = []

remote_path_date= datetime.now().strftime('%Y/%m/%d') 
Batch_Consumption_Path = 'https://'+conf.storage_account_name+'.blob.core.windows.net/'+conf.consumption_container_name+'/'+remote_path_date


#Create dir
data_dir = consumption_dir + "\\" + "data"
log_dir = consumption_dir + "\\" + "log_files"
job_dir = consumption_dir + "\\" + "Job_SPUR"
job_clob_dir = job_dir + "\\" + "ClobFiles"
job_dat_dir = job_dir + "\\" + "DatFiles"
position_dir = consumption_dir + "\\" + "Position_SPUR"
position_blob_dir = position_dir + "\\" + "BlobFiles"
position_clob_dir = position_dir + "\\" + "ClobFiles"
position_dat_dir = position_dir + "\\" + "DatFiles"

spur_details_file_path = data_dir + "\\" + "final_processed_data\\{}_details.xlsx".format(process_datetime)
spur_position_file_path = data_dir + "\\final_processed_data\\{}_position_profile_data.xlsx".format(process_datetime)

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
    os.makedirs(job_clob_dir)
    os.makedirs(job_dat_dir)
    os.makedirs(position_dir)
    os.makedirs(position_blob_dir)
    os.makedirs(position_clob_dir)
    os.makedirs(position_dat_dir)


#Download required documents
download_from_adls(
    service_client = service_client, 
    container = conf.master_container_name, 
    remote_path = 'PET_Job SPUR.xlsx',
    local_dir = job_dir,
    new_file_name = f'PET_Job_SPUR_{process_datetime}.xlsx')
download_from_adls(
    service_client = service_client, 
    container = conf.master_container_name, 
    remote_path = 'PET_Position SPUR.xlsx',
    local_dir = position_dir,
    new_file_name = f'PET_Position_SPUR_{process_datetime}.xlsx')


#region SPUR
#Load approved SPUR
sql_execute(sql_engine, f'EXEC uspGetSPURBatchForConsumption @pStagingTableName = \'SPUR_{process_datetime}\'')

spur_df = sql_read(sql_engine, f'SELECT * FROM [Staging].[SPUR_{process_datetime}]')

if not spur_df.empty:
    #Load SPUR staging data
    spur_data_processor_sql.data_processor(
        consumption_dir=consumption_dir,
        process_datetime=process_datetime,
        sql_engine = sql_engine
    ).spur_data()
    job_template_file_path = os.path.abspath(glob.glob(job_dir + "\\" + "*.xlsx")[0])

    #Create dat files
    spur_job_profile.spur_job_profile(
            job_template_file_path=job_template_file_path,
            spur_df=spur_df,
            spur_details_file_path=spur_details_file_path,
            job_dat_dir=job_dat_dir,
            log_dir=log_dir,
        )
    os.rename(job_template_file_path, job_template_file_path.replace('_' + process_datetime, ''))
    
    #Create blob files
    create_clob_file.create_clob_file(df = spur_df, clob_folder_path = job_clob_dir)

    #Upload dat and blob files
    uploaded_files = upload_to_adls(
            service_client = service_client,
            container = conf.consumption_container_name,
            local_path = job_dir,
            remote_path= datetime.now().strftime('%Y/%m/%d') + '/Job_SPUR'
        )
    
    #Copy clob files
    create_blob_file.create_blob_file(service_client = service_client, df = spur_df, destination_remote_path = datetime.now().strftime('%Y/%m/%d') + '/Job_SPUR/BlobFiles')

sql_execute(sql_engine, f'UPDATE A\
    SET \
        A.BatchCompleteStatus = \'Completed\',\
        A.BackendUserModifiedBy = \'{conf.sql_user}\',\
        A.BackendUserModifiedTimeStamp = GETUTCDATE(),\
        A.BatchConsumptionPath = \'{Batch_Consumption_Path}\'\
    FROM Batch A\
    INNER JOIN Staging.[SPUR_{process_datetime}] B ON A.BatchID = B.BatchID\
    DROP TABLE Staging.[SPUR_{process_datetime}]')
#endregion


#region Position
#Load approved position
sql_execute(sql_engine, f'EXEC uspGetPositionBatchForConsumption @pStagingTableName = \'Position_{process_datetime}\'')

position_df = sql_read(sql_engine, f'SELECT * FROM [Staging].[Position_{process_datetime}]')

if not position_df.empty:
    #Load SPUR staging data
    position_data_processor_sql.data_processor(
        consumption_dir=consumption_dir,
        process_datetime=process_datetime,
        sql_engine = sql_engine
    ).position_data()
    position_template_file_path = os.path.abspath(glob.glob(position_dir + "\\" + "*.xlsx")[0])

    #Create dat files
    spur_position_profile.spur_position_profile(
        position_blob_dir=position_blob_dir,
        position_template_file_path=position_template_file_path,
        position_profile_df = position_df,
        position_details_file_path = spur_position_file_path,
        spur_data_file_path='',
        spur_details_file_path='',
        spur_position_file_path=spur_position_file_path,
        LC_file_path='',
        TC_file_path='',
        jcp_file_path='',
        position_dat_dir=position_dat_dir,
        SPUR_ID_ignore_list='',
        content_item_file_path='',
        read_jcp='',
        WS_path='',
        log_dir=log_dir,
    )
    os.rename(position_template_file_path, position_template_file_path.replace('_' + process_datetime, ''))
    
    #Create blob files
    create_clob_file.create_clob_file(df = position_df, clob_folder_path = position_clob_dir)

    #Upload dat and blob files
    uploaded_files = upload_to_adls(
            service_client = service_client,
            container = conf.consumption_container_name,
            local_path = position_dir,
            remote_path= datetime.now().strftime('%Y/%m/%d') + '/Job_Position'
        )

    #Copy clob files
    create_blob_file.create_blob_file(service_client = service_client, df = position_df, destination_remote_path = datetime.now().strftime('%Y/%m/%d') + '/Job_Position/BlobFiles')



sql_execute(sql_engine, f'UPDATE A\
    SET \
        A.BatchCompleteStatus = \'Completed\',\
        A.BackendUserModifiedBy = \'{conf.sql_user}\',\
        A.BackendUserModifiedTimeStamp = GETUTCDATE(),\
        A.BatchConsumptionPath = \'{Batch_Consumption_Path}\'\
    FROM Batch A\
    INNER JOIN (SELECT BatchID FROM Staging.[SPUR_{process_datetime}] UNION SELECT BatchID FROM Staging.[Position_{process_datetime}])  B ON A.BatchID = B.BatchID\
    DROP TABLE Staging.[Position_{process_datetime}]')


# post(
#     url = conf.mulesoft_api_url,
#     header = mulesoft_api_header,
#     payload=[
#     {
#         "url": "https://adlspetronashrprofiledev.blob.core.windows.net/consumption/2023/07/07/"
#     }
# ]
# )
#endregion