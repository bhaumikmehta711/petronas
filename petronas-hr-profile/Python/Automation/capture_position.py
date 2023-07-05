import pandas as pd
from sqlalchemy import create_engine
import urllib
import logging
import config as conf
from azure.storage.filedatalake import DataLakeServiceClient
from Helper.sql_helper import *
from Helper.storage_account_helper import *
import shutil

#Initialization
sql_engine = create_engine("mssql+pyodbc:///?autocommit=true&odbc_connect=%s" % urllib.parse.quote_plus(conf.sql_connection_string))
service_client = DataLakeServiceClient(account_url="{}://{}.dfs.core.windows.net".format("https", conf.storage_account_name), credential=conf.storage_account_key)

#Capture non-processed batch
pd_batch = sql_read(engine = sql_engine, query = 'EXEC [uspGetBatchForProcessing] @pBatchType = \'Position\'')

#Process each batch
for index, position in pd_batch.iterrows():
    try:
        batch_id = position['BatchID']
        batch_name = position['BatchName']
        pd_position = pd.DataFrame(columns = ['maintenance_mode', 'effective_start_date', 'effective_end_date', 'position_profile_code', 'SPUR_ID', 'position_name', 'job_grade', 'role_level', 'company_code'])
        
        local_batch_dir = conf.main_local_dir + '/' + batch_name

        #Download files for a batch
        downloaded_files = download_from_adls(service_client = service_client, 
            container = conf.intake_container_name, 
            remote_path = conf.intake_job_position_dir + '/' + batch_name,
            local_dir = local_batch_dir + '/data/Write_up')
        
        position_writeups = downloaded_files

        for position_writeup in position_writeups:
            pd_writeup = pd.read_excel(position_writeup, sheet_name='Position Profile Maintenance', skiprows=[1])
            for index, position in pd_writeup.iterrows():                
                current_position = {}
                current_position['maintenance_mode'] = position['Maintenance Mode']
                current_position['effective_start_date'] = position['Effective Start Date']
                current_position['effective_end_date'] = position['Effective End Date']
                current_position['position_profile_code'] = position['Position Profile Code']
                current_position['SPUR_ID'] = position['SPUR ID']
                current_position['position_name'] = position['Position Name'] 
                current_position['job_grade'] = position['Job Grade']
                current_position['role_level'] = position['Role Level']
                current_position['company_code'] = position['Company Code']
                pd_position = pd_position.append(current_position, ignore_index = True)


        # Merge position in database
        if not pd_position.empty:
            sql_insert(engine = sql_engine, df = pd_position, table_name = batch_name, schema_name = 'Staging')
            sql_query = f'\
                UPDATE [dbo].[Batch] \
                SET \
                    BatchStatus = \'Pending Submit\',\
                    BatchProcessStatus = \'Completed\',\
                    BackendUserModifiedBy = \'{conf.sql_user}\',\
                    BackendUserModifiedTimestamp = GETUTCDATE()\
                WHERE BatchName = \'{batch_name}\'\
                \
                MERGE [dbo].[Position] A\
                USING (SELECT \
                        F.PositionID position_id,\
                        A.maintenance_mode, \
                        CONVERT(DATETIME, effective_start_date, 103) effective_start_date, \
                        CONVERT(DATETIME, effective_end_date, 103) effective_end_date, \
                        A.position_profile_code, \
                        D.SPURCode SPUR_code, \
                        D.SPURID SPUR_ID, \
                        A.position_name, \
                        C.JobGradeID job_grade_id, \
                        E.RoleLevelID role_level_id, \
                        A.company_code, \
                        B.BatchID, \
                        B.SubmittedBy, \
                        B.SubmittedTimeStamp \
                    FROM [Staging].[{batch_name}] A \
                    INNER JOIN [dbo].[Batch] B ON \'{batch_name}\' = B.BatchName\
                    INNER JOIN Master.JobGrade C ON C.JobGradeName = A.job_grade\
                    INNER JOIN DBO.SPUR D ON D.SPURCode = A.SPUR_ID\
                    INNER JOIN Master.RoleLevel E ON E.RoleLevelName = A.role_level AND E.RoleLevelID = D.RoleLevelID\
                    LEFT JOIN dbo.Position F ON F.PositionCode = A.position_profile_code) B\
                ON A.PositionID = B.position_id\
                WHEN MATCHED THEN\
                UPDATE SET BatchID = B.BatchID, BackendUserModifiedBy = \'{conf.sql_user}\', BackendUserModifiedTimestamp = GETUTCDATE()\
                WHEN NOT MATCHED THEN\
                INSERT (MaintenanceMode, PositionCode, PositionName, EffectiveStartDate, EffectiveEndDate, SPURID, SPURCode, JobGradeID, RoleLevelID, CompanyCode, SubmittedTimeStamp, EndUserCreatedBy, EndUserCreatedTimestamp)\
                VALUES (B.maintenance_mode, B.position_profile_code, B.position_name, B.effective_start_date, B.effective_end_date, B.SPUR_ID, B.SPUR_code, B.job_grade_id, B.role_level_id, B.company_code, B.SubmittedTimeStamp, B.SubmittedBy, B.SubmittedTimeStamp);\
                \
                EXEC [dbo].[uspLoadPosition] @pBatchID = {batch_id}\
                \
                EXEC [audit].[uspAddEmail] @pBatchID = {batch_id}\
                \
                DROP TABLE [Staging].[{batch_name}]'
            sql_execute(engine=sql_engine, query = sql_query)

            # shutil.rmtree(local_batch_dir)
    except Exception as e:
        logging.error(f'Encountered error in create_spur.py: {str(e)}')
        raise ValueError(e)