import pandas as pd
from sqlalchemy import create_engine
import urllib
import logging
import config as conf
from azure.storage.filedatalake import DataLakeServiceClient
from Helper.sql_helper import *
from Helper.storage_account_helper import *
import spur_pptx_to_xlsx
import shutil

#Initialization
sql_engine = create_engine("mssql+pyodbc:///?autocommit=true&odbc_connect=%s" % urllib.parse.quote_plus(conf.sql_connection_string))
service_client = DataLakeServiceClient(account_url="{}://{}.dfs.core.windows.net".format("https", conf.storage_account_name), credential=conf.storage_account_key)

#Capture non-processed batch
pd_batch = sql_read(engine = sql_engine, query = 'EXEC [uspGetBatchForProcessing]')

#Process each batch
for index, row in pd_batch.iterrows():
    try:
        batch_id = row['BatchID']
        batch_name = row['BatchName']
        pd_spur = pd.DataFrame(columns = ['SPUR_code', 'SPUR_name', 'SKG_name', 'SPUR_file_storage_account_path', 'purpose_and_accountability', 'challenge', 'experience'])
        local_batch_dir = conf.main_local_dir + '/' + batch_name

        #Download files for a batch
        downloaded_files = download_from_adls(
            service_client = service_client, 
            container = conf.intake_container_name, 
            remote_path = conf.intake_job_spur_dir + '/' + batch_name,
            local_dir = local_batch_dir + '/data/Write_up')
        
        ppt_list = downloaded_files

        pd_pptx = spur_pptx_to_xlsx.pptx_to_xlsx(
            ppt_list=ppt_list,
            save_slide=True,
            job_blob_dir = local_batch_dir + '/Job_SPUR/BlobFiles',
            job_clob_dir = '',
            position_blob_dir= '',
            position_clob_dir = '',
        )


        for index, row in pd_pptx.iterrows():
            SPUR_code = row['UR_CODE']
            SPUR_name = row['UR_NAME']
            SKG_name = SPUR_code.split('-')[1]
            SPUR_file_storage_account_path = conf.management_job_spur_dir + '/SKG' + SKG_name + '/' + row['UR_CODE'] + '.pdf'
            purpose_and_accountability =  "<p>" + row["ROLEPURPOSE"].strip() + "</p>" + "\n\n\n" + row["ACCOUNTABILITIES"].strip()
            experience = row["EXPERIENCE"].strip()
            challenge = row["CHALLENGES"].strip()
            local_file_path = local_batch_dir + '/Job_SPUR/BlobFiles/' + row['UR_CODE'] + '.pdf'
            
            uploaded_file = upload_to_adls(
                service_client = service_client,
                container = conf.management_container_name,
                local_path = local_file_path,
                remote_path= SPUR_file_storage_account_path
            )[0]

            spur = {
                'SPUR_code': SPUR_code, 
                'SPUR_name': SPUR_name, 
                'SKG_name': SKG_name, 
                'SPUR_file_storage_account_path': uploaded_file,
                'purpose_and_accountability': purpose_and_accountability,
                'challenge': challenge,
                'experience': experience
            }
            pd_spur = pd_spur.append(spur, ignore_index=True)

        # Merge SPUR in database
        sql_insert(engine = sql_engine, df = pd_spur, table_name = batch_name, schema_name = 'Staging')
        sql_query = f'\
            UPDATE [dbo].[Batch] \
            SET \
                BatchStatus = \'Pending Submit\',\
                BatchProcessStatus = \'Completed\',\
                BackendUserModifiedBy = \'{conf.sql_user}\',\
                BackendUserModifiedTimestamp = GETUTCDATE()\
            WHERE BatchName = \'{batch_name}\'\
            \
            MERGE [dbo].[SPUR] A\
            USING (SELECT \
                    A.SPUR_code, \
                    A.SPUR_name, \
                    B.BatchID, \
                    B.SubmittedBy, \
                    B.SubmittedTimeStamp, \
                    A.SPUR_file_storage_account_path, \
                    A.purpose_and_accountability, \
                    A.challenge, \
                    A.experience \
                FROM [Staging].[{batch_name}] A \
                INNER JOIN [dbo].[Batch] B ON \'{batch_name}\' = B.BatchName) B\
            ON A.SPURCode = B.SPUR_code\
            WHEN MATCHED THEN\
            UPDATE SET BatchID = B.BatchID, BackendUserModifiedBy = \'{conf.sql_user}\', BackendUserModifiedTimestamp = GETUTCDATE()\
            WHEN NOT MATCHED THEN\
            INSERT (SPURCode, SPURName, BatchID, SPURFilePath, PurposeAndAccountability, Challenge, Experience, EndUserCreatedBy, SubmittedTimeStamp, EndUserCreatedTimestamp)\
            VALUES (B.SPUR_code, B.SPUR_name, B.BatchID, B.SPUR_file_storage_account_path, B.purpose_and_accountability, B.challenge, B.experience, B.SubmittedBy, B.SubmittedTimeStamp, B.SubmittedTimeStamp);\
            \
            EXEC [audit].[uspAddEmail] @pBatchID = {batch_id}\
            \
            DROP TABLE [Staging].[{batch_name}]'
        sql_execute(engine=sql_engine, query = sql_query)

        shutil.rmtree(local_batch_dir)
    except Exception as e:
        logging.error(f'Encountered error in create_spur.py: {str(e)}')
        raise ValueError(e)