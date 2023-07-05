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

#Capture non-processed batch
pd_batch = sql_read(engine = sql_engine, query = 'EXEC [uspGetBatchForProcessing] @pBatchType = \'SPUR\'')

if not pd_batch.empty:
    
    service_client = DataLakeServiceClient(account_url="{}://{}.dfs.core.windows.net".format("https", conf.storage_account_name), credential=conf.storage_account_key)
    
    #Download thmx template if not exist
    if not os.path.isfile(conf.theme_file_path):
        downloaded_files = download_from_adls(
            service_client = service_client, 
            container = conf.master_container_name, 
            remote_path = 'PETRONAS.thmx',
            local_dir = conf.main_local_dir + '/Master')

    #Process each batch
    for index, row in pd_batch.iterrows():
        try:
            batch_id = row['BatchID']
            batch_name = row['BatchName']
            pd_spur = pd.DataFrame(columns = ['SPUR_code', 'SPUR_name', 'SKG_name', 'SPUR_file_storage_account_path', 'purpose_and_accountability', 'challenge', 'experience', 'KPI'])
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
                job_blob_dir = local_batch_dir + '/Job_SPUR/BlobFiles'
            )


            for index, row in pd_pptx.iterrows():
                SPUR_code = row['UR_CODE']
                SPUR_name = row['UR_NAME']
                SKG_name = SPUR_code.split('-')[1]
                purpose_and_accountability =  "<p>" + row["ROLEPURPOSE"].strip() + "</p>" + "\n\n\n" + row["ACCOUNTABILITIES"].strip()
                experience = row["EXPERIENCE"].strip()
                challenge = row["CHALLENGES"].strip()
                KPI = row["KPI"].strip()
                file_name = row['UR_CODE'] + '.pdf'
                local_file_path = local_batch_dir + '/Job_SPUR/BlobFiles/' + file_name
                SPUR_file_storage_account_path = conf.management_job_spur_dir + '/SKG' + SKG_name
                
                uploaded_file = upload_to_adls(
                    service_client = service_client,
                    container = conf.management_container_name,
                    local_path = local_file_path,
                    remote_path= SPUR_file_storage_account_path,
                    remote_file_name = file_name
                )[0]

                spur = {
                    'SPUR_code': SPUR_code, 
                    'SPUR_name': SPUR_name, 
                    'SKG_name': SKG_name, 
                    'SPUR_file_storage_account_path': uploaded_file,
                    'purpose_and_accountability': purpose_and_accountability,
                    'challenge': challenge,
                    'experience': experience,
                    'KPI': KPI
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
                        A.KPI, \
                        A.experience \
                    FROM [Staging].[{batch_name}] A \
                    INNER JOIN [dbo].[Batch] B ON \'{batch_name}\' = B.BatchName) B\
                ON 1 != 1\
                WHEN MATCHED THEN\
                UPDATE SET BatchID = B.BatchID, BackendUserModifiedBy = \'{conf.sql_user}\', BackendUserModifiedTimestamp = GETUTCDATE()\
                WHEN NOT MATCHED THEN\
                INSERT (SPURCode, SPURName, BatchID, SPURFilePath, PurposeAndAccountability, Challenge, Experience, KPI, EndUserCreatedBy, SubmittedTimeStamp, EndUserCreatedTimestamp)\
                VALUES (B.SPUR_code, B.SPUR_name, B.BatchID, B.SPUR_file_storage_account_path, B.purpose_and_accountability, B.challenge, B.experience, B.KPI, B.SubmittedBy, B.SubmittedTimeStamp, B.SubmittedTimeStamp);\
                \
                EXEC [audit].[uspAddEmail] @pBatchID = {batch_id}\
                \
                DROP TABLE [Staging].[{batch_name}]'
            sql_execute(engine=sql_engine, query = sql_query)

            
        except Exception as e:
            logging.error(f'Encountered error in create_spur.py: {str(e)}')
            raise ValueError(e)
        finally:
            shutil.rmtree(local_batch_dir)