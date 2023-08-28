import pandas as pd
from config import *
from Helper.sql_helper import *
from Helper.storage_account_helper import *
import spur_pptx_to_xlsx
import shutil
import traceback

try:
    #Capture non-processed batch
    pd_batch = sql_read(engine = SQL_ENGINE, query = 'EXEC [uspGetBatchForProcessing] @pBatchType = \'SPUR\'')
    LOGGER.info(f'{len(pd_batch.index)} new SPUR batches found.')
    if not pd_batch.empty:
        
        #Download thmx template if not exist
        if not os.path.isfile(THEME_FILE_PATH):
            downloaded_files = download_from_adls(
                service_client = DATA_LAKE_SERVICE_CLIENT, 
                container = MASTER_CONTAINER_NAME, 
                remote_path = 'PETRONAS.thmx',
                local_dir = MAIN_LOCAL_DIR + '/Master')
            LOGGER.info(f'Downloaded thmx template.')

        #Process each batch
        for index, row in pd_batch.iterrows():
            try:
                batch_id = row['BatchID']
                batch_name = row['BatchName']
                batch_intake_path = row['BatchIntakePath']
                pd_spur = pd.DataFrame(columns = ['SPUR_code', 'SPUR_name', 'SKG_name', 'SPUR_file_storage_account_path', 'purpose_and_accountability', 'challenge', 'experience', 'KPI'])
                local_batch_dir = MAIN_LOCAL_DIR + '/' + batch_name

                #Download files for a batch
                downloaded_files = download_from_adls(
                    service_client = DATA_LAKE_SERVICE_CLIENT, 
                    container = INTAKE_CONTAINER_NAME, 
                    remote_path = batch_intake_path.split('dfs.core.windows.net/')[1].split('/', 1)[1], #INTAKE_JOB_SPUR_DIR + '/' + batch_name,
                    local_dir = local_batch_dir + '/data/Write_up')
                LOGGER.info(f'Downloaded write up for {batch_name} SPUR batch.')
                
                ppt_list = downloaded_files

                pd_pptx = spur_pptx_to_xlsx.pptx_to_xlsx(
                    ppt_list=ppt_list,
                    save_slide=True,
                    job_blob_dir = local_batch_dir + '/Job_SPUR/BlobFiles'
                )
                LOGGER.info(f'Generated PDF files for {batch_name} SPUR batch.')

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
                    SPUR_file_storage_account_path = MANAGEMENT_JOB_SPUR_DIR + '/SKG' + SKG_name
                    
                    uploaded_file = upload_to_adls(
                        service_client = DATA_LAKE_SERVICE_CLIENT,
                        container = MANAGEMENT_CONTAINER_NAME,
                        local_path = local_file_path,
                        remote_path= SPUR_file_storage_account_path,
                        remote_file_name = f"{file_name.split('.')[0]}_{batch_name}.{file_name.split('.')[-1]}"
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
                LOGGER.info(f'Captured {len(pd_spur.index)} SPURs from {batch_name} SPUR batch.')

                # Merge SPUR in database
                sql_insert(engine = SQL_ENGINE, df = pd_spur, table_name = batch_name, schema_name = 'Staging')
                sql_query = f'\
                    INSERT INTO SPUR (SPURCode, SPURName, BatchID, SPURFilePath, PurposeAndAccountability, Challenge, Experience, KPI, EndUserCreatedBy, SubmittedTimeStamp, EndUserCreatedTimestamp)\
                    SELECT \
                            A.SPUR_code, \
                            A.SPUR_name, \
                            B.BatchID, \
                            A.SPUR_file_storage_account_path, \
                            A.purpose_and_accountability, \
                            A.challenge, \
                            A.experience, \
                            A.KPI, \
                            B.SubmittedBy, \
                            B.SubmittedTimeStamp, \
                            B.SubmittedTimeStamp \
                        FROM [Staging].[{batch_name}] A \
                        INNER JOIN [dbo].[Batch] B ON \'{batch_name}\' = B.BatchName\
                    \
                    UPDATE [dbo].[Batch] \
                    SET \
                        BatchStatus = \'Pending Submit\',\
                        BatchProcessStatus = \'Completed\',\
                        BatchLog = CONCAT(\'{len(pd_pptx.index)} valid SPURs have been successfully captured:\',  \
                            ISNULL(STUFF((SELECT \'<li>\' + SPURCode + \'-\' + SPURName + \'</li>\'\
                                FROM (SELECT B.SPURCode, B.SPURName FROM Batch A LEFT JOIN SPUR B ON A.BatchID = B.BatchID AND A.BatchName = \'{batch_name}\') T\
                                FOR XML PATH(\'\'), TYPE).value(\'.\', \'NVARCHAR(MAX)\'), 1, 0, \'\'), \'\')),\
                        BackendUserModifiedBy = \'{SQL_USERNAME}\',\
                        BackendUserModifiedTimestamp = GETUTCDATE()\
                    WHERE BatchName = \'{batch_name}\'\
                    \
                    EXEC [audit].[uspAddEmail] @pBatchID = {batch_id}\
                    \
                    DROP TABLE [Staging].[{batch_name}]'
                sql_execute(engine=SQL_ENGINE, query = sql_query)
                LOGGER.info(f'Loaded len{pd_spur.index} SPURs in the system from {batch_name} SPUR batch.')
            except Exception as e:
                LOGGER.exception(f'Encountered error in capture_spur.py while processing batch.')
            finally:
                if 'local_batch_dir' in locals():
                    shutil.rmtree(local_batch_dir)
except Exception as e:
    LOGGER.exception(f'Encountered error in capture_spur.py: {traceback.format_exc()}')
    raise ValueError(e)