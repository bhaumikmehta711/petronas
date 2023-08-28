
import json
import shutil
import glob
import spur_position_profile
import spur_job_profile
import os
from datetime import datetime 
from config import *
import spur_data_processor_sql
import position_data_processor_sql
from Helper.sql_helper import *
from Helper.storage_account_helper import *
from Helper.key_vault_helper import *
from Helper.web_helper import *
import create_clob_file
import create_blob_file
import spur_pd_processor
import traceback

try:
    #Initialization
    process_datetime = datetime.utcnow().strftime('%Y%m%d%H%M')
    consumption_dir =  MAIN_LOCAL_DIR + '\\' + process_datetime
    batch_consumption_path = 'https://'+ STORAGE_ACCOUNT_NAME + '.blob.core.windows.net/' + CONSUMPTION_CONTAINER_NAME+'/' + datetime.utcnow().strftime('%Y/%m/%d')
    LOGGER.info(f'Process started for {process_datetime} consumption batch.')

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
    pd_dir = data_dir + "\\" + "PD" + "\\"

    spur_details_file_path = data_dir + "\\" + "final_processed_data\\{}_details.xlsx".format(process_datetime)
    spur_position_file_path = data_dir + "\\final_processed_data\\{}_position_profile_data.xlsx".format(process_datetime)

    if not os.path.exists(consumption_dir):
        os.makedirs(consumption_dir)
        os.makedirs(data_dir)
        os.makedirs(data_dir + "\\" + "final_processed_data")
        os.makedirs(data_dir + "\\" + "Write_up")
        os.makedirs(data_dir + "\\" + "PD")
        os.makedirs(log_dir)
        os.makedirs(job_dir)
        os.makedirs(job_clob_dir)
        os.makedirs(job_dat_dir)
        os.makedirs(position_dir)
        os.makedirs(position_blob_dir)
        os.makedirs(position_clob_dir)
        os.makedirs(position_dat_dir)

    LOGGER.info(f'Folders created for {process_datetime} consumption batch.')

    #region SPUR
    #Load approved SPUR
    sql_execute(SQL_ENGINE, f'EXEC uspGetSPURBatchForConsumption @pStagingTableName = \'SPUR_{process_datetime}\'')
    spur_df = sql_read(SQL_ENGINE, f'SELECT * FROM [Staging].[SPUR_{process_datetime}]')
    LOGGER.info(f'Captured {len(spur_df.index)} SPUR batches for {process_datetime} consumption batch.')

    if not spur_df.empty:
        #Download SPUR excel template
        download_from_adls(
            service_client = DATA_LAKE_SERVICE_CLIENT, 
            container = MASTER_CONTAINER_NAME, 
            remote_path = 'PET_Job SPUR.xlsx',
            local_dir = job_dir,
            new_file_name = f'PET_Job_SPUR_{process_datetime}.xlsx')
        LOGGER.info(f'Downloaded SPUR template for {process_datetime} consumption batch.')

        #Load SPUR staging data
        spur_data_processor_sql.data_processor(
            consumption_dir=consumption_dir,
            process_datetime=process_datetime,
            sql_engine = SQL_ENGINE
        ).spur_data()
        job_template_file_path = os.path.abspath(glob.glob(job_dir + "\\" + "*.xlsx")[0])
        LOGGER.info(f'Prepared SPUR staging data for {process_datetime} consumption batch.')

        #Create dat files
        spur_job_profile.spur_job_profile(
                job_template_file_path=job_template_file_path,
                spur_df=spur_df,
                spur_details_file_path=spur_details_file_path,
                job_dat_dir=job_dat_dir
            )
        LOGGER.info(f'Loaded PET Job SPUR excel and created SPUR dat files for {process_datetime} consumption batch.')
        
        os.rename(job_template_file_path, job_template_file_path.replace('_' + process_datetime, ''))
        LOGGER.info(f'Renamed PET Job SPUR excel for {process_datetime} consumption batch.')

        #Create clob files
        create_clob_file.create_clob_file(df = spur_df, clob_folder_path = job_clob_dir, destination_file_name_field = 'ProfileCode')
        LOGGER.info(f'Created SPUR clob files for {process_datetime} consumption batch.')

        #Upload dat and blob files
        uploaded_files = upload_to_adls(
                service_client = DATA_LAKE_SERVICE_CLIENT,
                container = CONSUMPTION_CONTAINER_NAME,
                local_path = job_dir,
                remote_path= datetime.utcnow().strftime('%Y/%m/%d') + '/Job_SPUR'
            )
        LOGGER.info(f'Uploaded SPUR dat, clob and excel files for {process_datetime} consumption batch.')
        
        #Copy blob files
        create_blob_file.create_blob_file(service_client = DATA_LAKE_SERVICE_CLIENT, df = spur_df, destination_remote_path = datetime.utcnow().strftime('%Y/%m/%d') + '/Job_SPUR/BlobFiles', destination_file_name_field = 'ProfileCode')
        LOGGER.info(f'Created SPUR blob files for {process_datetime} consumption batch.')
    #endregion

    #region Position
    #Load approved position
    sql_execute(SQL_ENGINE, f'EXEC uspGetPositionBatchForConsumption @pStagingTableName = \'Position_{process_datetime}\'')
    position_df = sql_read(SQL_ENGINE, f'SELECT * FROM [Staging].[Position_{process_datetime}]')
    LOGGER.info(f'Captured {len(position_df.index)} position batches for {process_datetime} consumption batch.')

    if not position_df.empty:
        #Download position excel template
        download_from_adls(
            service_client = DATA_LAKE_SERVICE_CLIENT, 
            container = MASTER_CONTAINER_NAME, 
            remote_path = 'PET_Position SPUR.xlsx',
            local_dir = position_dir,
            new_file_name = f'PET_Position_SPUR_{process_datetime}.xlsx')
        LOGGER.info(f'Downloaded position template for {process_datetime} consumption batch.')

        #Load position staging data
        position_data_processor_sql.data_processor(
            consumption_dir=consumption_dir,
            process_datetime=process_datetime,
            sql_engine = SQL_ENGINE
        ).position_data()
        position_template_file_path = os.path.abspath(glob.glob(position_dir + "\\" + "*.xlsx")[0])
        LOGGER.info(f'Prepared position staging data for {process_datetime} consumption batch.')
        
        management_container_path = f'https://{STORAGE_ACCOUNT_NAME}.dfs.core.windows.net/' + MANAGEMENT_CONTAINER_NAME
        for index, row in position_df[(position_df.PDFilePath is not None) & (position_df.PDFilePath != '')].iterrows():
            remote_file_path = row['PDFilePath'].replace(management_container_path, '')
            download_from_adls(
                service_client = DATA_LAKE_SERVICE_CLIENT, 
                container = MANAGEMENT_CONTAINER_NAME, 
                remote_path = remote_file_path,
                local_dir = pd_dir,
                new_file_name=row['PositionProfileCode'] + '_' + row['PositionCode'] + os.path.splitext(remote_file_path)[1]
            )
        LOGGER.info(f"Downloaded {len(position_df[(position_df.PDFilePath is not None) & (position_df.PDFilePath != '')].index)} PD files for {process_datetime} consumption batch.")
        
        spur_pd_processor.pd_processor(
            position_blob_dir=position_blob_dir,
            pd_folder=pd_dir,
        )

        #Create dat files
        spur_position_profile.spur_position_profile(
            position_blob_dir=position_blob_dir,
            position_template_file_path=position_template_file_path,
            position_profile_df = position_df,
            position_details_file_path = spur_position_file_path,
            position_dat_dir=position_dat_dir
        )
        LOGGER.info(f'Loaded PET Position SPUR excel and created position dat files for {process_datetime} consumption batch.')

        os.rename(position_template_file_path, position_template_file_path.replace('_' + process_datetime, ''))
        LOGGER.info(f'Renamed PET Position SPUR excel for {process_datetime} consumption batch.')

        #Create clob files
        create_clob_file.create_clob_file(df = position_df, clob_folder_path = position_clob_dir, destination_file_name_field = 'PositionProfileCode')
        LOGGER.info(f'Created position clob files for {process_datetime} consumption batch.')

        #Upload dat and blob files
        uploaded_files = upload_to_adls(
                service_client = DATA_LAKE_SERVICE_CLIENT,
                container = CONSUMPTION_CONTAINER_NAME,
                local_path = position_dir,
                remote_path= datetime.utcnow().strftime('%Y/%m/%d') + '/Job_Position'
            )
        LOGGER.info(f'Uploaded position dat, clob and excel files for {process_datetime} consumption batch.')

        #Copy blob files
        create_blob_file.create_blob_file(service_client = DATA_LAKE_SERVICE_CLIENT, df = position_df, destination_remote_path = datetime.utcnow().strftime('%Y/%m/%d') + '/Job_Position/BlobFiles', destination_file_name_field = 'PositionProfileCode')
        LOGGER.info(f'Created position blob files for {process_datetime} consumption batch.')
    #endregion

    sql_execute(SQL_ENGINE, f'UPDATE A\
            SET \
                A.BatchCompleteStatus = \'Completed\',\
                A.BackendUserModifiedBy = \'{SQL_USERNAME}\',\
                A.BackendUserModifiedTimeStamp = GETUTCDATE(),\
                A.BatchConsumptionPath = \'{batch_consumption_path}\'\
            FROM Batch A\
            INNER JOIN (SELECT BatchID FROM Staging.[SPUR_{process_datetime}] UNION SELECT BatchID FROM Staging.[Position_{process_datetime}])  B ON A.BatchID = B.BatchID\
            DROP TABLE Staging.[Position_{process_datetime}]\
            DROP TABLE Staging.[SPUR_{process_datetime}]')
    LOGGER.info(f'Updated data for SPUR and Position batches for {process_datetime} consumption batch.')

    #region Notify Mulesoft
    pd_consumption_path = pd.DataFrame(sql_read(SQL_ENGINE, f'SELECT DISTINCT BatchConsumptionPath url FROM Batch WHERE ISNULL(MulesoftNotifiedFlag, 0) = 0 AND ISNULL(BatchConsumptionPath,\'\') != \'\' AND BatchCompleteStatus = \'Completed\''))
    LOGGER.info(f'Captured {len(pd_consumption_path.index)} consumption batches to notify mulesoft.')

    if not pd_consumption_path.empty:
        consumption_paths = pd_consumption_path.to_dict('records')
        response = post(
            url = MULESOFT_API_URL,
            header = {'Authorization' : f'Basic {MULESOFT_API_SECRET}','Content-Type' : 'application/json'},
            payload=consumption_paths
        )
        LOGGER.info(f'Executed Mulesoft API.')

        if response.status_code in range(200, 300):
            pd_api_result = pd.DataFrame(response.json())
            pd_api_failure_result = pd_api_result[~pd_api_result['status']]
            if not pd_api_failure_result.empty:
              LOGGER.warning(f"Following paths are not found at mulesoft: {', '.join(pd_api_failure_result['url'])}")

            sql_insert(engine = SQL_ENGINE, df = pd_api_result, table_name = f'Consumption_{process_datetime}', schema_name = 'Staging')
            sql_execute(SQL_ENGINE, f'UPDATE A\
                    SET \
                        A.MulesoftNotifiedFlag = 1,\
                        A.MulesoftAcknowledgedFlag = B.status, \
                        A.BackendUserModifiedBy = \'{SQL_USERNAME}\',\
                        A.BackendUserModifiedTimeStamp = GETUTCDATE()\
                    FROM Batch A \
                    INNER JOIN [Staging].[Consumption_{process_datetime}] B\
                    ON A.BatchConsumptionPath = B.url\
                    DROP TABLE [Staging].[Consumption_{process_datetime}]\
                ')
            LOGGER.info(f'Updated mulesoft related flags for SPUR and position batches.')
        else:
            LOGGER.info(f'Encountered an error in mulesoft API: {response.text}')
    #endregion
except Exception as e:
    LOGGER.exception(f'Encountered error in consumption_data_processor.py: {traceback.format_exc()}')
    raise ValueError(e)
finally:
    shutil.rmtree(consumption_dir)