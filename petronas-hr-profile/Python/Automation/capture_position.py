import pandas as pd
from config import *
from Helper.sql_helper import *
from Helper.storage_account_helper import *
import shutil
import traceback

try:
    #Initialization
    destination_file_path = []

    #Capture non-processed batch
    pd_batch = sql_read(engine = SQL_ENGINE, query = 'EXEC [uspGetBatchForProcessing] @pBatchType = \'Position\'')
    LOGGER.info(f'{len(pd_batch.index)} new position batches found.')

    #Process each batch
    for index, position in pd_batch.iterrows():
        try:
            batch_id = position['BatchID']
            batch_name = position['BatchName']
            batch_intake_path = position['BatchIntakePath']
            pd_position = pd.DataFrame(columns = ['maintenance_mode', 'effective_start_date', 'effective_end_date', 'position_profile_code', 'SPUR_code', 'position_name', 'job_grade', 'role_level', 'company_code'])
            
            local_batch_dir = MAIN_LOCAL_DIR + '/' + batch_name

            #Download files for a batch
            downloaded_files = download_from_adls(service_client = DATA_LAKE_SERVICE_CLIENT, 
                container = INTAKE_CONTAINER_NAME, 
                remote_path = batch_intake_path.split('dfs.core.windows.net/')[1].split('/', 1)[1], #INTAKE_JOB_POSITION_DIR + '/' + batch_name,
                local_dir = local_batch_dir + '/data/Write_up')
            LOGGER.info(f'Downloaded write up for {batch_name} position batch.')
            
            position_writeups = downloaded_files
            for position_writeup in position_writeups:
                pd_writeup = pd.read_excel(position_writeup, sheet_name='Position Profile Maintenance', skiprows=[1], dtype = {'Company Code': str, 'Position ID': str})
                for index, position in pd_writeup.iterrows():                
                    current_position = {}
                    current_position['maintenance_mode'] = position['Maintenance Mode']
                    current_position['effective_start_date'] = position['Effective Start Date']
                    current_position['effective_end_date'] = position['Effective End Date']
                    current_position['position_profile_code'] = position['Position Profile Code']
                    current_position['position_code'] = position['Position ID']
                    current_position['SPUR_code'] = position['SPUR ID']
                    current_position['position_name'] = position['Position Name'] 
                    current_position['job_grade'] = position['Job Grade']
                    current_position['role_level'] = position['Role Level']
                    current_position['company_code'] = str(position['Company Code']).zfill(4)
                    pd_position = pd_position.append(current_position, ignore_index = True)
            pd_position.drop_duplicates(subset='position_profile_code', keep="first")
            LOGGER.info(f'Captured {len(pd_position.index)} positions from {batch_name} position batch.')

            # Merge position in database
            if not pd_position.empty:
                sql_insert(engine = SQL_ENGINE, df = pd_position, table_name = batch_name, schema_name = 'Staging', if_exists='replace')
                sql_query = f'\
                    INSERT INTO [dbo].[Position] \
                    (\
                          MaintenanceMode\
                        , PositionProfileCode\
                        , PositionName\
                        , PositionDescription\
                        , EffectiveStartDate\
                        , EffectiveEndDate\
                        , PositionCode\
                        , SPURCode\
                        , SPURName\
                        , JobGradeID\
                        , SPURFilePath\
                        , RoleLevelID\
                        , CompanyID\
                        , PurposeAndAccountability\
                        , Challenge\
                        , Experience\
                        , KPI\
                        , BatchID\
                        , SubmittedTimeStamp\
                        , EndUserCreatedBy\
                        , EndUserCreatedTimestamp\
                    )\
                    SELECT \
                          A.maintenance_mode\
                        , A.position_profile_code\
                        , A.position_name\
                        , A.position_name\
                        , CONVERT(DATETIME, A.effective_start_date, 103)\
                        , CONVERT(DATETIME, A.effective_end_date, 103)\
                        , A.position_code\
                        , A.SPUR_code \
                        , F.SPURName\
                        , D.JobGradeID\
                        , F.SPURFilePath\
                        , E.RoleLevelID\
                        , C.CompanyID\
                        , F.PurposeAndAccountability\
                        , F.Challenge\
                        , F.Experience\
                        , F.KPI\
                        , B.BatchID\
                        , B.SubmittedTimeStamp\
                        , B.SubmittedBy\
                        , B.SubmittedTimeStamp\
                    FROM [Staging].[{batch_name}] A \
                    INNER JOIN [dbo].[Batch] B ON \'{batch_name}\' = B.BatchName AND A.maintenance_mode = \'Create\'\
                    INNER JOIN Master.Company C ON C.CompanyCode = A.company_code\
                    LEFT JOIN Master.JobGrade D ON D.JobGradeName = A.job_grade\
                    INNER JOIN Master.RoleLevel E ON E.RoleLevelName = A.role_level\
                    INNER JOIN Ref.SPUR F ON F.SPURCode = A.SPUR_code AND F.RoleLevelID = E.RoleLevelID AND F.Status = \'A\'\
                    LEFT JOIN Ref.Position G ON G.PositionProfileCode = position_profile_code\
                    WHERE G.PositionID IS NULL\
                    \
                    INSERT INTO [dbo].[Position] (\
                        PositionCode\
                        , PositionName\
                        , PositionDescription\
                        , MaintenanceMode\
                        , JobGradeID\
                        , SPURCode\
                        , SPURName\
                        , CompanyID\
                        , EffectiveStartDate\
                        , EffectiveEndDate\
                        , PurposeAndAccountability\
                        , Challenge\
                        , Experience\
                        , KPI\
                        , BatchID\
                        , SPURFilePath\
                        , RoleLevelID\
                        , PositionProfileCode\
                        , PDFilePath\
                        , SubmittedTimeStamp\
                        , EndUserCreatedBy\
                        , EndUserCreatedTimestamp\
                    )\
                    SELECT \
                          C.PositionCode,\
                        , C.positionName\
                        , C.PositionDescription\
                        , A.maintenance_mode\
                        , COALESCE(D.JobGradeID, C.JobGradeID)\
                        , C.SPURCode\
                        , C.SPURName\
                        , C.CompanyID\
                        , C.EffectiveStartDate\
                        , C.EffectiveEndDate\
                        , C.PurposeAndAccountability\
                        , C.Challenge\
                        , C.Experience\
                        , C.KPI\
                        , B.BatchID\
                        , C.SPURFilePath\
                        , C.RoleLevelID\
                        , C.PositionProfileCode\
                        , C.PDFilePath\
                        , B.SubmittedTimeStamp\
                        , B.SubmittedBy\
                        , B.EndUserCreatedTimestamp\
                    FROM [Staging].[{batch_name}] A \
                    INNER JOIN [dbo].[Batch] B ON \'{batch_name}\' = B.BatchName AND A.maintenance_mode IN (\'Update\', \'Delimit\')\
                    INNER JOIN Ref.Position C ON C.PositionProfileCode = A.position_profile_code\
                    LEFT JOIN Master.JobGrade D ON D.JobGradeName = A.job_grade\
                    \
                    EXEC [dbo].[uspLoadPosition] @pBatchID = {batch_id}'
                sql_execute(engine=SQL_ENGINE, query = sql_query)

                pd_file = sql_read(engine = SQL_ENGINE, query = f'SELECT PositionID, PositionProfileCode, MaintenanceMode, SPURFilePath, PDFilePath FROM Position A INNER JOIN BATCH B ON A.BatchID = B.BatchID WHERE BatchName = \'{batch_name}\'')
                for index, row in pd_file.iterrows():
                    destination_spur_file_path = ''
                    destination_pd_file_path = ''
                    position_id = row['PositionID']
                    position_profile_code = row['PositionProfileCode']
                    SPUR_code = position_profile_code.split('_')[0]
                    skg = SPUR_code.split('-')[0]
                    company_code = position_profile_code.split('_')[1]
                    position_code = position_profile_code.split('_')[2]
                    
                    if row['SPURFilePath'] is not None and row['SPURFilePath'] != '':
                        destination_spur_file_path = f'https://{STORAGE_ACCOUNT_NAME}.dfs.core.windows.net/{MANAGEMENT_CONTAINER_NAME}/Position/SKG{skg}/{SPUR_code}/{company_code}/{position_code}/{position_profile_code}_{SPUR_code}.pdf'
                        copy_file_in_adls(
                            service_client=DATA_LAKE_SERVICE_CLIENT,
                            source_container = REFERENCE_CONTAINER_NAME,
                            source_remote_path = row["SPURFilePath"].split(REFERENCE_CONTAINER_NAME)[1],
                            destination_container = MANAGEMENT_CONTAINER_NAME,
                            destination_remote_path = destination_spur_file_path.split(MANAGEMENT_CONTAINER_NAME)[1]
                        )
                    if row['PDFilePath'] is not None and row['PDFilePath'] != '':
                        destination_pd_file_path = f'https://{STORAGE_ACCOUNT_NAME}.dfs.core.windows.net/{MANAGEMENT_CONTAINER_NAME}/Position/SKG{skg}/{SPUR_code}/{company_code}/{position_code}/{position_profile_code}_{position_code}.pdf'
                        copy_file_in_adls(
                            service_client=DATA_LAKE_SERVICE_CLIENT,
                            source_container = REFERENCE_CONTAINER_NAME,
                            source_remote_path = row["PDFilePath"].split(REFERENCE_CONTAINER_NAME)[1],
                            destination_container = MANAGEMENT_CONTAINER_NAME,
                            destination_remote_path = destination_pd_file_path.split(MANAGEMENT_CONTAINER_NAME)[1]
                        )
                    destination_file_path.append({'PositionID': position_id, 'SPURFilePath': destination_spur_file_path, 'PDFilePath': destination_pd_file_path})

                df_position_path = pd.DataFrame(destination_file_path)
                sql_insert(engine = SQL_ENGINE, df = df_position_path, table_name = 'file_' + batch_name, schema_name = 'Staging', if_exists='replace')
                sql_execute(engine = SQL_ENGINE, query = f'\
                    UPDATE A \
                    SET \
                        SPURFilePath = B.SPURFilePath,\
                        PDFilePath = B.PDFilePath,\
                        BackendUserModifiedBy = \'{SQL_USERNAME}\',\
                        BackendUserModifiedTimestamp = GETUTCDATE()\
                    FROM [Position] A\
                    INNER JOIN Staging.[file_{batch_name}] B ON A.PositionID = B.PositionID;\
                    \
                    UPDATE [dbo].[Batch] \
                    SET \
                        BatchStatus = \'Pending Submit\',\
                        BatchProcessStatus = \'Completed\',\
                        BatchLog = CONCAT(\'Out of total {len(pd_position.index)} uploaded positions, the following positions have been successfully captured:\',  \
                            ISNULL(STUFF((SELECT \'<li>\' + MaintenanceMode + \' - \' + PositionProfileCode + \'</li>\'\
                                FROM (SELECT B.PositionProfileCode, B.MaintenanceMode FROM Batch A LEFT JOIN Position B ON A.BatchID = B.BatchID AND A.BatchName = \'{batch_name}\') T\
                                FOR XML PATH(\'\'), TYPE).value(\'.\', \'NVARCHAR(MAX)\'), 1, 0, \'\'), \'\')),\
                        BackendUserModifiedBy = \'{SQL_USERNAME}\',\
                        BackendUserModifiedTimestamp = GETUTCDATE()\
                    WHERE BatchName = \'{batch_name}\'\
                    \
                    \
                    EXEC [audit].[uspAddEmail] @pBatchID = {batch_id}\
                    \
                    DROP TABLE [Staging].[file_{batch_name}]\
                    DROP TABLE [Staging].[{batch_name}]\
                ')
                LOGGER.info(f'Loaded positions in the system from {batch_name} position batch.')
        except Exception as e:
            LOGGER.exception(f'Encountered error in capture_position.py while processing batch.')
        finally:
            if 'local_batch_dir' in locals():
                shutil.rmtree(local_batch_dir)
except Exception as e:
    LOGGER.exception(f'Encountered error in capture_position.py: {traceback.format_exc()}')
    raise ValueError(e)