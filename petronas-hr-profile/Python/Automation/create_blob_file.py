import os
from Helper.storage_account_helper import copy_file_in_adls
from config import *

def create_blob_file(
    service_client,
    df,
    destination_remote_path,
    destination_file_name_field
):
    try:
        for i in range(len(df)):
            copy_file_in_adls(
                service_client=service_client,
                source_container = MANAGEMENT_CONTAINER_NAME,
                source_remote_path = df.iloc[i]["SPURFilePath"].split(MANAGEMENT_CONTAINER_NAME)[1],
                destination_container = CONSUMPTION_CONTAINER_NAME,
                destination_remote_path = destination_remote_path + '/' + os.path.basename(df.iloc[i][destination_file_name_field] + '_' + df.iloc[i]['UR_CODE'] +'.pdf')
            )
    except Exception as e:
        raise ValueError(e)