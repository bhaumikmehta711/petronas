import os
from Helper.storage_account_helper import copy_file_in_adls
import config as conf

def create_blob_file(
    service_client,
    df,
    destination_remote_path
):
    for i in range(len(df)):
        copy_file_in_adls(
            service_client=service_client,
            source_container = conf.management_container_name,
            source_remote_path = df.iloc[i]["SPURFilePath"].split(conf.management_container_name)[1],
            destination_container = conf.consumption_container_name,
            destination_remote_path = destination_remote_path + '/' + os.path.basename(df.iloc[i]["SPURFilePath"])
        )