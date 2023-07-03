import os
import urllib.parse

def download_from_adls(service_client, container, remote_path, local_dir, new_file_name = None):
    file_system_client = service_client.get_file_system_client(file_system = container)
    paths = file_system_client.get_paths(path = remote_path)
    directory_client = file_system_client.get_directory_client('/')
    downloaded_files = []
    for path in paths:
        #Create folder if not exists
        local_file_path = local_dir + '/' + new_file_name if new_file_name else local_dir + '/' + os.path.basename(path.name)
        os.makedirs(local_dir, exist_ok=True)
        
        local_file = open(local_file_path,'wb')
        file_client = directory_client.get_file_client(path.name)
        download = file_client.download_file()
        downloaded_bytes = download.readall()
        local_file.write(downloaded_bytes)
        local_file.close()
        downloaded_files.append(local_file_path)
    return downloaded_files

def upload_to_adls(service_client, container, remote_path, local_path, remote_file_name = None):
    file_list = []
    primary_endpoints = []
    file_system_client = service_client.get_file_system_client(file_system = container)
    directory_client = file_system_client.get_directory_client("/")

    if os.path.isfile(local_path):
        file_list.append(local_path)
    else:
        for root, dirs, files in os.walk(local_path):
            for file in files:
                file_list.append(os.path.join(root,file))

    for file in file_list:
        if not remote_file_name:
            file_path = remote_path + '\\' + os.path.relpath(file, local_path)
        else:
            file_path = remote_path + '\\' + remote_file_name
        file_client = directory_client.get_file_client(file_path)
        local_file = open(file,'rb')
        file_contents = local_file.read()
        file_client.upload_data(file_contents, overwrite=True)
        primary_endpoints.append(urllib.parse.unquote(file_client.primary_endpoint))
    return primary_endpoints

def copy_file_in_adls(service_client, source_container, source_remote_path, destination_container, destination_remote_path):
    source_file_system_client = service_client.get_file_system_client(file_system=source_container)
    directory_client = source_file_system_client.get_directory_client('/')    
    source_file_client = directory_client.get_file_client(source_remote_path)
    download = source_file_client.download_file()
    downloaded_bytes = download.readall()

    destination_file_system_client = service_client.get_file_system_client(file_system=destination_container)
    directory_client = destination_file_system_client.get_directory_client('/')
    destination_file_client = directory_client.get_file_client(destination_remote_path)
    destination_file_client.upload_data(downloaded_bytes, overwrite=True)