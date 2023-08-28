from azure.identity import ChainedTokenCredential, ManagedIdentityCredential, AzureCliCredential
from azure.keyvault.secrets import SecretClient
import logging
from opencensus.ext.azure.log_exporter import AzureLogHandler
from azure.storage.filedatalake import DataLakeServiceClient
from sqlalchemy import create_engine
import urllib

try:
    #Define global variables
    MAIN_LOCAL_DIR = r'C:\Users\bhaumikmehta\OneDrive - Microsoft\Desktop\Project\Petronas\Debug' #r'C:\Users\hradmin\hrprofile'
    STORAGE_ACCOUNT_NAME = 'adlspetronashrprofiledev'
    MULESOFT_API_URL = 'https://devapi.petronas.com/dev/pet/web/hrempmgmt/exp/api/v1/employees'
    THEME_FILE_PATH = MAIN_LOCAL_DIR + r'\Master\PETRONAS.thmx'
    MASTER_CONTAINER_NAME = 'master'
    CONSUMPTION_CONTAINER_NAME = 'consumption'
    INTAKE_CONTAINER_NAME = 'intake'
    INTAKE_JOB_SPUR_DIR = 'spur'
    INTAKE_JOB_POSITION_DIR = 'position'
    MANAGEMENT_JOB_SPUR_DIR = 'job_spur'
    MANAGEMENT_CONTAINER_NAME = 'management'
    REFERENCE_CONTAINER_NAME = 'reference'

    #Get secret from Key Vault
    credential = ChainedTokenCredential(ManagedIdentityCredential(), AzureCliCredential())
    SECRET_CLIENT = SecretClient(vault_url="https://kv-ptrns-hrprofile-dev-1.vault.azure.net/", credential=credential)
    STORAGE_ACCOUNT_KEY = SECRET_CLIENT.get_secret("storage-account-key").value
    MULESOFT_API_SECRET = SECRET_CLIENT.get_secret("mulesoft-api-secret").value
    SQL_SERVER = SECRET_CLIENT.get_secret("sql-server").value
    SQL_DATABASE = SECRET_CLIENT.get_secret("sql-database").value
    SQL_USERNAME = SECRET_CLIENT.get_secret("sql-username").value
    SQL_PASSWORD = SECRET_CLIENT.get_secret("sql-password").value
    INSTRUMENTATION_KEY = SECRET_CLIENT.get_secret("instrumentation-key").value

    #Storage client
    DATA_LAKE_SERVICE_CLIENT = DataLakeServiceClient(account_url="{}://{}.dfs.core.windows.net".format("https", STORAGE_ACCOUNT_NAME), credential = STORAGE_ACCOUNT_KEY)

    #SQL client
    SQL_CONNECTION_STRING = f'DRIVER={{ODBC Driver 18 for SQL Server}};SERVER=tcp:{SQL_SERVER};DATABASE={SQL_DATABASE};UID={SQL_USERNAME};PWD={SQL_PASSWORD}'
    SQL_ENGINE = create_engine("mssql+pyodbc:///?autocommit=true&odbc_connect=%s" % urllib.parse.quote_plus(SQL_CONNECTION_STRING))

    #Log client
    LOGGER = logging.getLogger(__name__)
    APP_INSIGHT_CONNECTION_STRING = f'InstrumentationKey={INSTRUMENTATION_KEY}'
    LOGGER.addHandler(AzureLogHandler(connection_string=APP_INSIGHT_CONNECTION_STRING))
    LOGGER.setLevel(logging.INFO)
except Exception as e:
    raise ValueError(e)