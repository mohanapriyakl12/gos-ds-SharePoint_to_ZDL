# Databricks notebook source
 !pip install office365

# COMMAND ----------

!pip install office365-REST-Python-Client

# COMMAND ----------

!pip install google-cloud-secret-manager

# COMMAND ----------

from google.cloud import secretmanager

client = secretmanager.SecretManagerServiceClient()
secret = 'projects/1018680839633/secrets/zbr-sharepoint-gscr-secret/versions/latest'
response = client.access_secret_version(request={"name":secret})
credentials = eval(response.payload.data.decode("UTF-8"))
user_sc = credentials["username"]
password_sc = credentials["password"]

# COMMAND ----------

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File 
import os
import re

# Retrieve credentials from Google Secret Manager
from google.cloud import secretmanager

client = secretmanager.SecretManagerServiceClient()
secret = 'projects/1018680839633/secrets/zbr-sharepoint-gscr-secret/versions/latest'
response = client.access_secret_version(request={"name": secret})
credentials = eval(response.payload.data.decode("UTF-8"))
user_sc = credentials["username"]
password_sc = credentials["password"]

# SharePoint authentication
user = user_sc
password = password_sc

auth_url = "https://zebra.sharepoint.com/sites/PBI_GlobalMaterials"
folder_url = "/sites/PBI_GlobalMaterials/Shared%20Documents/Supplier%20visibility%20DNS/Files/"
#folder_url= "/sites/PBI_GlobalMaterials/Shared%20Documents/Supplier%20visibility%20DNS/Files/Previews%20week/WK%2029/"
gs_folder_path = '/Volumes/prod_catalog/shared_volume/gscr_ds/SV_DNS_reports_2024/2024/'

user_credentials = UserCredential(user, password)
ctx = ClientContext(auth_url).with_credentials(user_credentials)
web = ctx.web
ctx.load(web)
ctx.execute_query()

folder = ctx.web.get_folder_by_server_relative_url(folder_url)
files = folder.files
ctx.load(files)
ctx.execute_query()

# Regex pattern to match the file names
pattern = re.compile(r"Supplier Visibility Analytic DNS Report \d{2}-\d{2}-\d{4} WK\.(\d+)")

for file in files:
    file_name = file.properties['Name']
    match = pattern.match(file_name)
    if match:
        wk_number = match.group(1)
        wk_folder_path = os.path.join(gs_folder_path, f"WK {wk_number}")
        os.makedirs(wk_folder_path, exist_ok=True)
        
        file_url = folder_url + file_name
        source_file = ctx.web.get_file_by_server_relative_url(file_url)
        
        local_file_name = os.path.join(wk_folder_path, file_name)
        with open(local_file_name, "wb") as local_file:
            source_file.download(local_file).execute_query()
        print(f"[Ok] file has been downloaded to: {local_file_name}")

# COMMAND ----------

# from office365.runtime.auth.authentication_context import AuthenticationContext
# from office365.sharepoint.client_context import ClientContext
# from office365.runtime.auth.user_credential import UserCredential
# from office365.sharepoint.files.file import File 
# import os

# user = user_sc
# password = password_sc

# # CHANGE THIS PART  
# # --------------------------------------------------------------------------------------------------

# # all files/folders needs to be shared with GSCRSFSR@zebra.com
# auth_url = "https://zebra.sharepoint.com/sites/PBI_GlobalMaterials" # for authorization purpose
# folder_url = "/sites/PBI_GlobalMaterials/Shared%20Documents/Supplier%20visibility%20DNS/Files/Previews%20week/2024/WK%2027/" # this can be problematic (open in desktop and copy path from there if necessary)
# gs_folder_path = '/Volumes/dev_catalog/shared_volume/gscr_user_da/Sharepoint_TEST/' # to where do you want to load all files
# wk_folder_path = os.path.join(gs_folder_path, 'WK27') # Create WK28 directory inside the specified output path
# os.makedirs(wk_folder_path, exist_ok=True) # Ensure the directory exists
# # --------------------------------------------------------------------------------------------------

# user_credentials = UserCredential(user, password)
# ctx = ClientContext(auth_url).with_credentials(user_credentials)
# web = ctx.web
# ctx.load(web)
# ctx.execute_query()
# folder = ctx.web.get_folder_by_server_relative_url(folder_url)
# files = folder.files #Replace files with folders for getting list of folders
# ctx.load(files)
# ctx.execute_query()

# for file in files: 
#     file_url = folder_url+file.properties['Name']
#     print(file_url)
# local_file_name = ''

# for file in files:
#     file_url = folder_url+file.properties['Name']
#     source_file = ctx.web.get_file_by_server_relative_url(file_url)

#     local_file_name = os.path.join(wk_folder_path, file.properties['Name']) # Save to WK28 directory
#     with open(local_file_name, "wb") as local_file:
#         source_file.download(local_file).execute_query()
#     print("[Ok] file has been downloaded to: {0}".format(str(local_file_name)))

# COMMAND ----------

# from office365.runtime.auth.authentication_context import AuthenticationContext
# from office365.sharepoint.client_context import ClientContext
# from office365.runtime.auth.user_credential import UserCredential
# from office365.sharepoint.files.file import File 
# import os
# import re

# # Retrieve credentials from Google Secret Manager
# from google.cloud import secretmanager

# user_sc = credentials["username"]
# password_sc = credentials["password"]

# # SharePoint authentication
# user = user_sc
# password = password_sc

# auth_url = "https://zebra.sharepoint.com/sites/PBI_GlobalMaterials"
# folder_url = "/sites/PBI_GlobalMaterials/Shared%20Documents/Supplier%20visibility%20DNS/Files/"
# gs_folder_path = '/Volumes/dev_catalog/shared_volume/gscr_user_da/Sharepoint_TEST'

# user_credentials = UserCredential(user, password)
# ctx = ClientContext(auth_url).with_credentials(user_credentials)
# web = ctx.web
# ctx.load(web)
# ctx.execute_query()

# folder = ctx.web.get_folder_by_server_relative_url(folder_url)
# files = folder.files
# ctx.load(files)
# ctx.execute_query()

# # Regex pattern to match the file names
# pattern = re.compile(r"Supplier Visibility Analytic DNS Report \d{2}-\d{2}-\d{4} WK(\d+)")

# for file in files:
#     file_name = file.properties['Name']
#     match = pattern.match(file_name)
#     if match:
#         wk_number = match.group(1)
#         wk_folder_path = os.path.join(gs_folder_path, f"WK{wk_number}")
#         os.makedirs(wk_folder_path, exist_ok=True)
        
#         file_url = folder_url + file_name
#         source_file = ctx.web.get_file_by_server_relative_url(file_url)
        
#         local_file_name = os.path.join(wk_folder_path, file_name)
#         with open(local_file_name, "wb") as local_file:
#             source_file.download(local_file).execute_query()
#         print(f"[Ok] file has been downloaded to: {local_file_name}")

# COMMAND ----------

##https://zebra.sharepoint.com/sites/GOSDatascience/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FGOSDatascience%2FShared%20Documents%2FGeneral%2Fautomation%5Ftesting&viewid=06a030ef%2De4ae%2D4fd8%2D80da%2D3fc62564a704