from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.auth.ms_identity.authentication_context import DeviceCodeCredential
import os

site_url = "https://airtelworld.sharepoint.com/sites/MobilityServices"
relative_folder_url = "/sites/MobilityServices/Shared Documents/NPO/Ground Hog - Combined Hygiene Reports/FEB'25"
download_path = r"D:\Ground_Hog_Hyigine_Reports"

# Use DeviceCodeCredential for modern auth
credentials = DeviceCodeCredential("your-tenant-id", "your-client-id")
ctx = ClientContext(site_url).with_credentials(credentials)

# Access folder and download files (same as before)
folder = ctx.web.get_folder_by_server_relative_url(relative_folder_url)
files = folder.files
ctx.load(files)
ctx.execute_query()

os.makedirs(download_path, exist_ok=True)

for file in files:
    file_name = file.properties["Name"]
    if file_name.lower().endswith(".xlsb"):
        print(f"⬇️ Downloading: {file_name}")
        file_obj = ctx.web.get_file_by_server_relative_url(f"{relative_folder_url}/{file_name}").download()
        ctx.execute_query()
        with open(os.path.join(download_path, file_name), "wb") as f:
            f.write(file_obj.content)
