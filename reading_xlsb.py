import io
import pandas as pd
from pyxlsb import open_workbook
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# SharePoint details
SHAREPOINT_SITE_URL = "https://airtelworld.sharepoint.com/MobilityServices"
CLIENT_ID = "B0321755"
CLIENT_SECRET = "Bihar@1949"
FILE_RELATIVE_URL = "/Shared Documents/Documents/NPO/Ground Hog - Combined Hygiene Reports/FEB'25/AP_FEB'25.xlsb"  # Adjust this as per your folder structure
# Authenticate SharePoint
ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(ClientCredential(CLIENT_ID,CLIENT_SECRET))

# Get the file
file = ctx.web.get_file_by_server_relative_url(FILE_RELATIVE_URL)
response = file.download(ctx)
ctx.execute_query()

# Read file into memory
xlsb_data = io.BytesIO(response.content)

# Read specific sheet
df = pd.read_excel(xlsb_data, engine='pyxlsb', sheet_name="Sheet1")  # Change sheet_name as needed

# Display DataFrame
print(df.head())