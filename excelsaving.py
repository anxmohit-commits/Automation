import win32com.client
import os
import re
import requests
from datetime import datetime, timedelta
from getpass import getpass
from urllib.parse import unquote, urlparse, parse_qs

# **1️⃣ Ask for Credentials**
USERNAME = input("Enter your username: ").strip()
PASSWORD = getpass("Enter your password: ").strip()  

# **2️⃣ Define Download Path**
save_path = r"D:\Automation\WO Dump\Upload\WO Dump"
os.makedirs(save_path, exist_ok=True)

# **3️⃣ Connect to Outlook**
try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Inbox
    target_folder = inbox.Folders("Required to Work")  # Subfolder
except Exception as e:
    print(f"❌ Outlook connection error: {e}")
    exit()

# **4️⃣ Get Emails (Last 7 Days, Sorted by Newest)**
messages = target_folder.Items
messages.Sort("[ReceivedTime]", True)
date_filter = datetime.now() - timedelta(days=7)
messages = messages.Restrict(f"[ReceivedTime] >= '{date_filter.strftime('%m/%d/%Y %H:%M %p')}'")

# **5️⃣ Subject Keywords to Match**
subject_keywords = {
    "MobilityActiveClosedTT15-30DayDump V2": "ClosedTT15-30",
    "MobilityActiveClosedTT15DayDump V2": "ClosedTT15",
    "MobilityActiveOpenTT30DayDump V2": "OpenTT30"
}

# **6️⃣ Find the Latest Email for Each Subject Type**
latest_emails = {}  

for message in messages:
    try:
        subject = message.Subject.strip()
        
        for subject_keyword, filename_part in subject_keywords.items():
            if subject_keyword in subject:
                if filename_part not in latest_emails:
                    latest_emails[filename_part] = message
                    print(f"✅ Found latest email for: {filename_part} → {subject}")
                break  

    except Exception as e:
        print(f"⚠️ Error processing email: {e}")

if not latest_emails:
    print("❌ No matching emails found in the last 7 days.")
    exit()

# **7️⃣ Extract Download Links & Download Files**
def extract_real_url(safelink):
    """Extract real URL from Outlook SafeLink"""
    if "safelinks.protection.outlook.com" in safelink:
        parsed_url = urlparse(safelink)
        query_params = parse_qs(parsed_url.query)
        if "url" in query_params:
            return unquote(query_params["url"][0])
    return safelink

for filename_part, latest_email in latest_emails.items():
    email_body = latest_email.Body
    match = re.search(r"Link\s*[:\-]?\s*<?(https?://[^\s<>\"']+)>?", email_body)

    if not match:
        print(f"⚠️ No valid download link found in email: {filename_part}")
        continue

    download_url = extract_real_url(match.group(1))  # Extract real URL from SafeLink if necessary
    print(f"🔗 Downloading from: {download_url}")

    # **🔄 Retry logic for server errors (500)**
    retry_count = 3
    for attempt in range(retry_count):
        try:
            response = requests.get(download_url, auth=(USERNAME, PASSWORD), allow_redirects=True)

            if response.status_code == 200:
                today_date = datetime.now().strftime('%Y%m%d')
                file_name = f"Mobility_{filename_part}_{today_date}.csv"
                file_path = os.path.join(save_path, file_name)

                with open(file_path, "wb") as file:
                    file.write(response.content)

                print(f"📂 File downloaded successfully: {file_path}")
                break  # Exit retry loop after success

            elif response.status_code == 401:
                print(f"❌ Unauthorized access (401) for {filename_part}. Check credentials or token expiry.")
                break  # No point retrying

            elif response.status_code == 500:
                print(f"⚠️ Server error (500) for {filename_part}, retrying... ({attempt+1}/{retry_count})")

            else:
                print(f"❌ Failed to download {filename_part}, Status Code: {response.status_code}")
                break  # Stop retries for other unexpected errors

        except Exception as e:
            print(f"❌ Error downloading {filename_part}: {e}")

    else:
        print(f"❌ Failed to download {filename_part} after {retry_count} retries.")
