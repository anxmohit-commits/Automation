import time
import os
import requests
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from datetime import datetime

# === CONFIGURATION ===
FOLDER_TO_WATCH = r'D:\Daily Cell Outage Report\Output'
TEAMS_WEBHOOK_URL = 'https://teams.microsoft.com/l/chat/19:182a8a38-fd88-435b-a4ae-cf3181078f54_31130419-15e3-4967-be8a-b85c06bc7357@unq.gbl.spaces/conversations?context=%7B%22contextType%22%3A%22chat%22%7DRE'  # Step 1 se copy karo
EXTENSIONS = ['.csv', '.xlsx', '.xls']
LAST_FILE = None

class FolderMonitor(FileSystemEventHandler):
    def on_modified(self, event):
        if not event.is_directory:
            self.check_file()
    
    def on_created(self, event):
        if not event.is_directory:
            self.check_file()
    
    def check_file(self):
        check_and_notify()

def get_latest_file():
    """Latest modified file find karo"""
    try:
        files = []
        for f in os.listdir(FOLDER_TO_WATCH):
            path = os.path.join(FOLDER_TO_WATCH, f)
            if os.path.isfile(path):
                ext = os.path.splitext(f)[1].lower()
                if ext in EXTENSIONS:
                    files.append(path)
        
        if files:
            return max(files, key=os.path.getmtime)
    except:
        pass
    return None

def send_to_teams(filepath):
    """File info Teams ko bhejo"""
    global LAST_FILE
    
    if LAST_FILE == filepath:
        return
    
    try:
        filename = os.path.basename(filepath)
        file_size = os.path.getsize(filepath)
        mod_time = datetime.fromtimestamp(os.path.getmtime(filepath)).strftime("%d-%m-%Y %H:%M:%S")
        
        # Teams message
        message = {
            "@type": "MessageCard",
            "@context": "https://schema.org/extensions",
            "summary": f"New Report: {filename}",
            "themeColor": "0078D4",
            "sections": [
                {
                    "activityTitle": f"📊 New Report Available",
                    "activitySubtitle": filename,
                    "facts": [
                        {"name": "File Name:", "value": filename},
                        {"name": "Size:", "value": f"{file_size:,} bytes"},
                        {"name": "Modified:", "value": mod_time},
                        {"name": "Location:", "value": filepath}
                    ],
                    "markdown": True
                },
                {
                    "activityTitle": "📁 Open File",
                    "text": f"```\n{filepath}\n```"
                }
            ]
        }
        
        response = requests.post(
            TEAMS_WEBHOOK_URL,
            json=message,
            timeout=10,
            headers={"Content-Type": "application/json"}
        )
        
        if response.status_code == 200:
            print(f"✅ {datetime.now().strftime('%H:%M:%S')} - Sent to Teams: {filename}")
            LAST_FILE = filepath
        else:
            print(f"⚠️ Teams error: {response.status_code}")
    
    except Exception as e:
        print(f"❌ Error: {e}")

def check_and_notify():
    """Check latest file aur bhejo"""
    latest = get_latest_file()
    if latest:
        send_to_teams(latest)

if __name__ == "__main__":
    if not os.path.exists(FOLDER_TO_WATCH):
        print(f"❌ Folder not found: {FOLDER_TO_WATCH}")
        exit(1)
    
    print(f"📁 Folder: {FOLDER_TO_WATCH}")
    print(f"📄 Types: {', '.join(EXTENSIONS)}")
    print(f"✅ Connected to Teams")
    print(f"⏹️ Press Ctrl+C to stop\n")
    
    # Initial check
    check_and_notify()
    
    # Monitor
    monitor = FolderMonitor()
    observer = Observer()
    observer.schedule(monitor, FOLDER_TO_WATCH, recursive=False)
    observer.start()
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n🛑 Stopped")
        observer.stop()
    
    observer.join()