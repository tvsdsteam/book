import os
import time
import datetime
import urllib.parse
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'eastern-surface-456208-r2-b0fcbc3d5f53.json'

SPREADSHEET_ID = '1Y_D1zGRUgaK0B87hSUBb7T1D6UDJXcXFVr3gndAxCiU'
SOURCE_SHEET_NAME = 'list'
ARCHIVE_SHEET_NAME = 'Archived'

TEXTMEBOT_API_KEY = 'T55wPwawmj2V'

def get_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


def get_sheet_id(service, sheet_name):
    metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in metadata.get('sheets', []):
        if sheet.get("properties", {}).get("title") == sheet_name:
            return sheet.get("properties", {}).get("sheetId")
    return None

# ---------------------------------------------
# SEND WHATSAPP MESSAGE USING TEXTMEBOT
# ---------------------------------------------
def send_whatsapp_message(to_phone, message_text):
    try:
        print(f"📲 Sending WhatsApp message to {to_phone}...")

        full_number = "+91" + to_phone.strip()
        encoded_message = urllib.parse.quote(message_text)

        api_url = (
            f"https://api.textmebot.com/send.php"
            f"?recipient={full_number}"
            f"&apikey={TEXTMEBOT_API_KEY}"
            f"&text={encoded_message}"
        )

        response = requests.get(api_url)

        if response.status_code == 200:
            if "SENT" in response.text.upper():
                print(f"✅ Message successfully sent to {to_phone}")
            else:
                print(f"⚠️ API responded: {response.text}")
        else:
            print(f"❌ Failed to send message to {to_phone}: HTTP {response.status_code} - {response.text}")

    except Exception as e:
        print(f"❌ Error sending message to {to_phone}: {e}")

# ---------------------------------------------
# SHEET MANIPULATION
# ---------------------------------------------
def append_row_to_archive(service, row):
    body = {'values': [row]}
    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{ARCHIVE_SHEET_NAME}!A1",
        valueInputOption='RAW',
        body=body
    ).execute()
    print("✅ Row archived.")

def delete_row_from_source(service, sheet_id, row_index):
    body = {
        "requests": [{
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": row_index,
                    "endIndex": row_index + 1
                }
            }
        }]
    }
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=body
    ).execute()
    print(f"🗑️ Row {row_index+1} deleted.")

# ---------------------------------------------
# PROCESS EACH ROW
# ---------------------------------------------
def process_sheet(service, source_sheet_id):
    RANGE_NAME = f"{SOURCE_SHEET_NAME}!A2:F"
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=RANGE_NAME
    ).execute()
    rows = result.get('values', [])

    if not rows:
        print("📋 No rows found.")
        return

    now = datetime.datetime.now()
    current_str = now.strftime("%Y-%m-%d %H:%M")
    print("⏰ Current time:", current_str)

    for i in range(len(rows) - 1, -1, -1):
        try:
            row = rows[i]
            if len(row) < 5:
                print(f"⚠️ Row {i+2} missing fields. Skipping.")
                continue

            scheduled_date = row[4].strip()
            scheduled_time = row[5].strip()
            to_phone = row[1].strip()
            message_text = "\n Just a Reminder : " + row[2] + "\n \n" +row[3].strip()

            scheduled_datetime = datetime.datetime.strptime(
                f"{scheduled_date} {scheduled_time}", "%m/%d/%Y %I:%M:%S %p"
            )
            scheduled_str = scheduled_datetime.strftime("%Y-%m-%d %H:%M")

            if current_str == scheduled_str:
                print(f"✅ Time match at row {i+2}: sending to {to_phone}")
                send_whatsapp_message(to_phone, message_text)
                append_row_to_archive(service, row)
                delete_row_from_source(service, source_sheet_id, i + 1)
            else:
                print(f"⏳ Row {i+2} not scheduled (Scheduled: {scheduled_str})")
        except Exception as e:
            print(f"❌ Error in row {i+2}: {e}")

# ---------------------------------------------
# MAIN LOOP
# ---------------------------------------------
def main():
    service = get_service()
    source_sheet_id = get_sheet_id(service, SOURCE_SHEET_NAME)
    if source_sheet_id is None:
        print("❌ Sheet not found.")
        return

    print("🚀 Script started. Running every minute...")

    while True:
        now = datetime.datetime.now()
        seconds_until_next_minute = 60 - now.second
        print(f"🕒 Sleeping {seconds_until_next_minute} seconds...")
        time.sleep(seconds_until_next_minute)
        print("\n🔄 Checking sheet at", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        process_sheet(service, source_sheet_id)

# ---------------------------------------------
if __name__ == '__main__':
    main()
