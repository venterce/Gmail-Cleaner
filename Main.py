import os
import sys
import csv
import pandas as pd
import traceback
import base64

from google.cloud import gmail_v1
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
CSV_FILENAME = 'mail_headers.csv'

ACTIONS = [
    'Delete',
    'Star',
    'Move to folder'
    # Export and Export and Star are handled via prefix detection
]

def authenticate_gmail():
    """Authenticate with Gmail API using the new Cloud Client library and return the client."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    client = gmail_v1.GmailServiceClient(credentials=creds)
    return client

def export_headers():
    """Export Gmail INBOX message headers to a CSV file using the new client."""
    client = authenticate_gmail()
    user = "users/me"
    all_headers = set()
    rows = []
    messages = list(client.list_messages(parent=user, label_ids=["INBOX"], max_results=1000))
    print(f"Found {len(messages)} messages in INBOX.")
    for msg in messages:
        msg_data = client.get_message(name=msg.name, format_="METADATA")
        headers = {h.name: h.value for h in msg_data.payload.headers}
        headers['Message-ID'] = msg_data.id
        rows.append(headers)
        all_headers.update(headers.keys())
    all_headers = sorted(list(all_headers))
    with open(CSV_FILENAME, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=['actions'] + all_headers)
        writer.writeheader()
        for row in rows:
            row_out = {h: row.get(h, '') for h in all_headers}
            row_out['actions'] = ''
            writer.writerow(row_out)
    print(f"Exported {len(rows)} messages to {CSV_FILENAME}")

def test_actions():
    """Read the CSV and display actions for each message."""
    df = pd.read_csv(CSV_FILENAME)
    for idx, row in df.iterrows():
        print(f"Message-ID: {row.get('Message-ID', 'N/A')} | Action: {row['actions']}")

def get_message_body(msg):
    """Extract and return the plain text body from a Gmail message."""
    def extract_parts(payload):
        if hasattr(payload, 'parts') and payload.parts:
            for part in payload.parts:
                if getattr(part, 'mime_type', '') == 'text/plain' and getattr(part.body, 'data', None):
                    return base64.urlsafe_b64decode(part.body.data).decode('utf-8', errors='replace')
                elif hasattr(part, 'parts'):
                    result = extract_parts(part)
                    if result:
                        return result
        elif getattr(payload, 'mime_type', '') == 'text/plain' and getattr(payload.body, 'data', None):
            return base64.urlsafe_b64decode(payload.body.data).decode('utf-8', errors='replace')
        return ""
    body = extract_parts(msg.payload)
    if not body and hasattr(msg.payload, 'body') and getattr(msg.payload.body, 'data', None):
        body = base64.urlsafe_b64decode(msg.payload.body.data).decode('utf-8', errors='replace')
    return body

def run_actions():
    """Read the CSV and execute actions on Gmail messages."""
    client = authenticate_gmail()
    user = "users/me"
    df = pd.read_csv(CSV_FILENAME)
    summary = {
        'Delete': 0,
        'Star': 0,
        'Move to folder': 0,
        'Export': 0,
        'Export and Star': 0,
        'Unknown': 0
    }
    errors = 0

    try:
        for idx, row in df.iterrows():
            action = str(row['actions']).strip()
            msg_id = row.get('Message-ID', '').strip()
            if not action or not msg_id:
                continue

            msg_name = f"{user}/messages/{msg_id}"

            # Export and Star
            if action.lower().startswith('export and star:'):
                _, filepath = action.split(':', 1)
                filepath = filepath.strip()
                try:
                    msg = client.get_message(name=msg_name, format_="FULL")
                    headers = {h.name: h.value for h in msg.payload.headers}
                    body = get_message_body(msg)
                    os.makedirs(os.path.dirname(filepath), exist_ok=True)
                    with open(filepath, 'w', encoding='utf-8') as f:
                        for h, v in headers.items():
                            f.write(f"{h}: {v}\n")
                        f.write("\n--- MESSAGE BODY ---\n")
                        f.write(body)
                    # Star the message
                    client.modify_message(
                        name=msg_name,
                        add_label_ids=['STARRED']
                    )
                    summary['Export and Star'] += 1
                    print(f"Exported and starred message {msg_id} to '{filepath}'")
                except Exception as e:
                    print(f"ERROR exporting and starring message {msg_id} to '{filepath}': {e}")
                    errors += 1
                    continue

            # Export only
            elif action.lower().startswith('export:'):
                _, filepath = action.split(':', 1)
                filepath = filepath.strip()
                try:
                    msg = client.get_message(name=msg_name, format_="FULL")
                    headers = {h.name: h.value for h in msg.payload.headers}
                    body = get_message_body(msg)
                    os.makedirs(os.path.dirname(filepath), exist_ok=True)
                    with open(filepath, 'w', encoding='utf-8') as f:
                        for h, v in headers.items():
                            f.write(f"{h}: {v}\n")
                        f.write("\n--- MESSAGE BODY ---\n")
                        f.write(body)
                    summary['Export'] += 1
                    print(f"Exported message {msg_id} to '{filepath}'")
                except Exception as e:
                    print(f"ERROR exporting message {msg_id} to '{filepath}': {e}")
                    errors += 1
                    continue

            elif action == 'Delete':
                try:
                    client.delete_message(name=msg_name)
                    summary['Delete'] += 1
                    print(f"Deleted message {msg_id}")
                except Exception as e:
                    print(f"ERROR deleting message {msg_id}: {e}")
                    errors += 1
            elif action == 'Star':
                try:
                    client.modify_message(
                        name=msg_name,
                        add_label_ids=['STARRED']
                    )
                    summary['Star'] += 1
                    print(f"Starred message {msg_id}")
                except Exception as e:
                    print(f"ERROR starring message {msg_id}: {e}")
                    errors += 1
            elif action == 'Move to folder':
                try:
                    # Example: Move to 'Work' - replace 'Label_123456' with your label ID
                    client.modify_message(
                        name=msg_name,
                        add_label_ids=['Label_123456'],
                        remove_label_ids=['INBOX']
                    )
                    summary['Move to folder'] += 1
                    print(f"Moved message {msg_id} to folder/label")
                except Exception as e:
                    print(f"ERROR moving message {msg_id}: {e}")
                    errors += 1
            else:
                print(f"ERROR: Action '{action}' for message {msg_id} not recognized.")
                summary['Unknown'] += 1
                errors += 1

    except Exception as e:
        print('FATAL ERROR:', e)
        traceback.print_exc()
        sys.exit(1)

    print("\nSummary of actions:")
    for k, v in summary.items():
        print(f"{k}: {v}")
    if errors:
        print(f"Errors encountered: {errors}")

if __name__ == "__main__":
    if len(sys.argv) == 1:
        print("Usage: python main.py [EXPORT|TEST|RUN]")
        sys.exit(1)
    cmd = sys.argv[1].upper()
    try:
        if cmd == 'EXPORT':
            export_headers()
        elif cmd == 'TEST':
            test_actions()
        elif cmd == 'RUN':
            run_actions()
        else:
            print(f"Unknown command: {cmd}")
    except Exception as e:
        print("Fatal error:", e)
        sys.exit(1)