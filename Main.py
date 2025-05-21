import os
import sys
import csv
import pandas as pd
import traceback
import base64

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
CSV_FILENAME = 'mail_headers.csv'

ACTIONS = [
    'Delete',
    'Star',
    'Move to folder'
    # Export and Export and Star are handled via prefix detection
]

def authenticate_gmail():
    """Authenticate with Gmail API and return the service object."""
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

def export_headers():
    """Export Gmail INBOX message headers to a CSV file."""
    service = authenticate_gmail()
    results = service.users().messages().list(userId='me', labelIds=['INBOX'], maxResults=1000).execute()
    messages = results.get('messages', [])
    all_headers = set()
    rows = []
    print(f"Found {len(messages)} messages in INBOX.")
    for msg in messages:
        msg_data = service.users().messages().get(userId='me', id=msg['id'], format='metadata').execute()
        headers = {h['name']: h['value'] for h in msg_data['payload']['headers']}
        headers['Message-ID'] = msg['id']
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
        if 'parts' in payload:
            for part in payload['parts']:
                if part.get('mimeType') == 'text/plain' and 'data' in part.get('body', {}):
                    return base64.urlsafe_b64decode(part['body']['data']).decode('utf-8', errors='replace')
                elif 'parts' in part:
                    result = extract_parts(part)
                    if result:
                        return result
        elif payload.get('mimeType') == 'text/plain' and 'data' in payload.get('body', {}):
            return base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8', errors='replace')
        return ""
    body = extract_parts(msg['payload'])
    if not body and 'body' in msg['payload'] and 'data' in msg['payload']['body']:
        body = base64.urlsafe_b64decode(msg['payload']['body']['data']).decode('utf-8', errors='replace')
    return body

def run_actions():
    """Read the CSV and execute actions on Gmail messages."""
    service = authenticate_gmail()
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

            # Export and Star
            if action.lower().startswith('export and star:'):
                _, filepath = action.split(':', 1)
                filepath = filepath.strip()
                try:
                    msg = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
                    headers = {h['name']: h['value'] for h in msg['payload']['headers']}
                    body = get_message_body(msg)
                    os.makedirs(os.path.dirname(filepath), exist_ok=True)
                    with open(filepath, 'w', encoding='utf-8') as f:
                        for h, v in headers.items():
                            f.write(f"{h}: {v}\n")
                        f.write("\n--- MESSAGE BODY ---\n")
                        f.write(body)
                    # Star the message
                    service.users().messages().modify(userId='me', id=msg_id, body={'addLabelIds': ['STARRED']}).execute()
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
                    msg = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
                    headers = {h['name']: h['value'] for h in msg['payload']['headers']}
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
                    service.users().messages().delete(userId='me', id=msg_id).execute()
                    summary['Delete'] += 1
                    print(f"Deleted message {msg_id}")
                except Exception as e:
                    print(f"ERROR deleting message {msg_id}: {e}")
                    errors += 1
            elif action == 'Star':
                try:
                    service.users().messages().modify(userId='me', id=msg_id, body={'addLabelIds': ['STARRED']}).execute()
                    summary['Star'] += 1
                    print(f"Starred message {msg_id}")
                except Exception as e:
                    print(f"ERROR starring message {msg_id}: {e}")
                    errors += 1
            elif action == 'Move to folder':
                try:
                    # Example: Move to 'Work' - replace 'Label_123456' with your label ID
                    service.users().messages().modify(userId='me', id=msg_id, body={'addLabelIds': ['Label_123456'], 'removeLabelIds': ['INBOX']}).execute()
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
