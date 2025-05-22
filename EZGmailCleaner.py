import ezgmail
import csv
import sys
import os
import pandas as pd
import traceback

CSV_FILENAME = 'mail_headers.csv'
ACTIONS = ['Delete', 'Export', 'Move to folder', 'AddStar']

def export_headers():
    ezgmail.init()
    threads = ezgmail.search('in:inbox')
    all_headers = set()
    rows = []
    for thread in threads:
        for msg in thread.messages:
            headers = {
                'Message-ID': msg.id,
                'Subject': msg.subject,
                'From': msg.sender,
                'To': msg.recipient,
                'Date': msg.date,
                'Snippet': msg.snippet
            }
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
    df = pd.read_csv(CSV_FILENAME)
    for idx, row in df.iterrows():
        print(f"Message-ID: {row.get('Message-ID', 'N/A')} | Action: {row['actions']}")

def run_actions():
    ezgmail.init()
    df = pd.read_csv(CSV_FILENAME)
    summary = {action: 0 for action in ACTIONS}
    summary['Unknown'] = 0
    errors = 0
    try:
        for idx, row in df.iterrows():
            action = str(row['actions']).strip()
            msg_id = row.get('Message-ID', '').strip()
            if not action or not msg_id:
                continue
            msg = None
            # Find the message by ID
            for thread in ezgmail.search(f'rfc822msgid:{msg_id}'):
                for m in thread.messages:
                    if m.id == msg_id:
                        msg = m
                        break
                if msg:
                    break
            if not msg:
                print(f"ERROR: Message {msg_id} not found.")
                errors += 1
                continue
            if action == 'Delete':
                try:
                    msg.trash()
                    summary['Delete'] += 1
                    print(f"Deleted message {msg_id}")
                except Exception as e:
                    print(f"ERROR deleting message {msg_id}: {e}")
                    errors += 1
            elif action == 'Export':
                try:
                    export_path = f"export_{msg_id}.txt"
                    with open(export_path, 'w', encoding='utf-8') as f:
                        f.write(f"Subject: {msg.subject}\nFrom: {msg.sender}\nTo: {msg.recipient}\nDate: {msg.date}\n\n{msg.body}")
                    summary['Export'] += 1
                    print(f"Exported message {msg_id} to {export_path}")
                except Exception as e:
                    print(f"ERROR exporting message {msg_id}: {e}")
                    errors += 1
            elif action == 'Move to folder':
                try:
                    # Example: Move to "Work" label (must exist)
                    msg.addLabel('Work')
                    msg.removeLabel('INBOX')
                    summary['Move to folder'] += 1
                    print(f"Moved message {msg_id} to 'Work'")
                except Exception as e:
                    print(f"ERROR moving message {msg_id}: {e}")
                    errors += 1
            elif action == 'AddStar':
                try:
                    msg.star()
                    summary['AddStar'] += 1
                    print(f"Starred message {msg_id}")
                except Exception as e:
                    print(f"ERROR starring message {msg_id}: {e}")
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
        print("Usage: python main_ezgmail.py [EXPORT|TEST|RUN]")
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