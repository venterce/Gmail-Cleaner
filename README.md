# EZGmailCleaner

**EZGmailCleaner** is a Python tool for managing your Gmail INBOX using [ezgmail](https://pypi.org/project/ezgmail/) by Al Sweigart.  
It exports message headers to a CSV file, lets you plan actions in Excel, and then processes those actions (delete, export, move, star) back in your mailbox.

---

## Features

- **EXPORT:** Exports Gmail INBOX message headers to a CSV file (`mail_headers.csv`). The CSV includes a blank "actions" column for user planning.
- **TEST:** Reads the CSV and shows the planned actions for each message. No changes are made to Gmail.
- **RUN:** Reads the CSV and performs the specified actions on your Gmail mailbox:
  - `Delete`: Moves the message to Trash.
  - `Export`: Saves the message's headers and text to a file.
  - `Move to folder`: Moves the message to a specific label (edit code for your label).
  - `AddStar`: Adds a star to the message.
- **Error Handling:** Stops execution and closes files on error. Unknown actions are reported and skipped.
- **Summary:** After RUN, prints a summary of actions taken.

---

## Setup

1. **Install dependencies:**
   ```sh
   pip install ezgmail pandas
   ```

2. **First-time authentication:**  
   You must run `ezgmail.init()` once in your project folder to set up Gmail API credentials.  
   This will open a browser window for you to log in and authorize access.

   ```python
   import ezgmail
   ezgmail.init()
   ```

   This creates `token.json` and `credentials.json` in your folder.

---

## Usage

```sh
python EZGmailCleaner.py EXPORT   # Export headers to CSV
python EZGmailCleaner.py TEST     # Show actions to be performed (no changes made)
python EZGmailCleaner.py RUN      # Apply actions to Gmail
```

---

## CSV Structure Example

| actions      | Message-ID | Subject | From | To | Date | Snippet | ... |
|--------------|------------|---------|------|----|------|---------|-----|
| Delete       | ...        | ...     | ...  | ...| ...  | ...     | ... |
| Export       | ...        | ...     | ...  | ...| ...  | ...     | ... |
| Move to folder | ...      | ...     | ...  | ...| ...  | ...     | ... |
| AddStar      | ...        | ...     | ...  | ...| ...  | ...     | ... |

- Fill in the `actions` column in Excel or another editor before running `RUN`.

---

## Notes

- For "Move to folder", edit the label name in the code (`msg.addLabel('Work')`) to match your Gmail label.
- For "Export", exported files are named `export_<Message-ID>.txt` by default.
- If an action is not recognized, it will be reported and skipped.
- If any error occurs, the program will halt and print the error.

---

## Credits

- [ezgmail](https://pypi.org/project/ezgmail/) by Al Sweigart