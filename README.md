# Gmail-Cleaner
A python program that first exports Gmail message headers to a CSV file, adding an Action Column. Add actions to the action column, then pass the file back to the program. Actions will be taken for each message.

Intial files were all created by Github's Copilot based on the following prompt:

Create a python program that will read my gmail is mailbox and export a list of all the headers. Write a list of all the headers to a CSV file (with headers) including the properties of each message and a new first column called actions.  The files will be loaded into excel and the actions column will be filled in with actions including Delete, Export, Move to folder, Export and Star.  Then the program will be executed with a command line argument of TEST  which will read the csv file and show what actions are listed for each message.  The program will also have another command option Run which will execute the action against the gmail mailbox. If the action found is not listed, write an appropriate error message. If there is an error reading, writing the data or running the actions against the mailbox, halt with an error but close all open files. Show a summary of actions taken.

![image](https://github.com/user-attachments/assets/5a844451-f16c-4c58-bf11-70b1cba2fdd9)

# mailcleaner*
*mailcleaner** 
is a Python tool for managing your Gmail INBOX by exporting message headers, planning actions in a CSV file, and then automatically processing those actions (delete, star, move, export, etc.) back in your mailbox.
---## Features
-**EXPORT:** Exports Gmail INBOX message headers to a CSV file (`mail_headers.csv`). The CSV includes a blank "actions" column for user planning.
-**TEST:** Reads the CSV and shows the planned actions for each message. No changes are made to Gmail.
- **RUN:** Reads the CSV and performs the specified actions on your Gmail mailbox:
- `Delete`: Permanently deletes the message.
- `Star`: Adds the “Starred” label.
- `Move to folder`: Moves the message to a specific label (edit code for your label).
- `Export:<filepath>`: Saves the message's headers and text to the given file path.
- `Export and Star:<filepath>`: Saves the message's headers and text and adds the “Starred” label.
- **Error Handling:** Stops execution and closes files on error. Unknown actions are reported and skipped.-
-
**Summary:** After RUN, prints a summary of actions taken.-
-
-## CSV Structure Example
| actions                   | Message-ID | Subject     | From       | To          | Date               | ... 
|---------------------------|------------|-------------|------------|-------------|--------------------|-----|
| Delete                    | 123abc...  | Hello       | a@b.com    | me@gmail.com| 2025-05-21 10:00   | ... |
| Export:exports/mail1.txt  | 456def...  | Notice      | c@d.com    | me@gmail.com| 2025-05-20 15:00   | ... |
| Export and Star:star2.txt | 789ghi...  | Promotion   | e@f.com    | me@gmail.com| 2025-05-19 08:00   | ... |
| Move to folder            | ...        | ...         | ...        | ...         | ...                | ... |

**Note:** For `Export` and `Export and Star`, the path after the colon specifies where the message details will be saved.
-
-
## Usage```bashpython main.py EXPORT   # Export headers to CSV
python main.py TEST     # Show actions to be performed (no changes made)
python main.py RUN      # Apply actions to Gmail

![image](https://github.com/user-attachments/assets/feced8f8-51ad-44cd-a1be-d35ec31ce0c7)
