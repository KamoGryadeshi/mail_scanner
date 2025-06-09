==============================
Mail XLSX Scanner
==============================

A simple program to automatically retrieve Excel attachments from unread emails.

----------------------------------------
What It Does
----------------------------------------

- Connects to an IMAP server
- Checks the inbox for unread messages
- Saves .xlsx attachments to the attachments folder
- (Optional) Loads data from the files into an SQLite database using pandas

----------------------------------------
Project Structure
----------------------------------------

mail-xlsx-scanner/
├── main.py               # Entry point (triggered on schedule)
├── mail_scanning.py      # Email fetching and attachment saving
├── to_db.py              # Import data from Excel into the database
├── .env                  # Environment variables
├── attachments/          # Folder for downloaded attachments
├── requirements.txt      # Dependencies list
└── README.txt            # This file

----------------------------------------
Installation & Usage
----------------------------------------

1. Install dependencies:

   pip install -r requirements.txt

2. Create a `.env` file with your credentials:

   USERNAME=your_email@example.com  
   MAIL_PASS=your_password  
   IMAP_SERVER=imap.yourmail.com  

3. Run manually (for testing):

   python main.py

----------------------------------------
Dependencies
----------------------------------------

- imaplib, email (standard libraries)
- pandas, openpyxl
- python-dotenv