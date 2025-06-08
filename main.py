from mail_to_db import import_xlsx_to_db
import schedule
import time
from mail_scanning import get_email

def main_def():
    get_email()
    import_xlsx_to_db()

schedule.every().day.at("12:00").do(main_def)

while True:
    schedule.run_pending()
    time.sleep(60)

