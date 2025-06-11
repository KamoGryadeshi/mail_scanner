import imaplib
import email
import os
from email.header import decode_header
from dotenv import load_dotenv

load_dotenv()
username = os.getenv("MAILUSERNAME")
mail_pass = os.getenv("MAIL_PASS")
imap_server = os.getenv("IMAP_SERVER")

files_dir = "./files"

def decode_filename(raw_name):
    if not raw_name:
        return None
    parts = decode_header(raw_name)
    return ''.join([
        part.decode(encoding or 'utf-8') if isinstance(part, bytes) else part
        for part, encoding in parts
    ])

def get_email():
    try:
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(username, mail_pass)
        imap.select("INBOX")
        status, data = imap.search(None, 'UNSEEN')

        if status != 'OK' or not data[0]:
            print('Новых писем нет')
            return

        message_ids = data[0].decode().split()
        for msg_id in message_ids:
            res, msg = imap.fetch(msg_id, '(RFC822)')
            msg = email.message_from_bytes(msg[0][1])
            found = False
            # letter_date = email.utils.parsedate_tz(msg["Date"])
            # letter_from = msg["Return-path"]
            # print(letter_date, letter_from)
            for part in msg.walk():
                if (part.get_content_type() ==
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'):
                    raw_filename = part.get_filename()
                    filename = decode_filename(raw_filename)
                    if not filename:
                        continue

                    payload = part.get_payload(decode=True)
                    filepath = os.path.join(files_dir, filename)
                    with open(filepath, "wb") as f:
                        f.write(payload)

                    print(f"Найден файл: {filename} ({len(payload)} байт)")
                    found = True
            if not found:
                print('В письме вложений не обнаружено')

    except Exception as e:
        print("Ошибка:", e)

    finally:
        imap.logout()

if __name__ == '__main__':
    get_email()