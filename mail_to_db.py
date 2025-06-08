import pandas as pd
import sqlite3
import os
from datetime import datetime

DB_FILE = "mail_data.db"
FILES_DIR = "./files"

def import_xlsx_to_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS log_imports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            import_time TEXT
        )
    ''')

    for filename in os.listdir(FILES_DIR):
        if filename.endswith('.xlsx'):
            filepath = os.path.join(FILES_DIR, filename)

            cursor.execute("SELECT 1 FROM log_imports WHERE filename = ?", (filename,))
            if cursor.fetchone():
                print("Файл уже импортирован.")
                continue

            try:
                df = pd.read_excel(filepath)
                table_name = os.path.splitext(filename)[0].lower().replace(" ", "_")
                df.to_sql(table_name, conn, if_exists='replace', index=False)

                cursor.execute("INSERT INTO log_imports (filename, import_time) VALUES (?, ?)", (
                    filename,
                    datetime.now().isoformat()
                ))

                print(f"[Успешно] Импортирован {filename} → таблица '{table_name}'")
                show_table(table_name)

            except Exception as e:
                print(f"[Ошибка] Не удалось импортировать {filename}: {e}")

    conn.commit()
    conn.close()


def show_table(table_name):
    conn = sqlite3.connect(DB_FILE)

    try:
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        print(df)
    except Exception as e:
        print(f"Не удалось прочитать таблицу '{table_name}': {e}")
    finally:
        conn.close()