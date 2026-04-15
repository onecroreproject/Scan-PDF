import sqlite3
import os

db_path = 'db.sqlite3'
if not os.path.exists(db_path):
    print("Database not found")
else:
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cols = [
        ("qr_data", "JSON"),
        ("qr_type", "VARCHAR(30) DEFAULT 'url'"),
        ("logo", "VARCHAR(100) NULL")
    ]
    
    for col_name, col_type in cols:
        try:
            cursor.execute(f"ALTER TABLE dynamic_qr_dynamicqrcode ADD COLUMN {col_name} {col_type};")
            print(f"Added {col_name}")
        except sqlite3.OperationalError:
            print(f"{col_name} already exists")
            
    conn.commit()
    conn.close()
    print("Done")
