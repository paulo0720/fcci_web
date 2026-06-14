import sqlite3

conn = sqlite3.connect("database/fcci.db")
cursor = conn.cursor()

try:
    cursor.execute("""
    ALTER TABLE members
    ADD COLUMN photo_path TEXT
    """)
except:
    pass

conn.commit()
conn.close()

print("DONE")