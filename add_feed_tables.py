import sqlite3
import os

# Hahanapin ang fcci.db sa loob ng "database" folder,
# kasing-level ng script na ito (parehong setup ng app.py mo)
db_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "database",
    "fcci.db"
)

if not os.path.exists(db_path):
    print("=" * 60)
    print("HINDI NAHANAP ANG fcci.db!")
    print(f"Hinanap dito: {db_path}")
    print("")
    print("Siguraduhin na:")
    print("1. Ang add_feed_tables.py ay naka-save sa PAREHONG")
    print("   folder ng iyong app.py")
    print("2. May 'database' folder sa loob nito na may fcci.db")
    print("=" * 60)
    exit()

conn = sqlite3.connect(db_path)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS feed_posts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id TEXT,
    full_name TEXT,
    content TEXT,
    photo_path TEXT,
    is_pinned INTEGER DEFAULT 0,
    post_date TEXT,
    post_time TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS feed_comments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    post_id INTEGER,
    member_id TEXT,
    full_name TEXT,
    comment_text TEXT,
    comment_date TEXT,
    comment_time TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS feed_likes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    post_id INTEGER,
    member_id TEXT,
    UNIQUE(post_id, member_id)
)
""")

conn.commit()
conn.close()

print("=" * 60)
print("SUCCESS! Feed tables ready.")
print(f"Na-update ang database dito: {db_path}")
print("=" * 60)
