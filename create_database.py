import sqlite3
import os

db_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "fcci.db"
)

conn = sqlite3.connect(db_path)
cursor = conn.cursor()
try:
    cursor.execute("""
    ALTER TABLE members
    ADD COLUMN birthday TEXT
    """)
    conn.commit()
    print("Birthday column added.")
except:
    print("Birthday column already exists.")

# =========================
# MEMBERS
# =========================

cursor.execute("""
CREATE TABLE IF NOT EXISTS members (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id TEXT UNIQUE,
    full_name TEXT,
    contact TEXT,
    birthday TEXT,
    email TEXT,
    address TEXT,
    registration_fee INTEGER DEFAULT 0,
    date_registered TEXT
)
""")

# =========================
# PAYMENTS
# =========================

cursor.execute("""
CREATE TABLE IF NOT EXISTS payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_no TEXT,
    member_id TEXT,
    payment_type TEXT,
    amount INTEGER,
    payment_date TEXT,
    payment_year TEXT,
    payment_month TEXT
)
""")

# =========================
# EXPENSES
# =========================

cursor.execute("""
CREATE TABLE IF NOT EXISTS expenses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    expense_type TEXT,
    description TEXT,
    amount INTEGER,
    receipt_path TEXT,
    expense_date TEXT
)
""")

# =========================
# MEMBER PHOTOS
# =========================

cursor.execute("""
CREATE TABLE IF NOT EXISTS member_photos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id TEXT,
    photo_path TEXT
)
""")

# =========================
# ID CARDS
# =========================

cursor.execute("""
CREATE TABLE IF NOT EXISTS id_cards (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id TEXT,
    issue_date TEXT,
    qr_data TEXT
)
""")

# =========================
# ATTENDANCE
# =========================

cursor.execute("""
CREATE TABLE IF NOT EXISTS attendance (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id TEXT,
    attendance_date TEXT,
    attendance_time TEXT,
    time_out TEXT
)
""")

# =========================
# USERS
# =========================

cursor.execute("""
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE,
    password TEXT,
    role TEXT
)
""")

# DEFAULT USERS

cursor.execute("""
INSERT OR IGNORE INTO users
(username, password, role)
VALUES
('admin', 'admin123', 'Admin')
""")

cursor.execute("""
INSERT OR IGNORE INTO users
(username, password, role)
VALUES
('treasurer', 'treasurer123', 'Treasurer')
""")

cursor.execute("""
INSERT OR IGNORE INTO users
(username, password, role)
VALUES
('secretary', 'secretary123', 'Secretary')
""")

# =========================
# ACTIVITY LOGS
# =========================

cursor.execute("""
CREATE TABLE IF NOT EXISTS activity_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT,
    activity TEXT,
    log_date TEXT,
    log_time TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS donations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    donor_name TEXT,
    contact TEXT,
    amount INTEGER,
    purpose TEXT,
    donation_date TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_name TEXT,
    event_date TEXT,
    event_time TEXT,
    venue TEXT,
    description TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_type TEXT,
    event_name TEXT,
    event_date TEXT,
    event_time TEXT,
    venue TEXT,
    description TEXT,
    background_image TEXT,
    qr_enabled INTEGER DEFAULT 0
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS event_registrations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_id INTEGER,
    member_id TEXT,
    full_name TEXT,
    contact TEXT,
    registration_date TEXT
)
""")

try:
    cursor.execute("""
    ALTER TABLE members
    ADD COLUMN date_registered TEXT
    """)
    conn.commit()
    print("date_registered column added.")
except:
    print("date_registered column already exists.")


cursor.execute("""
CREATE TABLE IF NOT EXISTS withdrawals (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    member_id TEXT,

    full_name TEXT,

    total_contributions INTEGER,

    refund_amount INTEGER,

    community_share INTEGER,

    withdrawal_date TEXT
)
""")

try:
    cursor.execute("""
    ALTER TABLE members
    ADD COLUMN status TEXT DEFAULT 'Active'
    """)
    conn.commit()
except:
    pass


conn.commit()
conn.close()

print("================================")
print("FCCI DATABASE READY")
print("================================")
print(db_path)