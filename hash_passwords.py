HASH_SCRIPT = '''
import sqlite3
from werkzeug.security import generate_password_hash
 
conn = sqlite3.connect("database/fcci.db")
cursor = conn.cursor()
 
cursor.execute("SELECT id, username, password FROM users")
users = cursor.fetchall()
 
for user in users:
    user_id   = user[0]
    username  = user[1]
    password  = user[2]
 
    # Skip na hashed na
    if password.startswith("pbkdf2:") or password.startswith("scrypt:"):
        print(f"[SKIP] {username} - already hashed")
        continue
 
    hashed = generate_password_hash(password)
    cursor.execute("UPDATE users SET password=? WHERE id=?", (hashed, user_id))
    print(f"[HASHED] {username}")
 
conn.commit()
conn.close()
print("Done! All passwords hashed.")
'''