from flask import (
    Flask,
    render_template,
    request,
    redirect,
    session,
    send_file,
    jsonify
)
import shutil
from werkzeug.utils import secure_filename
import os
from openpyxl import Workbook
import sqlite3
import psycopg2
from dotenv import load_dotenv
import cv2
import qrcode
from datetime import datetime
from cloudinary_helper import upload_photo, upload_file
from email_helper import send_welcome_email, send_payment_confirmation_email

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Image,
    Table
)

from reportlab.lib.styles import (
    getSampleStyleSheet
)

from flask import send_file
from werkzeug.security import check_password_hash, generate_password_hash

load_dotenv()

app = Flask(__name__)

app.secret_key = "FCCI_SECRET_KEY"

UPLOAD_FOLDER = "static/uploads"

app.config[
    "UPLOAD_FOLDER"
] = UPLOAD_FOLDER

DATABASE_URL = os.environ.get("DATABASE_URL")


def get_db():
    """
    Kumokonekta sa Supabase PostgreSQL database gamit ang
    connection string mula sa .env file (DATABASE_URL).
    """
    return psycopg2.connect(DATABASE_URL)


def fetch_member_with_photo(cursor, member_id):
    """
    Kinukuha ang lahat ng columns ng isang member (SELECT *) PLUS
    ang photo_path mula sa member_photos table, idinadagdag bilang
    HULING ELEMENT ng tuple. Ginagawa itong list para magamit ang
    .append(), tapos ibinabalik bilang tuple para gumana pa rin ang
    member[index] access pattern sa templates.

    Kung walang member na nahanap, nagbabalik ng None.
    """
    cursor.execute("SELECT id, member_id, full_name, contact, address, registration_fee, member_since, email, birthday, date_registered, status, proof_of_payment FROM members WHERE member_id = %s", (member_id,))
    row = cursor.fetchone()

    if row is None:
        return None

    cursor.execute("""
    SELECT photo_path FROM member_photos
    WHERE member_id = %s
    ORDER BY id DESC LIMIT 1
    """, (member_id,))

    photo_row = cursor.fetchone()
    photo_path = photo_row[0] if photo_row else None

    return tuple(list(row) + [photo_path])


def fetch_all_members_with_photo(cursor, where_clause="", params=()):
    """
    Kinukuha ang lahat ng members (o filtered gamit ang where_clause)
    PLUS photo_path bilang huling column ng bawat row. Ginagamit ito
    sa mga listahan tulad ng registration_approval at members list.
    """
    query = "SELECT id, member_id, full_name, contact, address, registration_fee, member_since, email, birthday, date_registered, status, proof_of_payment FROM members"
    if where_clause:
        query += f" WHERE {where_clause}"

    cursor.execute(query, params)
    rows = cursor.fetchall()

    results = []

    for row in rows:
        cursor.execute("""
        SELECT photo_path FROM member_photos
        WHERE member_id = %s
        ORDER BY id DESC LIMIT 1
        """, (row[1],))  # row[1] = member_id column

        photo_row = cursor.fetchone()
        photo_path = photo_row[0] if photo_row else None

        results.append(tuple(list(row) + [photo_path]))

    return results


def download_photo_for_pdf(member_id):
    """
    Kunin ang Cloudinary photo URL ng member mula sa member_photos
    table, i-download sa temporary file, at ibalik ang local path
    para magamit ng ReportLab Image(). Nagbabalik ng None kung
    walang photo o nag-error ang download.
    """
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("""
        SELECT photo_path FROM member_photos
        WHERE member_id = %s
        ORDER BY id DESC LIMIT 1
        """, (member_id,))
        photo_row = cursor.fetchone()
        conn.close()

        if photo_row and photo_row[0]:
            photo_url = photo_row[0]
            if photo_url.startswith("http"):
                import urllib.request
                import tempfile
                tmp_photo = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
                urllib.request.urlretrieve(photo_url, tmp_photo.name)
                return tmp_photo.name
    except Exception as e:
        print(f"[PDF] Photo download error: {e}")

    return None


@app.route("/")
def home():
    return redirect("/login")


@app.route(
    "/login",
    methods=["GET", "POST"]
)
def login():

    if request.method == "POST":

        username = request.form["username"]
        password = request.form["password"]

        conn = get_db()
        cursor = conn.cursor()

        cursor.execute("""
        SELECT role, password
        FROM users
        WHERE username = %s
        """, (username,))

        user = cursor.fetchone()
        conn.close()

        # Supports both hashed and plain passwords (para hindi masisira ang existing users)
        if user:
            stored_password = user[1]
            role = user[0]

            # Check if password is already hashed (starts with 'pbkdf2:' or 'scrypt:')
            if stored_password.startswith("pbkdf2:") or stored_password.startswith("scrypt:"):
                password_ok = check_password_hash(stored_password, password)
            else:
                # Plain text fallback (old accounts)
                password_ok = (stored_password == password)

            if password_ok:
                session["username"] = username
                session["role"] = role
                return redirect("/dashboard")

        return render_template(
            "login.html",
            error="Invalid username or password."
        )

    return render_template("login.html")

@app.route(
    "/member_registration",
    methods=["GET","POST"]
)
def member_registration():

    if request.method == "POST":

        conn = get_db()
        cursor = conn.cursor()

        # Hanapin ang pinakamataas na existing APP- number,
        # hindi yung total count (para hindi mag-clash kapag may
        # na-delete o na-convert na applicant sa gitna ng sequence)
        cursor.execute("""
        SELECT member_id
        FROM members
        WHERE member_id LIKE 'APP-%%'
        """)

        existing_app_ids = cursor.fetchall()

        highest_num = 0
        for row in existing_app_ids:
            try:
                num_part = int(row[0].split("-")[-1])
                if num_part > highest_num:
                    highest_num = num_part
            except (ValueError, IndexError):
                continue

        count = highest_num + 1

        member_id = (
            f"APP-{datetime.now().year}-{count:06d}"
        )

        full_name = request.form["full_name"]
        contact = request.form["contact"]
        birthday = request.form["birthday"]
        email = request.form["email"]
        address = request.form["address"]

        photo = request.files.get("photo")

        photo_filename = ""

        if photo and photo.filename:
            photo_filename = upload_photo(photo, folder="fcci_member_photos") or ""

        cursor.execute("""
        INSERT INTO members
        (
            member_id,
            full_name,
            contact,
            address,
            registration_fee,
            member_since,
            email,
            birthday,
            date_registered,
            status,
            photo_path
        )
        VALUES
        (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """,
        (
            member_id,
            full_name,
            contact,
            address,
            0,
            "",
            email,
            birthday,
            datetime.now().strftime("%Y-%m-%d"),
            "Applicant",
            photo_filename
        ))

        if photo_filename:
            cursor.execute("""
            INSERT INTO member_photos
            (member_id, photo_path)
            VALUES (%s, %s)
            """, (
                member_id,
                photo_filename
            ))

        conn.commit()
        conn.close()

        return redirect(
            f"/registration_confirmation/{member_id}"
        )

    return render_template(
        "member_registration.html"
    )

@app.route("/dashboard")
def dashboard():

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    # TOTAL MEMBERS
    cursor.execute("SELECT COUNT(*) FROM members")
    result = cursor.fetchone()
    total_members = result[0] if result else 0

    # COLLECTIONS
    cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM payments")
    result = cursor.fetchone()
    collections = result[0] if result else 0

    # DONATIONS
    cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM donations")
    result = cursor.fetchone()
    donations = result[0] if result else 0

    # EXPENSES
    cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM expenses")
    result = cursor.fetchone()
    expenses = result[0] if result else 0

    # BALANCE
    balance = collections + donations - expenses

    # RECENT PAYMENTS
    cursor.execute("""
    SELECT
        p.member_id,
        p.amount,
        m.full_name,
        p.payment_type,
        p.payment_month,
        p.payment_year
    FROM payments p
    LEFT JOIN members m ON p.member_id = m.member_id
    ORDER BY p.id DESC
    LIMIT 10
    """)
    recent_payments = cursor.fetchall()

    # APPLICANTS COUNT
    cursor.execute("SELECT COUNT(*) FROM members WHERE status='Applicant'")
    result = cursor.fetchone()
    applicants = result[0] if result else 0

    # ACTIVE / INACTIVE / OUTSTANDING
    active_members = 0
    inactive_members = 0
    total_outstanding = 0

    cursor.execute("""
        SELECT member_id, member_since
        FROM members
        WHERE status = 'Active'
    """)
    members = cursor.fetchall()

    current_date = datetime.now()

    month_map = {
        "January":1, "February":2, "March":3,
        "April":4, "May":5, "June":6,
        "July":7, "August":8, "September":9,
        "October":10, "November":11, "December":12
    }

    for member in members:
        member_id = member[0]
        member_since = member[1]

        if not member_since:
            continue

        try:
            parts = member_since.split()
            month_name = parts[0]
            year = int(parts[1])
            month = month_map[month_name]
        except:
            continue

        missing_count = 0

        while (
            year < current_date.year
            or (year == current_date.year and month <= current_date.month)
        ):
            month_name_loop = datetime(year, month, 1).strftime("%B")

            cursor.execute("""
            SELECT COUNT(*)
            FROM payments
            WHERE member_id = %s
            AND payment_type = 'Monthly Contribution'
            AND payment_month = %s
            AND payment_year = %s
            """, (member_id, month_name_loop, str(year)))

            result = cursor.fetchone()
            paid = result[0] if result else 0

            if paid == 0:
                missing_count += 1

            month += 1
            if month > 12:
                month = 1
                year += 1

        if missing_count < 5:
            active_members += 1
        else:
            inactive_members += 1

        total_outstanding += (missing_count * 10000)

    # ── BIRTHDAY THIS MONTH ──────────────────────────────────
    # Kinukuha ang lahat ng Active members na may birthday ngayong buwan
    current_month = current_date.month
    current_day   = current_date.day

    cursor.execute("""
    SELECT full_name, birthday, photo_path
    FROM members
    WHERE status = 'Active'
    AND birthday IS NOT NULL
    AND birthday != ''
    ORDER BY birthday
    """)

    all_members = cursor.fetchall()
    birthday_list = []

    for row in all_members:
        full_name  = row[0]
        birthday   = row[1]   # expected format: YYYY-MM-DD
        photo_path = row[2]

        if not birthday:
            continue

        try:
            # Try YYYY-MM-DD format first
            bday_obj = datetime.strptime(birthday, "%Y-%m-%d")
        except ValueError:
            try:
                # Fallback: MM/DD/YYYY
                bday_obj = datetime.strptime(birthday, "%m/%d/%Y")
            except ValueError:
                continue

        if bday_obj.month == current_month:
            is_today = (bday_obj.day == current_day)

            birthday_list.append({
                "full_name": full_name,
                "birthday":  bday_obj.strftime("%B %d"),
                "photo":     photo_path or "",
                "is_today":  is_today,
                "day":       bday_obj.day   # para ma-sort
            })

    # I-sort: today first, then by day of month
    birthday_list.sort(key=lambda x: (not x["is_today"], x["day"]))

    conn.close()

    return render_template(
        "dashboard.html",
        username=session["username"],
        total_members=total_members,
        active_members=active_members,
        inactive_members=inactive_members,
        applicants=applicants,
        collections=collections,
        donations=donations,
        expenses=expenses,
        balance=balance,
        total_outstanding=total_outstanding,
        birthday_list=birthday_list,
        recent_payments=recent_payments
    )

@app.route("/members")
def members():

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    search = request.args.get(
        "search",
        ""
    )

    if search:

        cursor.execute("""
        SELECT
            member_id,
            full_name,
            contact,
            email,
            member_since
        FROM members
        WHERE
            member_id LIKE %s
            OR full_name LIKE %s
        ORDER BY full_name
        """, (
            f"%{search}%",
            f"%{search}%"
        ))

    else:

        cursor.execute("""
        SELECT
            member_id,
            full_name,
            contact,
            email,
            member_since
        FROM members
        ORDER BY full_name
        """)

    members = cursor.fetchall()

    updated_members = []

    month_map = {
        "January":1,
        "February":2,
        "March":3,
        "April":4,
        "May":5,
        "June":6,
        "July":7,
        "August":8,
        "September":9,
        "October":10,
        "November":11,
        "December":12
    }

    current_date = datetime.now()

    for member in members:

        member_id = member[0]
        full_name = member[1]
        contact = member[2]
        email = member[3]
        member_since = member[4]

        status = "Applicant"

        if member_since:

            try:

                month_name, year = member_since.split()

                year = int(year)

                month = month_map[month_name]

                missing_count = 0

                temp_year = year
                temp_month = month

                while (
                    temp_year < current_date.year
                    or
                    (
                        temp_year == current_date.year
                        and temp_month <= current_date.month
                    )
                ):

                    current_month = datetime(
                        temp_year,
                        temp_month,
                        1
                   ).strftime("%B")

                    cursor.execute("""
                    SELECT COUNT(*)
                    FROM payments
                    WHERE member_id=%s
                    AND payment_type='Monthly Contribution'
                    AND payment_month=%s
                    AND payment_year=%s
                    """,(
                       member_id,
                       current_month,
                       str(temp_year)
                    ))

                    paid = cursor.fetchone()[0]

                    if paid == 0:
                        missing_count += 1

                    temp_month += 1

                    if temp_month > 12:
                        temp_month = 1
                        temp_year += 1

                if missing_count >= 5:
                    status = "Inactive"
                else:
                    status = "Active"

            except:
                status = "Active"

        updated_members.append(
            (
                member_id,
                full_name,
                contact,
                email,
                status
            )
        )

    members = updated_members

    conn.close()

    return render_template(
        "members.html",
        members=members,
        search=search,
        username=session["username"]
    )

@app.route(
    "/add_member",
    methods=["GET", "POST"]
)
def add_member():

    if "username" not in session:
        return redirect("/login")
    
    conn = get_db()
    cursor = conn.cursor()


    if request.method == "POST":

        # Hanapin ang pinakamataas na existing APP- number
        # (hindi FCCI-, dahil "Applicant" pa ang status nito —
        # dapat mag-bayad muna ng Registration Fee bago maging
        # Active member, gaya ng public registration flow)
        cursor.execute("""
        SELECT member_id
        FROM members
        WHERE member_id LIKE 'APP-%%'
        """)

        existing_app_ids = cursor.fetchall()

        highest_num = 0
        for row in existing_app_ids:
            try:
                num_part = int(row[0].split("-")[-1])
                if num_part > highest_num:
                    highest_num = num_part
            except (ValueError, IndexError):
                continue

        next_no = highest_num + 1

        current_year = datetime.now().year

        member_id = (
            f"APP-{current_year}-{next_no:06d}"
        )
        full_name = request.form["full_name"]
        contact = request.form["contact"]
        birthday = request.form["birthday"]
        email = request.form["email"]
        address = request.form["address"]
        photo = request.files["photo"]
        

        photo_filename = ""

        if photo and photo.filename:
            photo_filename = upload_photo(photo, folder="fcci_member_photos") or ""

        conn = get_db()
        cursor = conn.cursor()

        cursor.execute("""
        INSERT INTO members
        (
            member_id,
            full_name,
            contact,
            address,
            registration_fee,
            member_since,
            email,
            birthday,
            date_registered,
            status
        )
        VALUES
        (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            member_id,
            full_name,
            contact,
            address,
            0,
            "",
            email,
            birthday,
            datetime.now().strftime("%Y-%m-%d"),
            "Applicant"
        ))

        # I-upload ang photo sa Cloudinary at i-save sa member_photos
        if photo_filename:
            cursor.execute("""
            INSERT INTO member_photos (member_id, photo_path)
            VALUES (%s, %s)
            """, (member_id, photo_filename))

        conn.commit()
        conn.close()

        return redirect("/members")

    return render_template(
        "add_member.html",
        username=session["username"]
    )

@app.route(
    "/edit_member/<member_id>",
    methods=["GET", "POST"]
)
def edit_member(member_id):

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    if request.method == "POST":

        photo = request.files.get("photo")

        if photo and photo.filename:

            photo_url = upload_photo(photo, folder="fcci_member_photos")

            if photo_url:
                cursor.execute("""
                INSERT INTO member_photos (member_id, photo_path)
                VALUES (%s, %s)
                """, (
                    member_id,
                    photo_url
                ))

        cursor.execute("""
        UPDATE members
        SET
            full_name = %s,
            contact = %s,
            birthday = %s,
            email = %s,
            address = %s
        WHERE member_id = %s
        """, (
            request.form["full_name"],
            request.form["contact"],
            request.form["birthday"],
            request.form["email"],
            request.form["address"],
            member_id
        ))

        conn.commit()
        conn.close()

        return redirect(
            f"/view_member/{member_id}"
        )

    member = fetch_member_with_photo(cursor, member_id)

    conn.close()

    return render_template(
        "edit_member.html",
        member=member,
        username=session["username"]
    )


@app.route("/delete_member/<member_id>")
def delete_member(member_id):

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    # Payments
    cursor.execute("""
    DELETE FROM payments
    WHERE member_id = %s
    """, (member_id,))

    # Attendance
    cursor.execute("""
    DELETE FROM attendance
    WHERE member_id = %s
    """, (member_id,))

    # Photos
    cursor.execute("""
    DELETE FROM member_photos
    WHERE member_id = %s
    """, (member_id,))

    # Withdrawals
    cursor.execute("""
    DELETE FROM withdrawals
    WHERE member_id = %s
    """, (member_id,))

    # ID Cards
    cursor.execute("""
    DELETE FROM id_cards
    WHERE member_id = %s
    """, (member_id,))

    # Main Member Record
    cursor.execute("""
    DELETE FROM members
    WHERE member_id = %s
    """, (member_id,))

    conn.commit()
    conn.close()

    return redirect("/members")


@app.route("/view_member/<member_id>")
def view_member(member_id):

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    member = fetch_member_with_photo(cursor, member_id)

    conn.close()

    return render_template(
        "view_member.html",
        member=member,
        username=session["username"]
    )


@app.route("/check_duplicate_payment")
def check_duplicate_payment():

    if "username" not in session:
        return jsonify({"error": "Not logged in"}), 401

    member_id = request.args.get("member_id", "").strip()
    payment_type = request.args.get("payment_type", "")
    payment_month = request.args.get("payment_month", "")
    payment_year = request.args.get("payment_year", "")

    conn = get_db()
    cursor = conn.cursor()

    if payment_type == "Registration Fee":

        cursor.execute("""
        SELECT status FROM members WHERE member_id = %s
        """, (member_id,))

        row = cursor.fetchone()
        conn.close()

        if row and row[0] == "Active":
            return jsonify({
                "duplicate": True,
                "message": f"Ang member na ito ({member_id}) ay Active na at nakabayad na ng Registration Fee. Hindi na ito dapat bayaran ulit."
            })

        return jsonify({"duplicate": False})

    else:
        cursor.execute("""
        SELECT COUNT(*) FROM payments
        WHERE member_id = %s
        AND payment_type = 'Monthly Contribution'
        AND payment_month = %s
        AND payment_year = %s
        """, (member_id, payment_month, payment_year))

        count = cursor.fetchone()[0]
        conn.close()

        if count > 0:
            return jsonify({
                "duplicate": True,
                "message": f"Nakabayad na ang member na ito ng Monthly Contribution para sa {payment_month} {payment_year}. Hindi na ito dapat bayaran ulit."
            })

        return jsonify({"duplicate": False})


@app.route("/search_member_payments")
def search_member_payments():

    if "username" not in session:
        return jsonify({"error": "Not logged in"}), 401

    search_term = request.args.get("name", "").strip()

    if not search_term:
        return jsonify({"results": []})

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT
        p.id,
        p.receipt_no,
        p.member_id,
        p.payment_type,
        p.amount,
        p.payment_date,
        m.full_name,
        p.payment_month,
        p.payment_year
    FROM payments p
    LEFT JOIN members m ON p.member_id = m.member_id
    WHERE LOWER(m.full_name) LIKE LOWER(%s)
    ORDER BY p.id DESC
    """, (f"%{search_term}%",))

    rows = cursor.fetchall()
    conn.close()

    results = []
    for r in rows:
        results.append({
            "id": r[0],
            "receipt_no": r[1],
            "member_id": r[2],
            "payment_type": r[3],
            "amount": r[4],
            "payment_date": r[5],
            "full_name": r[6] or r[2],
            "payment_month": r[7],
            "payment_year": r[8]
        })

    return jsonify({"results": results})


@app.route(
    "/payments",
    methods=["GET", "POST"]
)
def payments():

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    if request.method == "POST":

        member_id = request.form["member_id"]
        payment_type = request.form["payment_type"]

        payment_month = request.form["payment_month"]
        payment_year = request.form["payment_year"]

        if payment_type == "Registration Fee":
            amount = 20000
        else:
            amount = 10000

        payment_date = datetime.now().strftime(
            "%Y-%m-%d"
        )

        cursor.execute("""
        SELECT COUNT(*)
        FROM payments
        """)

        count = cursor.fetchone()[0] + 1

        receipt_no = (
            f"RCPT-{datetime.now().year}-{count:06d}"
        )

        cursor.execute("""
        SELECT COUNT(*)
        FROM members
        WHERE member_id = %s
        """, (member_id,))

        exists = cursor.fetchone()[0]

        if exists == 0:

            conn.close()

            return "Member ID Not Found"

        # Kunin ang full_name at email ng member BAGO mag-INSERT,
        # para magamit sa pagpapadala ng confirmation email mamaya
        cursor.execute("""
        SELECT full_name, email FROM members WHERE member_id = %s
        """, (member_id,))
        payer_info = cursor.fetchone()

        cursor.execute("""
        INSERT INTO payments
        (
            receipt_no,
            member_id,
            payment_type,
            amount,
            payment_date,
            payment_year,
            payment_month
            
        )
        VALUES
        (%s, %s, %s, %s, %s, %s, %s)
        """, (
            receipt_no,
            member_id,
            payment_type,
            amount,
            payment_date,
            payment_year,
            payment_month
            
        ))

        if payment_type == "Registration Fee":

            # Gumamit ng MAX existing number imbes na COUNT(*),
            # para hindi mag-clash kapag may na-delete na member
            # sa gitna ng sequence (hal. FCCI-2026-000002 na-delete,
            # pero may FCCI-2026-000003 pa rin)
            cursor.execute("""
            SELECT member_id
            FROM members
            WHERE member_id LIKE 'FCCI-%%'
            """)

            existing_ids = cursor.fetchall()

            highest_num = 0
            for row in existing_ids:
                try:
                    num_part = int(row[0].split("-")[-1])
                    if num_part > highest_num:
                        highest_num = num_part
                except (ValueError, IndexError):
                    continue

            count = highest_num + 1

            new_member_id = (
                f"FCCI-{datetime.now().year}-{count:06d}"
            )

            member_since = (
                f"{payment_month} {payment_year}"
            )

            old_member_id = member_id

            cursor.execute("""
            UPDATE members
            SET
                member_id = %s,
                registration_fee = 20000,
                member_since = %s,
                status = 'Active'
            WHERE member_id = %s
            """, (
                new_member_id,
                member_since,
                old_member_id
            ))

            cursor.execute("""
            UPDATE payments
            SET member_id = %s
            WHERE member_id = %s
            """, (
                new_member_id,
                old_member_id
            ))

            # I-update din ang member_photos para hindi mawala ang photo
            cursor.execute("""
            UPDATE member_photos
            SET member_id = %s
            WHERE member_id = %s
            """, (
                new_member_id,
                old_member_id
            ))

        conn.commit()

        # I-send ang payment confirmation email (hindi mag-crash
        # ang app kung walang email o nag-error ang pagpapadala)
        if payer_info and payer_info[1]:
            try:
                send_payment_confirmation_email(
                    payer_info[1],
                    payer_info[0],
                    payment_type,
                    amount,
                    receipt_no,
                    payment_date
                )
            except Exception as e:
                print(f"[EMAIL] Failed to send payment confirmation: {e}")

    cursor.execute("""
    SELECT
        member_id,
        full_name,
        status
    FROM members
    ORDER BY full_name
    """)

    members = cursor.fetchall()

    cursor.execute("""
    SELECT
        p.id,
        p.receipt_no,
        p.member_id,
        p.payment_type,
        p.amount,
        p.payment_date,
        m.full_name,
        p.payment_month,
        p.payment_year
    FROM payments p
    LEFT JOIN members m ON p.member_id = m.member_id
    ORDER BY p.id DESC
    LIMIT 50
    """)

    payment_history = cursor.fetchall()

    conn.close()

    return render_template(
        "payments.html",
        members=members,
        payment_history=payment_history,
        username=session["username"]
    )


@app.route("/delete_payment/<int:payment_id>")
def delete_payment(payment_id):

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute(
        "DELETE FROM payments WHERE id = %s",
        (payment_id,)
    )

    conn.commit()
    conn.close()

    return redirect("/payments")


@app.route(
    "/member_id_card/<member_id>"
)
def member_id_card(member_id):

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    member = fetch_member_with_photo(cursor, member_id)

    conn.close()

    if not member:
        return "Member Not Found"

    qr = qrcode.make(member_id)

    qr_path = (
        f"static/qr/{member_id}.png"
    )

    qr.save(qr_path)

    return render_template(
        "member_id_card.html",
        member=member,
        qr_file=f"{member_id}.png"
    )


@app.route("/bulk_id_cards", methods=["POST"])
def bulk_id_cards():

    if "username" not in session:
        return redirect("/login")

    member_ids = request.form.getlist("selected_members")

    if not member_ids:
        return redirect("/members")

    conn = get_db()
    cursor = conn.cursor()

    members_list = []

    for mid in member_ids:

        member = fetch_member_with_photo(cursor, mid)

        if member:

            qr = qrcode.make(mid)
            qr_path = f"static/qr/{mid}.png"
            qr.save(qr_path)

            members_list.append({
                "member": member,
                "qr_file": f"{mid}.png"
            })

    conn.close()

    return render_template(
        "bulk_id_cards.html",
        members_list=members_list
    )


@app.route(
"/expenses",
methods=["GET","POST"]
)
def expenses():

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    if request.method == "POST":

        expense_type = request.form["expense_type"]
        description = request.form["description"]

        try:
            amount = int(
                request.form["amount"]
            )
        except:
            amount = 0

        receipt_file = request.files.get(
            "receipt"
        )

        receipt_filename = ""

        if (
            receipt_file
            and
            receipt_file.filename
        ):

            receipt_filename = upload_file(receipt_file, folder="fcci_receipts") or ""

        cursor.execute("""
        INSERT INTO expenses
        (
            expense_type,
            description,
            amount,
            receipt_path,
            expense_date
        )
        VALUES
        (%s, %s, %s, %s, %s)
        """, (
            expense_type,
            description,
            amount,
            receipt_filename,
            datetime.now().strftime(
               "%Y-%m-%d"
           )
       ))
    
        conn.commit()

    cursor.execute("""
    SELECT
        id,
        expense_type,
        description,
        amount,
        expense_date,
        receipt_path
    FROM expenses
    ORDER BY id DESC
    """)

    expense_history = cursor.fetchall()

    conn.close()

    return render_template(
        "expenses.html",
        expense_history=expense_history,
        username=session["username"]
    )

@app.route(
"/donations",
methods=["GET","POST"]
)
def donations():


    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    if request.method == "POST":

        donor_name = request.form["donor_name"]
        contact = request.form["contact"]
        amount = int(request.form["amount"])
        purpose = request.form["purpose"]

        cursor.execute("""
        INSERT INTO donations
        (
            donor_name,
            contact,
            amount,
            purpose,
            donation_date
        )
        VALUES
        (%s, %s, %s, %s, %s)
        """, (
            donor_name,
            contact,
            amount,
            purpose,
            datetime.now().strftime(
                "%Y-%m-%d"
            )
        ))

        conn.commit()

    cursor.execute("""
    SELECT
        id,
        donor_name,
        contact,
        amount,
        purpose,
        donation_date
    FROM donations
    ORDER BY id DESC
    """)

    donation_history = cursor.fetchall()

    conn.close()

    return render_template(
        "donations.html",
        donation_history=donation_history,
        username=session["username"]
    )

@app.route("/exports")
def exports():


    if "username" not in session:
        return redirect("/login")
    
    return render_template(
        "exports.html",
        username=session["username"]
    )


@app.route("/export_members")
def export_members():

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT
        member_id,
        full_name,
        contact,
        email,
        address,
        member_since,
        status
    FROM members
    ORDER BY full_name
    """)

    rows = cursor.fetchall()

    conn.close()

    wb = Workbook()

    ws = wb.active

    ws.title = "Members"

    ws.append([
        "Member ID",
        "Full Name",
        "Contact",
        "Email",
        "Address",
        "Member Since",
        "Status"
    ])

    for row in rows:
        ws.append(row)

    export_dir = "exports"

    os.makedirs(
        export_dir,
        exist_ok=True
    )

    filename = os.path.join(
        export_dir,
        "FCCI_Members.xlsx"
    )

    wb.save(filename)

    return send_file(
        filename,
        as_attachment=True
    )

@app.route("/export_payments")
def export_payments():


    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT
        receipt_no,
        member_id,
        payment_type,
        amount,
        payment_month,
        payment_year,
        payment_date
    FROM payments
    ORDER BY id DESC
    """)

    rows = cursor.fetchall()

    conn.close()

    wb = Workbook()

    ws = wb.active

    ws.title = "Payments"

    ws.append([
        "Receipt No",
        "Member ID",
        "Payment Type",
        "Amount",
        "Month",
        "Year",
        "Payment Date"
    ])

    for row in rows:
        ws.append(row)

    export_dir = "exports"

    os.makedirs(
        export_dir,
        exist_ok=True
    )

    filename = os.path.join(
        export_dir,
        "FCCI_Payments.xlsx"
    )

    wb.save(filename)

    return send_file(
        filename,
        as_attachment=True
    )

@app.route("/export_donations")
def export_donations():


    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT
        donor_name,
        contact,
        amount,
        purpose,
        donation_date
    FROM donations
    ORDER BY id DESC
    """)

    rows = cursor.fetchall()

    conn.close()

    wb = Workbook()

    ws = wb.active

    ws.title = "Donations"

    ws.append([
        "Donor Name",
        "Contact",
        "Amount",
        "Purpose",
        "Donation Date"
    ])

    for row in rows:
        ws.append(row)

    export_dir = "exports"

    os.makedirs(
        export_dir,
        exist_ok=True
    )

    filename = os.path.join(
        export_dir,
        "FCCI_Donations.xlsx"
    )

    wb.save(filename)

    return send_file(
        filename,
        as_attachment=True
    )

@app.route(
"/donation_certificate/<int:donation_id>"
)
def donation_certificate(donation_id):


    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT
        donor_name,
        contact,
        amount,
        purpose,
        donation_date
    FROM donations
    WHERE id = %s
    """, (
        donation_id,
    ))

    donation = cursor.fetchone()

    conn.close()

    if not donation:
        return "Donation Not Found"

    os.makedirs(
        "exports",
        exist_ok=True
    )

    filename = (
        f"exports/"
        f"Donation_Certificate_{donation_id}.pdf"
    )

    pdf = SimpleDocTemplate(
        filename
    )

    styles = getSampleStyleSheet()

    content = []

    logo_path = "logo/fcci_logo.jpeg"

    if os.path.exists(
        logo_path
    ):

        content.append(
            Image(
                logo_path,
                width=100,
                height=100
            )
        )

    content.append(
        Spacer(1,10)
    )

    content.append(
        Paragraph(
            "FILIPINO COMMUNITY CENTER INTERNATIONAL",
            styles["Title"]
        )
    )

    content.append(
        Spacer(1,20)
    )

    content.append(
        Paragraph(
            "CERTIFICATE OF APPRECIATION",
            styles["Title"]
        )
    )

    content.append(
        Spacer(1,30)
    )

    content.append(
        Paragraph(
            f"""
            This certificate is proudly presented to

            <b>{donation[0]}</b>

            In recognition and appreciation of your
            generous donation and support to FCCI.
            """,
            styles["BodyText"]
        )
    )

    content.append(
        Spacer(1,20)
    )

    content.append(
        Paragraph(
            f"Donation Amount: ₩{donation[2]:,}",
            styles["BodyText"]
        )
    )

    content.append(
        Paragraph(
            f"Purpose: {donation[3]}",
            styles["BodyText"]
        )
    )

    content.append(
        Paragraph(
            f"Date: {donation[4]}",
            styles["BodyText"]
        )
    )

    content.append(
        Spacer(1,50)
    )

    content.append(
        Paragraph(
            "TO GOD BE THE GLORY!",
            styles["Heading2"]
        )
    )

    content.append(
        Spacer(1,40)
    )

    content.append(
        Paragraph(
            "________________________",
            styles["BodyText"]
        )
    )

    content.append(
        Paragraph(
            "FCCI President",
            styles["BodyText"]
        )
    )

    pdf.build(content)

    return send_file(
        filename,
        as_attachment=True
    )


@app.route("/export_expenses")
def export_expenses():


    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT
        expense_type,
        description,
        amount,
        expense_date
    FROM expenses
    ORDER BY id DESC
    """)

    rows = cursor.fetchall()

    conn.close()

    wb = Workbook()

    ws = wb.active

    ws.title = "Expenses"

    ws.append([
        "Expense Type",
        "Description",
        "Amount",
        "Expense Date"
    ])

    for row in rows:
        ws.append(row)

    export_dir = "exports"

    os.makedirs(
        export_dir,
        exist_ok=True
    )

    filename = os.path.join(
        export_dir,
        "FCCI_Expenses.xlsx"
    )

    wb.save(filename)
    
    return send_file(
        filename,
        as_attachment=True
    )

@app.route("/export_attendance")
def export_attendance():

    if "username" not in session:
        return redirect("/login")

    export_date = request.args.get(
        "export_date"
    )

    if not export_date:
        return "Please select a date."

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT
        a.member_id,
        m.full_name,
        a.attendance_date,
        a.attendance_time,
        a.time_out
    FROM attendance a
    LEFT JOIN members m
    ON a.member_id = m.member_id
    WHERE a.attendance_date = %s
    ORDER BY a.id DESC
    """, (
        export_date,
    ))

    rows = cursor.fetchall()

    conn.close()

    wb = Workbook()

    ws = wb.active

    ws.title = "Attendance"

    ws.append([
        "Member ID",
        "Full Name",
        "Date",
        "Time In",
        "Time Out"
    ])

    for row in rows:
        ws.append(row)

    export_dir = "exports"

    os.makedirs(
        export_dir,
        exist_ok=True
    )

    filename = os.path.join(
        export_dir,
        f"Attendance_{export_date}.xlsx"
    )

    wb.save(filename)

    return send_file(
        filename,
        as_attachment=True
    )

@app.route("/export_monitoring")
def export_monitoring():


    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT
        member_id,
        full_name
    FROM members
    ORDER BY full_name
    """)

    members = cursor.fetchall()

    wb = Workbook()

    ws = wb.active

    ws.title = "Monthly Monitoring"

    ws.append([
        "Member ID",
        "Full Name",
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec"
    ])

    month_map = {
        "January":"Jan",
        "February":"Feb",
        "March":"Mar",
        "April":"Apr",
        "May":"May",
        "June":"Jun",
        "July":"Jul",
        "August":"Aug",
        "September":"Sep",
        "October":"Oct",
        "November":"Nov",
        "December":"Dec"
    }

    for member in members:

        member_id = member[0]
        full_name = member[1]

        row = {
            "Jan":"-",
            "Feb":"-",
            "Mar":"-",
            "Apr":"-",
            "May":"-",
            "Jun":"-",
            "Jul":"-",
            "Aug":"-",
            "Sep":"-",
            "Oct":"-",
            "Nov":"-",
            "Dec":"-"
        }

        cursor.execute("""
        SELECT payment_month
        FROM payments
        WHERE member_id = %s
        """, (member_id,))

        payments = cursor.fetchall()

        for p in payments:

            month = p[0]

            if month in month_map:

                row[
                    month_map[month]
                ] = "✔"

        ws.append([
            member_id,
            full_name,
            row["Jan"],
            row["Feb"],
            row["Mar"],
            row["Apr"],
            row["May"],
            row["Jun"],
            row["Jul"],
            row["Aug"],
            row["Sep"],
            row["Oct"],
            row["Nov"],
            row["Dec"]
        ])

    conn.close()

    export_dir = "exports"

    os.makedirs(
        export_dir,
        exist_ok=True
    )

    filename = os.path.join(
        export_dir,
        "FCCI_Monthly_Monitoring.xlsx"
    )

    wb.save(filename)

    return send_file(
        filename,
        as_attachment=True
    )

@app.route(
"/attendance",
methods=["GET","POST"]
)
def attendance():


    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    message = ""

    if request.method == "POST":

        member_id = request.form[
            "member_id"
        ].strip()

        cursor.execute("""
        SELECT full_name
        FROM members
        WHERE member_id = %s
        """, (
            member_id,
        ))

        member = cursor.fetchone()

        if member:

            today = datetime.now().strftime(
                "%Y-%m-%d"
            )

            current_time = datetime.now().strftime(
                "%H:%M:%S"
            )

            cursor.execute("""
            SELECT
                id,
                time_out
            FROM attendance
            WHERE member_id = %s
            AND attendance_date = %s
            """, (
                member_id,
                today
            ))

            record = cursor.fetchone()

            if record:

                if not record[1]:

                    cursor.execute("""
                    UPDATE attendance
                    SET time_out = %s
                    WHERE id = %s
                    """, (
                        current_time,
                        record[0]
                    ))

                    conn.commit()

                    message = (
                        f"{member[0]}"
                        f" TIME OUT "
                        f"{current_time}"
                    )

                else:

                    message = (
                        "Already Timed Out Today"
                    )

            else:

                cursor.execute("""
                INSERT INTO attendance
                (
                    member_id,
                    attendance_date,
                    attendance_time,
                    time_out
                )
                VALUES (%s, %s, %s, %s)
                """, (
                    member_id,
                    today,
                    current_time,
                    ""
                ))


                conn.commit()

                message = (
                    f"{member[0]}"
                    f" TIME IN "
                    f"{current_time}"
                )

        else:

            message = "Member Not Found"

    today = datetime.now().strftime(
        "%Y-%m-%d"
    )

    cursor.execute("""
    SELECT
        a.member_id,
        m.full_name,
        a.attendance_date,
        a.attendance_time,
        a.time_out
    FROM attendance a
    LEFT JOIN members m
    ON a.member_id = m.member_id
    WHERE attendance_date = %s
    ORDER BY a.id DESC
    """, (
        today,
    ))

    attendance_history = (
        cursor.fetchall()
    )

    conn.close()

    return render_template(
        "attendance.html",
        attendance_history=attendance_history,
        message=message,
        username=session["username"]
    )

@app.route("/qr_attendance")
def qr_attendance():


    if "username" not in session:
        return redirect("/login")

    cap = cv2.VideoCapture(0)

    detector = cv2.QRCodeDetector()

    while True:

        success, frame = cap.read()

        if not success:
            break

        data, bbox, _ = detector.detectAndDecode(
            frame
        )

        if data:

            member_id = data.strip()

            conn = get_db()
            cursor = conn.cursor()

            cursor.execute("""
            SELECT full_name
            FROM members
            WHERE member_id = %s
            """, (
                member_id,
            ))

            member = cursor.fetchone()

            if member:

                today = datetime.now().strftime(
                    "%Y-%m-%d"
                )

                current_time = datetime.now().strftime(
                    "%H:%M:%S"
                )

                cursor.execute("""
                SELECT
                    id,
                    time_out
                FROM attendance
                WHERE member_id = %s
                AND attendance_date = %s
                """, (
                    member_id,
                    today
                ))

                record = cursor.fetchone()

                if record:

                    if not record[1]:

                        cursor.execute("""
                        UPDATE attendance
                        SET time_out = %s
                        WHERE id = %s
                        """, (
                            current_time,
                            record[0]
                        ))

                    conn.commit()

                else:

                    cursor.execute("""
                    INSERT INTO attendance
                    (
                        member_id,
                        attendance_date,
                        attendance_time,
                        time_out
                    )
                    VALUES (%s, %s, %s, %s)
                    """, (
                        member_id,
                        today,
                        current_time,
                        ""
                    ))

                    conn.commit()

            conn.close()

            break

        cv2.imshow(
            "FCCI QR Attendance",
            frame
        )

        if cv2.waitKey(1) == 27:
            break

        import time

        current_scan_time = time.time()

        if data:

            member_id = data.strip()

            if (
                member_id == last_scan
                and
                current_scan_time - last_scan_time < 5
            ):
                pass

            else:

                last_scan = member_id
                last_scan_time = current_scan_time

                # attendance logic dito

    cap.release()

    cv2.destroyAllWindows()

    return redirect("/attendance")

@app.route(
    "/qr_attendance_scan",
    methods=["POST"]
)
def qr_attendance_scan():

    if "username" not in session:
        return {"success":False}

    member_id = request.form["member_id"]

    

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT full_name
    FROM members
    WHERE member_id = %s
    """, (member_id,))

    member = cursor.fetchone()

    if not member:

        conn.close()

        return {
            "success":False,
            "message":"Member Not Found"
        }

    today = datetime.now().strftime(
        "%Y-%m-%d"
    )

    current_time = datetime.now().strftime(
        "%H:%M:%S"
    )

    cursor.execute("""
    SELECT id,time_out
    FROM attendance
    WHERE member_id = %s
    AND attendance_date = %s
    """, (
        member_id,
        today
    ))

    record = cursor.fetchone()

    if record:

        if not record[1]:

            cursor.execute("""
            UPDATE attendance
            SET time_out = %s
            WHERE id = %s
            """, (
                current_time,
                record[0]
            ))

            action = "TIME OUT"
            print("TIME OUT SAVED:", member_id)

        else:

            conn.close()

            return {
                "success":True,
                "message":"Already Timed Out"
            }

    else:

        cursor.execute("""
        INSERT INTO attendance
        (
            member_id,
            attendance_date,
            attendance_time,
            time_out
        )
        VALUES (%s, %s, %s, %s)
        """, (
            member_id,
            today,
            current_time,
            ""
        ))

        action = "TIME IN"
        print("TIME IN SAVED:", member_id)

    conn.commit()

    cursor.execute("""
    SELECT COUNT(*)
    FROM attendance
    WHERE member_id = %s
    AND attendance_date = %s
    """, (
        member_id,
        today
    ))

    count = cursor.fetchone()[0]

    conn.close()

    return {
        "success":True,
        "message":
        f"{member[0]} {action} {current_time}",
        "count": count
    }

@app.route("/member_profile/<member_id>")
def member_profile(member_id):

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    member = fetch_member_with_photo(cursor, member_id)

    if not member:

        conn.close()
        return "Member Not Found"

    cursor.execute("""
    SELECT
        payment_date,
        payment_type,
        amount,
        payment_month,
        payment_year
    FROM payments
    WHERE member_id = %s
    ORDER BY id DESC
    """, (member_id,))

    payments = cursor.fetchall()

    cursor.execute("""
    SELECT
        COALESCE(SUM(amount),0)
    FROM payments
    WHERE member_id = %s
    """, (member_id,))

    total_payment = cursor.fetchone()[0]
    
    cursor.execute("""
    SELECT attendance_date
    FROM attendance
    WHERE member_id = %s
    ORDER BY id DESC
    LIMIT 1
    """, (member_id,))

    attendance = cursor.fetchone()

    conn.close()

    return render_template(
        "member_profile.html",
        member=member,
        payments=payments,
        attendance=attendance,
        total_payment=total_payment
    )

@app.route(
    "/withdrawals",
    methods=["GET", "POST"]
)
def withdrawals():

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    member = None
    total_contributions = 0
    refund_amount = 0
    community_share = 0

    if request.method == "POST":

        action = request.form.get("action")

        member_id = request.form["member_id"]

        member = fetch_member_with_photo(cursor, member_id)

        if member:

            cursor.execute("""
            SELECT
                COALESCE(SUM(amount),0)
            FROM payments
            WHERE member_id = %s
            AND payment_type = 'Monthly Contribution'
            """, (member_id,))

            total_contributions = (
                cursor.fetchone()[0]
            )

            refund_amount = int(
                total_contributions * 0.75
            )

            community_share = (
                total_contributions
                - refund_amount
            )

            

            if action == "finalize":

                cursor.execute("""
                INSERT INTO withdrawals
                (
                    member_id,
                    full_name,
                    total_contributions,
                    refund_amount,
                    community_share,
                    withdrawal_date
                )
                VALUES
                (%s, %s, %s, %s, %s, %s)
                """, (
                    member[1],
                    member[2],
                    total_contributions,
                    refund_amount,
                    community_share,
                    datetime.now().strftime(
                        "%Y-%m-%d"
                    )
                ))

                cursor.execute("""
                DELETE FROM payments
                WHERE member_id = %s
                """, (member_id,))

                cursor.execute("""
                DELETE FROM attendance
                WHERE member_id = %s
                """, (member_id,))

                cursor.execute("""
                DELETE FROM members
                WHERE member_id = %s
                """, (member_id,))
                

                conn.commit()

                conn.close()

                return redirect(
                    "/withdrawals"
                )

    cursor.execute("""
    SELECT *
    FROM withdrawals
    ORDER BY id DESC
    """)

    withdrawal_history = (
        cursor.fetchall()
    )

    conn.close()

    return render_template(
        "withdrawals.html",
        member=member,
        total_contributions=
        total_contributions,
        refund_amount=
        refund_amount,
        community_share=
        community_share,
        withdrawal_history=
        withdrawal_history,
        username=session["username"]
    )

@app.route(
"/withdrawal_certificate/<member_id>"
)
def withdrawal_certificate(member_id):


    if "username" not in session:
        return redirect("/login")

    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Image,
        Table,
        TableStyle
    )

    from reportlab.lib import colors
    from reportlab.lib.styles import (
        getSampleStyleSheet
    )

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT
    id, member_id, full_name, contact, address, registration_fee, member_since, email, birthday, date_registered, status, proof_of_payment
    FROM members
    WHERE member_id = %s
    """, (member_id,))

    member = cursor.fetchone()

    if not member:
        conn.close()
        return "Member Not Found"
    
    cursor.execute("""
    SELECT
        COALESCE(SUM(amount),0)
    FROM payments
    WHERE member_id = %s
    AND payment_type =
    'Monthly Contribution'
    """, (member_id,))

    total_contributions = (
        cursor.fetchone()[0]
    )

    refund_amount = int(
        total_contributions * 0.75
    )

    community_share = (
        total_contributions
        - refund_amount
    )

    conn.close()

    os.makedirs(
        "exports",
        exist_ok=True
    )

    filename = (
        f"exports/"
        f"Withdrawal_{member_id}.pdf"
    )

    doc = SimpleDocTemplate(
        filename
    )

    styles = getSampleStyleSheet()

    story = []

    logo_path = (
        "logo/fcci_logo.jpeg"
    )

    if os.path.exists(
        logo_path
    ):
        logo = Image(
            logo_path,
            width=100,
            height=100
        )
        story.append(logo)

    story.append(
        Paragraph(
            "<b>FILIPINO COMMUNITY CENTER INTERNATIONAL</b>",
            styles["Title"]
        )
    )

    story.append(
        Paragraph(
            "Membership Withdrawal Certificate",
            styles["Heading2"]
        )
    )

    story.append(
        Spacer(1,20)
    )


    photo_path = download_photo_for_pdf(member[1])

    if photo_path:

        story.append(
            Image(
                photo_path,
                width=120,
                height=120
            )
        )

        story.append(
            Spacer(1,10)
        )

    data = [

        [
            "Member ID",
            member[1]
        ],

        [
            "Full Name",
            member[2]
        ],

        [
            "Contact",
            member[3]
        ],

        [
            "Email",
            member[7]
        ],

        [
            "Address",
            member[4]
        ],

        [
            "Withdrawal Date",
            datetime.now().strftime(
                "%Y-%m-%d"
            )
        ],

        [
            "Total Contributions",
            f"₩{total_contributions:,}"
        ],

        [
            "Refund (75%)",
            f"₩{refund_amount:,}"
        ],

        [
            "Community Share (25%)",
            f"₩{community_share:,}"
        ]

    ]

    table = Table(
        data,
        colWidths=[
            180,
            280
        ]
    )

    table.setStyle(
        TableStyle([
            (
                "GRID",
                (0,0),
                (-1,-1),
                1,
                colors.black
            ),
            (
                "BACKGROUND",
                (0,0),
                (0,-1),
                colors.lightgrey
            )
        ])
    )

    story.append(
        table
    )

    story.append(
        Spacer(1,40)
    )

    story.append(
        Paragraph(
            "This certifies that the above member voluntarily withdrew from FCCI membership.",
            styles["Normal"]
        )
    )

    story.append(
        Spacer(1,50)
    )

    signature_table = Table(
        
        [

            [
                "Member Signature",
                "President Signature"
            ],

            [
                "",
                ""
            ],

            [
                "__________________",
                "__________________"
            ]

        ],

        colWidths=[
            250,
            250
        ]

    )

    story.append(
        signature_table
    )

    doc.build(
        story
    )

    return send_file(
        filename,
        as_attachment=True
    )

@app.route(
"/member_profile_pdf/<member_id>"
)
def member_profile_pdf(member_id):


    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT
    id, member_id, full_name, contact, address, registration_fee, member_since, email, birthday, date_registered, status, proof_of_payment
    FROM members
    WHERE member_id = %s
    """, (member_id,))

    member = cursor.fetchone()

    if not member:

        conn.close()

        return "Member Not Found"

    cursor.execute("""
    SELECT
        COALESCE(SUM(amount),0)
    FROM payments
    WHERE member_id = %s
    """, (member_id,))

    total_payment = cursor.fetchone()[0]

    cursor.execute("""
    SELECT attendance_date
    FROM attendance
    WHERE member_id = %s
    ORDER BY id DESC
    LIMIT 1
    """, (member_id,))
    
    attendance = cursor.fetchone()

    cursor.execute("""
    SELECT
        payment_date,
        payment_type,
        amount
    FROM payments
    WHERE member_id = %s
    ORDER BY id DESC
    """, (member_id,))

    payments = cursor.fetchall()

    conn.close()

    os.makedirs(
        "exports",
        exist_ok=True
    )

    filename = (
        f"exports/"
        f"{member_id}_Profile.pdf"
    )

    pdf = SimpleDocTemplate(
        filename
    )

    styles = getSampleStyleSheet()

    content = []

    logo_path = (
        "logo/fcci_logo.jpeg"
    )

    if os.path.exists(
        logo_path
    ):

        logo = Image(
            logo_path,
            width=70,
            height=70
        )

        header = Table(
            [[
                Paragraph(
                    "FCCI MEMBER PROFILE",
                    styles["Title"]
                ),
                logo
            ]],
            colWidths=[400,80]
        )

        content.append(
            header
        )

    content.append(
        Spacer(1,20)
    )

    photo_path = download_photo_for_pdf(member[1])

    if photo_path:

        img = Image(
            photo_path,
            width=120,
            height=120
        )

        content.append(img)

        content.append(
            Spacer(1,10)
        )

    last_attendance = (
        attendance[0]
        if attendance
        else "No Record"
    )

    content.append(
        Paragraph(
            f"Member ID: {member[1]}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Name: {member[2]}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Contact: {member[3]}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Birthday: {member[6]}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Email: {member[7]}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Address: {member[4]}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Total Payments: ₩{total_payment:,}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Last Attendance: {last_attendance}",
            styles["Normal"]
        )
    )

    content.append(
        Spacer(1,20)
    )

    content.append(
        Paragraph(
            "PAYMENT HISTORY",
            styles["Heading2"]
        )
    )

    for payment in payments:

        content.append(
            Paragraph(
                f"{payment[0]} | "
                f"{payment[1]} | "
                f"₩{payment[2]:,}",
                styles["Normal"]
            )
        )

    pdf.build(
        content
    )

    return send_file(
        filename,
        as_attachment=True
    )

@app.route("/dashboard_outstanding")
def dashboard_outstanding():

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    data = []

    cursor.execute("""
    SELECT
        member_id,
        full_name,
        member_since
    FROM members
    WHERE status='Active'
    ORDER BY full_name
    """)

    members = cursor.fetchall()

    current_date = datetime.now()

    for member in members:

        member_id = member[0]
        full_name = member[1]
        member_since = member[2]

        if not member_since:
            continue

        try:

            month_name, year = member_since.split()

            year = int(year)

            month_map = {
                "January":1,
                "February":2,
                "March":3,
                "April":4,
                "May":5,
                "June":6,
                "July":7,
                "August":8,
                "September":9,
                "October":10,
                "November":11,
                "December":12
            }

            month = month_map[month_name]

        except:
            continue

        missing_months = []

        while (
            year < current_date.year
            or
            (
                year == current_date.year
                and month <= current_date.month
            )
        ):

            current_month = datetime(
                year,
                month,
                1
            ).strftime("%B")

            cursor.execute("""
            SELECT COUNT(*)
            FROM payments
            WHERE member_id=%s
            AND payment_type='Monthly Contribution'
            AND payment_month=%s
            AND payment_year=%s
            """, (
                member_id,
                current_month,
                str(year)
            ))

            paid = cursor.fetchone()[0]

            if paid == 0:

                missing_months.append(
                    f"{current_month} {year}"
                )

            month += 1

            if month > 12:
                month = 1
                year += 1

        if missing_months:

            data.append({
                "member_id": member_id,
                "full_name": full_name,
                "months": missing_months,
                "amount": len(missing_months)*10000
            })

    conn.close()

    return render_template(
        "dashboard_outstanding.html",
        data=data,
        username=session["username"]
    )


@app.route("/dashboard_active")
def dashboard_active():


    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    active_members = []

    cursor.execute("""
    SELECT
        member_id,
        full_name,
        contact,
        member_since
    FROM members
    ORDER BY full_name
    """)

    members = cursor.fetchall()

    current_date = datetime.now()

    month_map = {
        "January":1,
        "February":2,
        "March":3,
        "April":4,
        "May":5,
        "June":6,
        "July":7,
        "August":8,
        "September":9,
        "October":10,
        "November":11,
        "December":12
    }

    for member in members:

        member_id = member[0]
        full_name = member[1]
        contact = member[2]
        member_since = member[3]

        if not member_since:
            continue

        try:

            month_name, year = member_since.split()

            year = int(year)

            month = month_map[month_name]

        except:
            continue

        missing_count = 0

        temp_year = year
        temp_month = month

        while (
            temp_year < current_date.year
            or
            (
                temp_year == current_date.year
                and temp_month <= current_date.month
            )
        ):

            current_month = datetime(
                temp_year,
                temp_month,
                1
            ).strftime("%B")

            cursor.execute("""
            SELECT COUNT(*)
            FROM payments
            WHERE member_id = %s
            AND payment_type='Monthly Contribution'
            AND payment_month = %s
            AND payment_year = %s
            """, (
                member_id,
                current_month,
                str(temp_year)
            ))

            paid = cursor.fetchone()[0]

            if paid == 0:
                missing_count += 1

            temp_month += 1

            if temp_month > 12:
                temp_month = 1
                temp_year += 1

        if missing_count < 5:

            active_members.append(
                (
                    member_id,
                    full_name,
                    contact,
                    missing_count
                )
            )

    conn.close()

    return render_template(
        "dashboard_active.html",
        members=active_members,
        username=session["username"]
    )




@app.route("/dashboard_inactive")
def dashboard_inactive():

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    inactive_members = []

    cursor.execute("""
    SELECT
        member_id,
        full_name,
        contact,
        member_since
    FROM members
    ORDER BY full_name
    """)

    members = cursor.fetchall()

    current_date = datetime.now()

    month_map = {
        "January":1,
        "February":2,
        "March":3,
        "April":4,
        "May":5,
        "June":6,
        "July":7,
        "August":8,
        "September":9,
        "October":10,
        "November":11,
        "December":12
    }

    for member in members:

        member_id = member[0]
        full_name = member[1]
        contact = member[2]
        member_since = member[3]

        if not member_since:
            continue

        try:

            month_name, year = member_since.split()

            year = int(year)

            month = month_map[month_name]

        except:
            continue

        missing_count = 0

        while (
            year < current_date.year
            or
            (
                year == current_date.year
                and month <= current_date.month
            )
        ):

            current_month = datetime(
                year,
                month,
                1
            ).strftime("%B")

            cursor.execute("""
            SELECT COUNT(*)
            FROM payments
            WHERE member_id = %s
            AND payment_type='Monthly Contribution'
            AND payment_month = %s
            AND payment_year = %s
            """, (
                member_id,
                current_month,
                str(year)
            ))

            paid = cursor.fetchone()[0]

            if paid == 0:
                missing_count += 1

            month += 1

            if month > 12:
                month = 1
                year += 1

        if missing_count >= 5:

            inactive_members.append(
                (
                    member_id,
                    full_name,
                    contact,
                    missing_count
                )
            )

    conn.close()

    return render_template(
        "dashboard_inactive.html",
        members=inactive_members,
        username=session["username"]
    )

@app.route("/dashboard_applicants")
def dashboard_applicants():


    if "username" not in session:
        return redirect("/login")
    
    conn = get_db()
    cursor = conn.cursor()

    # AUTO DELETE APPLICANTS AFTER 3 DAYS

    cursor.execute("""
    SELECT
        member_id,
        date_registered
    FROM members
    WHERE status='Applicant'
    """)

    applicant_list = cursor.fetchall()

    today = datetime.now()

    for applicant in applicant_list:

        try:

            registered_date = datetime.strptime(
                applicant[1],
                "%Y-%m-%d"
            )

            days_pending = (
                today - registered_date
            ).days
    
            if days_pending >= 3:

                cursor.execute("""
                SELECT COUNT(*)
                FROM payments
                WHERE member_id = %s
                AND payment_type='Registration Fee'
                """, (
                    applicant[0],
                ))

                has_payment = cursor.fetchone()[0]

                if has_payment == 0:

                    cursor.execute("""
                    DELETE FROM members
                    WHERE member_id = %s
                    """, (
                        applicant[0],
                    ))

        except:
            pass

    conn.commit()
    
    cursor.execute("""
    SELECT
        member_id,
        full_name,
        contact,
        date_registered
    FROM members
    WHERE status='Applicant'
    ORDER BY id DESC
    """)

    applicants = cursor.fetchall()

    conn.close()

    return render_template(
        "dashboard_applicants.html",
        applicants=applicants,
        username=session["username"]
    )

@app.route("/export_pdf_report")
def export_pdf_report():


    if "username" not in session:
        return redirect("/login")

    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer
    )

    from reportlab.lib.styles import (
        getSampleStyleSheet
    )

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute(
        "SELECT COUNT(*) FROM members"
    )
    total_members = cursor.fetchone()[0]

    cursor.execute("""
    SELECT COUNT(*)
    FROM members
    WHERE status='Active'
    """)
    active_members = cursor.fetchone()[0]

    cursor.execute("""
    SELECT COUNT(*)
    FROM members
    WHERE status='Applicant'
    """)
    applicants = cursor.fetchone()[0]

    cursor.execute("""
    SELECT COALESCE(SUM(amount),0)
    FROM payments
    """)
    collections = cursor.fetchone()[0]

    cursor.execute("""
    SELECT COALESCE(SUM(amount),0)
    FROM donations
    """)
    donations = cursor.fetchone()[0]

    cursor.execute("""
    SELECT COALESCE(SUM(amount),0)
    FROM expenses
    """)
    expenses = cursor.fetchone()[0]

    balance = (
        collections +
        donations -
        expenses
    )

    conn.close()

    os.makedirs(
        "exports",
        exist_ok=True
    )

    filename = (
        f"exports/FCCI_Report_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    )

    doc = SimpleDocTemplate(
        filename
    )

    styles = getSampleStyleSheet()

    content = []

    content.append(
        Paragraph(
            "FCCI FINANCIAL REPORT",
            styles["Title"]
        )
    )

    content.append(
        Spacer(1,20)
    )

    content.append(
        Paragraph(
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            styles["Normal"]
        )
    )

    content.append(
        Spacer(1,20)
    )

    content.append(
        Paragraph(
            f"Total Members: {total_members}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Active Members: {active_members}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Applicants: {applicants}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Collections: ₩{collections:,}",
            styles["Normal"]
        )
    )
    
    content.append(
        Paragraph(
            f"Donations: ₩{donations:,}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Expenses: ₩{expenses:,}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"Current Balance: ₩{balance:,}",
            styles["Normal"]
        )
    )

    doc.build(content)

    return send_file(
        filename,
        as_attachment=True
    )



@app.route("/approve_applicant/<member_id>")
def approve_applicant(member_id):

    if "username" not in session:
        return redirect("/login")

    from datetime import datetime

    conn = get_db()
    cursor = conn.cursor()

    # Kunin ang susunod na FCCI number
    cursor.execute("""
        SELECT member_id
        FROM members
        WHERE member_id LIKE 'FCCI-%'
        ORDER BY member_id DESC
        LIMIT 1
    """)

    last_member = cursor.fetchone()

    if last_member:
        last_num = int(last_member[0].split("-")[-1])
        next_num = last_num + 1
    else:
        next_num = 1

    year = datetime.now().year

    new_member_id = f"FCCI-{year}-{next_num:06d}"

    # Update applicant
    cursor.execute("""
        UPDATE members
        SET
            member_id = %s,
            status = 'Active',
            registration_fee = 20000,
            member_since = CURRENT_DATE
        WHERE member_id = %s
    """, (
        new_member_id,
        member_id
    ))

    # Generate receipt number
    cursor.execute("SELECT COUNT(*) FROM payments")
    payment_count = cursor.fetchone()[0] + 1

    receipt_no = f"RCPT-{year}-{payment_count:06d}"

    # Create registration payment record
    cursor.execute("""
        INSERT INTO payments
        (
            receipt_no,
            member_id,
            payment_type,
            amount,
            payment_date,
            payment_year,
            payment_month
        )
        VALUES
        (
            %s,
            %s,
            'Registration Fee',
            20000,
            CURRENT_DATE,
            %s,
            %s
        )
    """, (
        receipt_no,
        new_member_id,
        year,
        datetime.now().strftime("%B")
    ))

    conn.commit()
    conn.close()


    return redirect("/registration_approval")

@app.route(
"/reject_applicant/<member_id>"
)
def reject_applicant(member_id):


    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    DELETE FROM members
    WHERE member_id=%s
    """, (
        member_id,
    ))

    conn.commit()
    conn.close()

    return redirect(
        "/registration_approval"
    )

@app.route("/registration_confirmation/<member_id>")
def registration_confirmation(member_id):

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT member_id, full_name, proof_of_payment
    FROM members
    WHERE member_id = %s
    """, (member_id,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        return "Applicant Not Found"

    member = {
        "member_id": row[0],
        "full_name": row[1],
        "proof_uploaded": bool(row[2])
    }

    return render_template(
        "registration_confirmation.html",
        member=member
    )


@app.route("/upload_proof_of_payment", methods=["POST"])
def upload_proof_of_payment():

    member_id = request.form["member_id"]
    proof_file = request.files.get("proof_of_payment")

    if proof_file and proof_file.filename:

        proof_filename = upload_photo(proof_file, folder="fcci_proof_of_payment") or ""

        conn = get_db()
        cursor = conn.cursor()

        cursor.execute("""
        UPDATE members
        SET proof_of_payment = %s
        WHERE member_id = %s
        """, (proof_filename, member_id))

        conn.commit()
        conn.close()

    return redirect(f"/registration_confirmation/{member_id}")


@app.route("/applicant_slip/<member_id>")
def applicant_slip(member_id):

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT
    id, member_id, full_name, contact, address, registration_fee, member_since, email, birthday, date_registered, status, proof_of_payment
    FROM members
    WHERE member_id = %s
    """, (
        member_id,
    ))

    member = cursor.fetchone()

    conn.close()

    if not member:
        return "Applicant Not Found"

    import os
    import qrcode

    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Image,
        Table
    )

    from reportlab.lib.styles import (
        getSampleStyleSheet
    )

    os.makedirs(
        "exports",
        exist_ok=True
    )

    os.makedirs(
        "static/qr",
        exist_ok=True
    )

    qr_file = f"static/qr/{member_id}.png"

    qr = qrcode.make(member_id)
    qr.save(qr_file)

    filename = (
        f"exports/Applicant_{member_id}.pdf"
    )

    pdf = SimpleDocTemplate(filename)

    styles = getSampleStyleSheet()

    content = []

    logo_path = "static/fcci_logo.jpeg"

    if os.path.exists(logo_path):

        content.append(
            Image(
                logo_path,
                width=80,
                height=80
            )
        )

    content.append(
        Paragraph(
            "FILIPINO COMMUNITY CENTER INTERNATIONAL",
            styles["Title"]
        )
    )

    content.append(
        Paragraph(
            "APPLICANT REGISTRATION SLIP",
            styles["Heading1"]
        )
    )

    content.append(
        Spacer(1,20)
    )

    # Kunin ang applicant photo mula sa member_photos table (Cloudinary)
    slip_photo_path = download_photo_for_pdf(member_id)

    if slip_photo_path:
        content.append(Image(slip_photo_path, width=120, height=120))
        content.append(Spacer(1, 10))

    content.append(
        Image(
            qr_file,
            width=120,
            height=120
        )
    )

    content.append(
        Spacer(1,20)
    )

    data = [

        ["Applicant ID", member[1]],
        ["Full Name", member[2]],
        ["Contact", member[3]],
        ["Email", member[7]],
        ["Birthday", member[8]],
        ["Date Registered", member[9]],
        ["Status", member[10]]

    ]

    table = Table(
        data,
        colWidths=[180,300]
    )

    content.append(table)

    content.append(
        Spacer(1,30)
    )

    content.append(
        Paragraph(
            "Please pay the Registration Fee of ₩20,000 to become an Official FCCI Member.",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            "Present this Applicant Slip during payment.",
            styles["Normal"]
        )
    )

    pdf.build(content)

    return send_file(
        filename,
        as_attachment=True
    )

@app.route("/registration_approval")
def registration_approval():

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    applicants = fetch_all_members_with_photo(
        cursor,
        where_clause="status='Applicant'"
    )

    conn.close()

    return render_template(
        "registration_approval.html",
        applicants=applicants
    )

@app.route("/get_member_info/<member_id>")
def get_member_info(member_id):
 
    if "username" not in session:
        return jsonify({"found": False, "error": "Not logged in"}), 401
 
    conn = get_db()
    cursor = conn.cursor()
 
    cursor.execute("""
    SELECT member_id, full_name, status
    FROM members
    WHERE member_id = %s
    """, (member_id,))
 
    member = cursor.fetchone()
    conn.close()
 
    if not member:
        return jsonify({"found": False})
 
    return jsonify({
        "found": True,
        "member_id": member[0],
        "full_name": member[1],
        "status": member[2]
    })


# ============================================================
# WIRELESS PHONE SCANNER PAIRING SYSTEM
# Ginagamit para makapag-scan ng QR gamit ang isang phone, at
# automatic na lalabas ang resulta sa ibang device (hal. iPad
# o PC) na naka-display sa Payments/Attendance/Members page.
# ============================================================

import random
import string


def generate_pair_code():
    return "".join(random.choices(string.digits, k=6))


@app.route("/create_pair_session", methods=["POST"])
def create_pair_session():

    if "username" not in session:
        return jsonify({"error": "Not logged in"}), 401

    target_page = request.form.get("target_page", "payments")

    conn = get_db()
    cursor = conn.cursor()

    # Gumawa ng unique 6-digit code (subukan ulit kung sakaling
    # may existing na pareho)
    pair_code = generate_pair_code()

    for _ in range(5):
        cursor.execute(
            "SELECT id FROM pairing_sessions WHERE pair_code = %s",
            (pair_code,)
        )
        if not cursor.fetchone():
            break
        pair_code = generate_pair_code()

    cursor.execute("""
    INSERT INTO pairing_sessions (pair_code, target_page, scanned_value)
    VALUES (%s, %s, NULL)
    """, (pair_code, target_page))

    conn.commit()
    conn.close()

    return jsonify({"pair_code": pair_code})


@app.route("/mobile_scan/<pair_code>")
def mobile_scan(pair_code):

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute(
        "SELECT id, target_page FROM pairing_sessions WHERE pair_code = %s",
        (pair_code,)
    )
    session_row = cursor.fetchone()
    conn.close()

    if not session_row:
        return render_template(
            "mobile_scan.html",
            valid_session=False,
            pair_code=pair_code
        )

    return render_template(
        "mobile_scan.html",
        valid_session=True,
        pair_code=pair_code,
        target_page=session_row[1]
    )


@app.route("/submit_mobile_scan", methods=["POST"])
def submit_mobile_scan():

    pair_code = request.form.get("pair_code", "")
    scanned_value = request.form.get("scanned_value", "")

    if not pair_code or not scanned_value:
        return jsonify({"success": False, "error": "Missing data"})

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    UPDATE pairing_sessions
    SET scanned_value = %s, updated_at = NOW()
    WHERE pair_code = %s
    """, (scanned_value, pair_code))

    conn.commit()
    conn.close()

    return jsonify({"success": True})


@app.route("/check_pair_scan/<pair_code>")
def check_pair_scan(pair_code):

    if "username" not in session:
        return jsonify({"error": "Not logged in"}), 401

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT scanned_value FROM pairing_sessions WHERE pair_code = %s
    """, (pair_code,))

    row = cursor.fetchone()

    if row and row[0]:
        # I-clear agad pagkatapos mabasa, para hindi paulit-ulit
        # mag-trigger ang parehong scan
        cursor.execute("""
        UPDATE pairing_sessions SET scanned_value = NULL WHERE pair_code = %s
        """, (pair_code,))
        conn.commit()
        conn.close()
        return jsonify({"has_scan": True, "value": row[0]})

    conn.close()
    return jsonify({"has_scan": False})


@app.route("/settings")
def settings():
 
    if "username" not in session:
        return redirect("/login")
 
    conn = get_db()
    cursor = conn.cursor()
 
    cursor.execute("SELECT username, role FROM users ORDER BY username")
    all_users = cursor.fetchall()
 
    conn.close()
 
    return render_template(
        "settings.html",
        username=session["username"],
        all_users=all_users
    )
 
 
@app.route("/settings/change_username", methods=["POST"])
def change_username():
 
    if "username" not in session:
        return redirect("/login")
 
    new_username = request.form["new_username"].strip()
    current_password = request.form["current_password"]
 
    conn = get_db()
    cursor = conn.cursor()
 
    # Kunin ang current user's stored password
    cursor.execute(
        "SELECT password FROM users WHERE username = %s",
        (session["username"],)
    )
    row = cursor.fetchone()
 
    if not row:
        conn.close()
        return redirect("/login")
 
    stored_password = row[0]
 
    # Verify current password (supports hashed or plain)
    if stored_password.startswith("pbkdf2:") or stored_password.startswith("scrypt:"):
        password_ok = check_password_hash(stored_password, current_password)
    else:
        password_ok = (stored_password == current_password)
 
    if not password_ok:
        conn.close()
        return render_template(
            "settings.html",
            username=session["username"],
            all_users=[],
            error="Maling current password. Subukan ulit."
        )
 
    # Check kung taken na ang bagong username
    cursor.execute(
        "SELECT COUNT(*) FROM users WHERE username = %s",
        (new_username,)
    )
    taken = cursor.fetchone()[0]
 
    if taken > 0 and new_username != session["username"]:
        conn.close()
        return render_template(
            "settings.html",
            username=session["username"],
            all_users=[],
            error=f'Ang username "{new_username}" ay ginagamit na.'
        )
 
    # I-update ang username
    cursor.execute(
        "UPDATE users SET username = %s WHERE username = %s",
        (new_username, session["username"])
    )
    conn.commit()
    conn.close()
 
    session["username"] = new_username
 
    cursor = get_db().cursor()
 
    return redirect("/settings")
 
 
@app.route("/settings/change_password", methods=["POST"])
def change_password():
 
    if "username" not in session:
        return redirect("/login")
 
    current_password = request.form["current_password"]
    new_password = request.form["new_password"]
    confirm_password = request.form["confirm_password"]
 
    conn = get_db()
    cursor = conn.cursor()
 
    cursor.execute(
        "SELECT password FROM users WHERE username = %s",
        (session["username"],)
    )
    row = cursor.fetchone()
 
    if not row:
        conn.close()
        return redirect("/login")
 
    stored_password = row[0]
 
    if stored_password.startswith("pbkdf2:") or stored_password.startswith("scrypt:"):
        password_ok = check_password_hash(stored_password, current_password)
    else:
        password_ok = (stored_password == current_password)
 
    if not password_ok:
        conn.close()
        cursor2 = get_db().cursor()
        cursor2.execute("SELECT username, role FROM users ORDER BY username")
        all_users = cursor2.fetchall()
        return render_template(
            "settings.html",
            username=session["username"],
            all_users=all_users,
            error="Maling current password."
        )
 
    if new_password != confirm_password:
        conn.close()
        cursor2 = get_db().cursor()
        cursor2.execute("SELECT username, role FROM users ORDER BY username")
        all_users = cursor2.fetchall()
        return render_template(
            "settings.html",
            username=session["username"],
            all_users=all_users,
            error="Hindi tugma ang New Password at Confirm Password."
        )
 
    hashed_password = generate_password_hash(new_password)
 
    cursor.execute(
        "UPDATE users SET password = %s WHERE username = %s",
        (hashed_password, session["username"])
    )
    conn.commit()
    conn.close()
 
    cursor2 = get_db().cursor()
    cursor2.execute("SELECT username, role FROM users ORDER BY username")
    all_users = cursor2.fetchall()
 
    return render_template(
        "settings.html",
        username=session["username"],
        all_users=all_users,
        message="Matagumpay na na-update ang password!"
    )
 
 
@app.route("/settings/add_user", methods=["POST"])
def add_user():
 
    if "username" not in session:
        return redirect("/login")
 
    new_username = request.form["username"].strip()
    new_password = request.form["password"]
    role = request.form["role"]
 
    conn = get_db()
    cursor = conn.cursor()
 
    cursor.execute(
        "SELECT COUNT(*) FROM users WHERE username = %s",
        (new_username,)
    )
    taken = cursor.fetchone()[0]
 
    if taken > 0:
        conn.close()
        cursor2 = get_db().cursor()
        cursor2.execute("SELECT username, role FROM users ORDER BY username")
        all_users = cursor2.fetchall()
        return render_template(
            "settings.html",
            username=session["username"],
            all_users=all_users,
            error=f'Ang username "{new_username}" ay ginagamit na.'
        )
 
    hashed_password = generate_password_hash(new_password)
 
    cursor.execute(
        "INSERT INTO users (username, password, role) VALUES (%s, %s, %s)",
        (new_username, hashed_password, role)
    )
    conn.commit()
    conn.close()
 
    cursor2 = get_db().cursor()
    cursor2.execute("SELECT username, role FROM users ORDER BY username")
    all_users = cursor2.fetchall()
 
    return render_template(
        "settings.html",
        username=session["username"],
        all_users=all_users,
        message=f'Matagumpay na nagawa ang user "{new_username}"!'
    )

@app.route("/delete_donation/<int:donation_id>")
def delete_donation(donation_id):
 
    if "username" not in session:
        return redirect("/login")
 
    conn = get_db()
    cursor = conn.cursor()
 
    cursor.execute(
        "DELETE FROM donations WHERE id = %s",
        (donation_id,)
    )
 
    conn.commit()
    conn.close()
 
    return redirect("/donations")

@app.route("/settings/backup_database")
def backup_database():

    if "username" not in session:
        return redirect("/login")

    return render_template(
        "settings.html",
        username=session["username"],
        all_users=get_all_users(),
        message="Ang database ay naka-Supabase na (cloud). Para mag-backup, pumunta sa iyong Supabase dashboard → Project → Backups."
    )

@app.route("/settings/restore_database", methods=["POST"])
def restore_database():

    if "username" not in session:
        return redirect("/login")

    return render_template(
        "settings.html",
        username=session["username"],
        all_users=get_all_users(),
        message="Ang database ay naka-Supabase na (cloud). Ang Restore from .db file ay hindi na available. Para mag-restore, gamitin ang Supabase dashboard → Project → Backups."
    )

@app.route("/settings/merge_database", methods=["POST"])
def merge_database():

    if "username" not in session:
        return redirect("/login")

    return render_template(
        "settings.html",
        username=session["username"],
        all_users=get_all_users(),
        message="Ang database ay naka-Supabase na (cloud). Ang Merge from .db file ay hindi na available."
    )

def merge_old_database(old_db_path):
    """
    Kinukuha ang lahat ng members, payments, donations, at expenses
    mula sa lumang offline .db file at idinadagdag (INSERT) sila sa
    kasalukuyang database, hindi pinapalitan ang existing records.

    SAFETY: Gumagamit lang ng mga column na PARESHAS sa dalawang
    database (matching column names). Kung may column sa luma na
    wala sa bago (o vice versa), hindi ito sasama sa insert at
    hindi magiging error — para hindi masira ang merge kung may
    konting pagkaiba sa schema.

    Member duplicates ay kinakheck gamit ang member_id (skip kung
    existing na). Payments/donations/expenses ay direktang idinadagdag
    dahil walang natural duplicate key dito.
    """

    old_conn = sqlite3.connect(old_db_path)
    old_cursor = old_conn.cursor()

    new_conn = get_db()
    new_cursor = new_conn.cursor()

    summary_parts = []

    def get_columns(cursor, table_name):
        cursor.execute(f"PRAGMA table_info({table_name})")
        return [col[1] for col in cursor.fetchall()]

    def table_exists(cursor, table_name):
        cursor.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=%s",
            (table_name,)
        )
        return cursor.fetchone() is not None

    # ── MEMBERS (may duplicate check gamit ang member_id) ──
    if table_exists(old_cursor, "members") and table_exists(new_cursor, "members"):

        old_columns = get_columns(old_cursor, "members")
        new_columns = get_columns(new_cursor, "members")

        # Gamitin lang ang columns na PARESHAS sa dalawa, laktawan ang 'id'
        common_columns = [c for c in old_columns if c in new_columns and c != "id"]

        if "member_id" not in common_columns:
            summary_parts.append("Members: hindi na-merge — walang 'member_id' column na pareho sa dalawang database")
        else:
            old_cursor.execute(f"SELECT {', '.join(common_columns)} FROM members")
            old_members = old_cursor.fetchall()

            member_id_index = common_columns.index("member_id")

            added_members = 0
            skipped_members = 0

            for row in old_members:
                member_id = row[member_id_index]

                new_cursor.execute(
                    "SELECT COUNT(*) FROM members WHERE member_id = %s",
                    (member_id,)
                )
                exists = new_cursor.fetchone()[0]

                if exists:
                    skipped_members += 1
                    continue

                placeholders = ", ".join(["%s"] * len(common_columns))
                columns_str = ", ".join(common_columns)

                new_cursor.execute(
                    f"INSERT INTO members ({columns_str}) VALUES ({placeholders})",
                    row
                )

                added_members += 1

            summary_parts.append(f"Members: {added_members} idinagdag, {skipped_members} na-skip (duplicate)")
    else:
        summary_parts.append("Members: walang nahanap na table")

    # ── MEMBER_PHOTOS (hiwalay na table, gamit ng offline version) ──
    if table_exists(old_cursor, "member_photos") and table_exists(new_cursor, "member_photos"):

        old_columns = get_columns(old_cursor, "member_photos")
        new_columns = get_columns(new_cursor, "member_photos")
        common_columns = [c for c in old_columns if c in new_columns and c != "id"]

        if common_columns and "member_id" in common_columns:
            old_cursor.execute(f"SELECT {', '.join(common_columns)} FROM member_photos")
            old_photos = old_cursor.fetchall()

            member_id_index = common_columns.index("member_id")

            placeholders = ", ".join(["%s"] * len(common_columns))
            columns_str = ", ".join(common_columns)

            added_photos = 0

            for row in old_photos:
                member_id = row[member_id_index]

                # Skip kung may existing na photo record na ang member na ito
                new_cursor.execute(
                    "SELECT COUNT(*) FROM member_photos WHERE member_id = %s",
                    (member_id,)
                )
                exists = new_cursor.fetchone()[0]

                if exists:
                    continue

                new_cursor.execute(
                    f"INSERT INTO member_photos ({columns_str}) VALUES ({placeholders})",
                    row
                )
                added_photos += 1

            summary_parts.append(f"Member Photos: {added_photos} idinagdag")
        else:
            summary_parts.append("Member Photos: walang pareho na columns")
    else:
        summary_parts.append("Member Photos: walang nahanap na table (okay lang, optional)")

    # ── PAYMENTS (walang duplicate check, direktang idagdag) ──
    if table_exists(old_cursor, "payments") and table_exists(new_cursor, "payments"):

        old_columns = get_columns(old_cursor, "payments")
        new_columns = get_columns(new_cursor, "payments")
        common_columns = [c for c in old_columns if c in new_columns and c != "id"]

        if common_columns:
            old_cursor.execute(f"SELECT {', '.join(common_columns)} FROM payments")
            old_payments = old_cursor.fetchall()

            placeholders = ", ".join(["%s"] * len(common_columns))
            columns_str = ", ".join(common_columns)

            for row in old_payments:
                new_cursor.execute(
                    f"INSERT INTO payments ({columns_str}) VALUES ({placeholders})",
                    row
                )

            summary_parts.append(f"Payments: {len(old_payments)} idinagdag")
        else:
            summary_parts.append("Payments: walang pareho na columns")
    else:
        summary_parts.append("Payments: walang nahanap na table")

    # ── DONATIONS (walang duplicate check, direktang idagdag) ──
    if table_exists(old_cursor, "donations") and table_exists(new_cursor, "donations"):

        old_columns = get_columns(old_cursor, "donations")
        new_columns = get_columns(new_cursor, "donations")
        common_columns = [c for c in old_columns if c in new_columns and c != "id"]

        if common_columns:
            old_cursor.execute(f"SELECT {', '.join(common_columns)} FROM donations")
            old_donations = old_cursor.fetchall()

            placeholders = ", ".join(["%s"] * len(common_columns))
            columns_str = ", ".join(common_columns)

            for row in old_donations:
                new_cursor.execute(
                    f"INSERT INTO donations ({columns_str}) VALUES ({placeholders})",
                    row
                )

            summary_parts.append(f"Donations: {len(old_donations)} idinagdag")
        else:
            summary_parts.append("Donations: walang pareho na columns")
    else:
        summary_parts.append("Donations: walang nahanap na table")

    # ── EXPENSES (walang duplicate check, direktang idagdag) ──
    if table_exists(old_cursor, "expenses") and table_exists(new_cursor, "expenses"):

        old_columns = get_columns(old_cursor, "expenses")
        new_columns = get_columns(new_cursor, "expenses")
        common_columns = [c for c in old_columns if c in new_columns and c != "id"]

        if common_columns:
            old_cursor.execute(f"SELECT {', '.join(common_columns)} FROM expenses")
            old_expenses = old_cursor.fetchall()

            placeholders = ", ".join(["%s"] * len(common_columns))
            columns_str = ", ".join(common_columns)

            for row in old_expenses:
                new_cursor.execute(
                    f"INSERT INTO expenses ({columns_str}) VALUES ({placeholders})",
                    row
                )

            summary_parts.append(f"Expenses: {len(old_expenses)} idinagdag")
        else:
            summary_parts.append("Expenses: walang pareho na columns")
    else:
        summary_parts.append("Expenses: walang nahanap na table")

    # ── ATTENDANCE (walang duplicate check, direktang idagdag) ──
    if table_exists(old_cursor, "attendance") and table_exists(new_cursor, "attendance"):

        old_columns = get_columns(old_cursor, "attendance")
        new_columns = get_columns(new_cursor, "attendance")
        common_columns = [c for c in old_columns if c in new_columns and c != "id"]

        if common_columns:
            old_cursor.execute(f"SELECT {', '.join(common_columns)} FROM attendance")
            old_attendance = old_cursor.fetchall()

            placeholders = ", ".join(["%s"] * len(common_columns))
            columns_str = ", ".join(common_columns)

            for row in old_attendance:
                new_cursor.execute(
                    f"INSERT INTO attendance ({columns_str}) VALUES ({placeholders})",
                    row
                )

            summary_parts.append(f"Attendance: {len(old_attendance)} idinagdag")
    else:
        summary_parts.append("Attendance: walang nahanap na table (okay lang, optional)")

    new_conn.commit()
    new_conn.close()
    old_conn.close()

    return "Matagumpay na na-merge ang backup! " + " | ".join(summary_parts)


def get_all_users():
    """Helper function para makuha ang listahan ng users (ginagamit sa settings page)."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT username, role FROM users ORDER BY username")
    users = cursor.fetchall()
    conn.close()
    return users

@app.route("/member_login", methods=["GET", "POST"])
def member_login():

    if request.method == "POST":

        member_id = request.form["member_id"].strip()
        birthday_input = request.form["birthday"]  # format: YYYY-MM-DD mula sa <input type="date">

        conn = get_db()
        cursor = conn.cursor()

        cursor.execute("""
        SELECT member_id, full_name, birthday, status
        FROM members
        WHERE member_id = %s
        """, (member_id,))

        member = cursor.fetchone()
        conn.close()

        if not member:
            return render_template(
                "member_login.html",
                error="Member ID na hindi natagpuan. Pakitiyak na tama ang inilagay mo."
            )

        stored_birthday = member[2]
        status = member[3]

        # I-normalize ang dalawang birthday format papuntang YYYY-MM-DD bago icompare
        normalized_stored = normalize_birthday(stored_birthday)

        if normalized_stored != birthday_input:
            return render_template(
                "member_login.html",
                error="Maling Member ID o Birthday. Subukan ulit."
            )

        if status != "Active":
            return render_template(
                "member_login.html",
                error=f"Hindi mo pa magagamit ang portal na ito. Ang status mo ay '{status}'. Pakibayaran muna ang Registration Fee sa FCCI office para ma-activate ang account mo."
            )

        # Successful login
        session["member_logged_in"] = True
        session["member_id"] = member[0]
        session["member_name"] = member[1]

        return redirect("/feed")

    return render_template("member_login.html")


def normalize_birthday(raw_value):
    """
    Sinusubukan i-convert ang anumang stored birthday format papunta sa
    'YYYY-MM-DD' para ma-compare sa value mula sa HTML date input.
    """
    if not raw_value:
        return ""

    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%B %d, %Y", "%b %d, %Y"):
        try:
            parsed = datetime.strptime(raw_value, fmt)
            return parsed.strftime("%Y-%m-%d")
        except ValueError:
            continue

    return raw_value  # fallback: ibalik na lang as-is


# ── MEMBER LOGOUT ──────────────────────────────────────────
@app.route("/member_logout")
def member_logout():
    session.pop("member_logged_in", None)
    session.pop("member_id", None)
    session.pop("member_name", None)
    return redirect("/member_login")


# ── DECORATOR/HELPER: I-CHECK KUNG NAKA-LOGIN ANG MEMBER ───
def require_member_login():
    """Tawagin ito sa simula ng bawat member-portal route."""
    if not session.get("member_logged_in"):
        return False
    return True

@app.route("/feed", methods=["GET", "POST"])
def feed():

    if not session.get("member_logged_in"):
        return redirect("/member_login")

    conn = get_db()
    cursor = conn.cursor()

    if request.method == "POST":

        content = request.form.get("content", "").strip()
        photo_file = request.files.get("photo")

        photo_filename = None

        if photo_file and photo_file.filename != "":
            photo_filename = upload_photo(photo_file, folder="fcci_feed") or None

        if content or photo_filename:
            now = datetime.now()

            cursor.execute("""
            INSERT INTO feed_posts
            (member_id, full_name, content, photo_path, is_pinned, post_date, post_time)
            VALUES (%s, %s, %s, %s, FALSE, %s, %s)
            """, (
                session["member_id"],
                session["member_name"],
                content,
                photo_filename,
                now.strftime("%B %d, %Y"),
                now.strftime("%I:%M %p")
            ))

            conn.commit()

        conn.close()
        return redirect("/feed")

    # GET request: kunin lahat ng posts, pinned muna, tapos pinaka-bago
    cursor.execute("""
    SELECT id, member_id, full_name, content, photo_path, is_pinned, post_date, post_time
    FROM feed_posts
    ORDER BY is_pinned DESC, id DESC
    """)
    posts = cursor.fetchall()

    # Bilangin ang likes at comments bawat post
    posts_with_meta = []

    for post in posts:
        post_id = post[0]

        cursor.execute("SELECT COUNT(*) FROM feed_likes WHERE post_id = %s", (post_id,))
        like_count = cursor.fetchone()[0]

        cursor.execute("""
        SELECT COUNT(*) FROM feed_likes WHERE post_id = %s AND member_id = %s
        """, (post_id, session["member_id"]))
        liked_by_me = cursor.fetchone()[0] > 0

        cursor.execute("""
        SELECT full_name, comment_text, comment_date, comment_time
        FROM feed_comments WHERE post_id = %s ORDER BY id ASC
        """, (post_id,))
        comments = cursor.fetchall()

        posts_with_meta.append({
            "id": post[0],
            "member_id": post[1],
            "full_name": post[2],
            "content": post[3],
            "photo_path": post[4],
            "is_pinned": post[5],
            "post_date": post[6],
            "post_time": post[7],
            "like_count": like_count,
            "liked_by_me": liked_by_me,
            "comments": comments
        })

    conn.close()

    return render_template(
        "feed.html",
        posts=posts_with_meta,
        member_name=session["member_name"],
        member_id=session["member_id"]
    )


# ── DELETE POST (sariling post lang) ───────────────────────
@app.route("/feed/delete/<int:post_id>")
def feed_delete_post(post_id):

    if not session.get("member_logged_in"):
        return redirect("/member_login")

    conn = get_db()
    cursor = conn.cursor()

    # Siguraduhing sariling post lang ng member ang puwedeng i-delete
    cursor.execute(
        "DELETE FROM feed_posts WHERE id = %s AND member_id = %s",
        (post_id, session["member_id"])
    )

    conn.commit()
    conn.close()

    return redirect("/feed")


# ── LIKE / UNLIKE POST ──────────────────────────────────────
@app.route("/feed/like/<int:post_id>")
def feed_like_post(post_id):

    if not session.get("member_logged_in"):
        return redirect("/member_login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute(
        "SELECT id FROM feed_likes WHERE post_id = %s AND member_id = %s",
        (post_id, session["member_id"])
    )
    existing = cursor.fetchone()

    if existing:
        cursor.execute("DELETE FROM feed_likes WHERE id = %s", (existing[False],))
    else:
        cursor.execute(
            "INSERT INTO feed_likes (post_id, member_id) VALUES (%s, %s)",
            (post_id, session["member_id"])
        )

    conn.commit()
    conn.close()

    return redirect("/feed")


# ── ADD COMMENT ──────────────────────────────────────────────
@app.route("/feed/comment/<int:post_id>", methods=["POST"])
def feed_add_comment(post_id):

    if not session.get("member_logged_in"):
        return redirect("/member_login")

    comment_text = request.form.get("comment_text", "").strip()

    if comment_text:
        conn = get_db()
        cursor = conn.cursor()

        now = datetime.now()

        cursor.execute("""
        INSERT INTO feed_comments (post_id, member_id, full_name, comment_text, comment_date, comment_time)
        VALUES (%s, %s, %s, %s, %s, %s)
        """, (
            post_id,
            session["member_id"],
            session["member_name"],
            comment_text,
            now.strftime("%B %d, %Y"),
            now.strftime("%I:%M %p")
        ))

        conn.commit()
        conn.close()

    return redirect("/feed")


# ── PIN / UNPIN POST (ADMIN/STAFF LANG, gamit ang main session) ──
@app.route("/feed/pin/<int:post_id>")
def feed_pin_post(post_id):

    # Gamit ang ADMIN session (yung "username" sa session, hindi member portal session)
    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT is_pinned FROM feed_posts WHERE id = %s", (post_id,))
    row = cursor.fetchone()

    if row:
        new_value = False if row[False] == True else True
        cursor.execute("UPDATE feed_posts SET is_pinned = %s WHERE id = %s", (new_value, post_id))
        conn.commit()

    conn.close()

    # Ibalik sa pinanggalingang page (admin feed view o member feed)
    return redirect(request.referrer or "/feed")


# ── MEMBER'S OWN PROFILE VIEW (sa loob ng portal) ──────────
@app.route("/member_portal_profile")
def member_portal_profile():

    if not session.get("member_logged_in"):
        return redirect("/member_login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT id, member_id, full_name, contact, address, registration_fee, member_since, email, birthday, date_registered, status, proof_of_payment FROM members WHERE member_id = %s", (session["member_id"],))
    member = cursor.fetchone()

    cursor.execute("""
    SELECT payment_date, payment_type, amount
    FROM payments WHERE member_id = %s ORDER BY id DESC
    """, (session["member_id"],))
    payments = cursor.fetchall()

    cursor.execute("""
    SELECT COALESCE(SUM(amount), 0) FROM payments WHERE member_id = %s
    """, (session["member_id"],))
    total_payment = cursor.fetchone()[0]

    cursor.execute("""
    SELECT photo_path FROM member_photos WHERE member_id = %s ORDER BY id DESC LIMIT 1
    """, (session["member_id"],))
    photo_row = cursor.fetchone()
    photo_path = photo_row[0] if photo_row else None

    conn.close()

    return render_template(
        "member_portal_profile.html",
        member=member,
        payments=payments,
        total_payment=total_payment,
        photo_path=photo_path
    )

@app.route("/admin_feed")
def admin_feed():

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT id, member_id, full_name, content, photo_path, is_pinned, post_date, post_time
    FROM feed_posts
    ORDER BY is_pinned DESC, id DESC
    """)
    posts = cursor.fetchall()

    conn.close()

    return render_template("admin_feed.html", posts=posts, username=session["username"])



@app.route("/logout")
def logout():

    session.clear()

    return redirect("/login")



@app.route("/admin_feed/post", methods=["POST"])
def admin_feed_post():
    if "username" not in session:
        return redirect("/login")
    content_text = request.form.get("content", "").strip()
    photo_file = request.files.get("photo")
    photo_filename = None
    if photo_file and photo_file.filename != "":
        photo_filename = upload_photo(photo_file, folder="fcci_feed")
    if content_text or photo_filename:
        conn = get_db()
        cursor = conn.cursor()
        now = datetime.now()
        cursor.execute("""
        INSERT INTO feed_posts
        (member_id, full_name, content, photo_path, is_pinned, post_date, post_time)
        VALUES (%s, %s, %s, %s, FALSE, %s, %s)
        """, (
            "ADMIN",
            session["username"],
            content_text,
            photo_filename,
            now.strftime("%B %d, %Y"),
            now.strftime("%I:%M %p")
        ))
        conn.commit()
        conn.close()
    return redirect("/admin_feed")


@app.route("/admin_feed/delete/<int:post_id>")
def admin_feed_delete(post_id):
    if "username" not in session:
        return redirect("/login")
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM feed_posts WHERE id = %s", (post_id,))
    conn.commit()
    conn.close()
    return redirect("/admin_feed")

if __name__ == "__main__":

    # ── TEST SUPABASE CONNECTION BAGO MAGSIMULA ──────────────
    try:
        test_conn = get_db()
        test_conn.close()
        print("[SUPABASE] Matagumpay na nakakonekta sa Supabase database!")
    except Exception as conn_error:
        print(f"[SUPABASE] WARNING: Hindi makakonekta sa Supabase: {conn_error}")
        print("[SUPABASE] Siguraduhing tama ang DATABASE_URL sa .env file mo.")

    app.run(
        debug=True
    )
