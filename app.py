from flask import (
    Flask,
    render_template,
    request,
    redirect,
    session,
    send_file
)

from werkzeug.utils import secure_filename
import os
from openpyxl import Workbook
import sqlite3
import os
import cv2
import qrcode
from datetime import datetime

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

app = Flask(__name__)

app.secret_key = "FCCI_SECRET_KEY"

UPLOAD_FOLDER = "static/uploads"

app.config[
    "UPLOAD_FOLDER"
] = UPLOAD_FOLDER

DB_PATH = os.path.join(
    "database",
    "fcci.db"
)


def get_db():
    return sqlite3.connect(DB_PATH)


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
        SELECT role
        FROM users
        WHERE username = ?
        AND password = ?
        """, (
            username,
            password
        ))

        user = cursor.fetchone()

        conn.close()

        if user:

            session["username"] = username
            session["role"] = user[0]

            return redirect("/dashboard")

        return render_template(
            "login.html",
            error="Invalid Login"
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

        cursor.execute("""
        SELECT COUNT(*)
        FROM members
        """)

        count = cursor.fetchone()[0] + 1

        member_id = (
            f"APP-{datetime.now().year}-{count:06d}"
        )

        full_name = request.form["full_name"]
        contact = request.form["contact"]
        birthday = request.form["birthday"]
        email = request.form["email"]
        address = request.form["address"]

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
        (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            ""
        ))

        conn.commit()
        conn.close()

        return redirect(
            f"/applicant_slip/{member_id}"
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

    cursor.execute("""
    SELECT COUNT(*)
    FROM members
    """)

    result = cursor.fetchone()
    total_members = result[0] if result else 0

    # COLLECTIONS

    cursor.execute("""
    SELECT COALESCE(
        SUM(amount),
        0
    )
    FROM payments
    """)

    result = cursor.fetchone()
    collections = result[0] if result else 0

    # DONATIONS

    cursor.execute("""
    SELECT COALESCE(
        SUM(amount),
        0
    )
    FROM donations
    """)

    result = cursor.fetchone()
    donations = result[0] if result else 0

    # EXPENSES

    cursor.execute("""
    SELECT COALESCE(
        SUM(amount),
        0
    )
    FROM expenses
    """)

    result = cursor.fetchone()
    expenses = result[0] if result else 0

    # CURRENT BALANCE

    balance = (
        collections
        + donations
        - expenses
    )

    # RECENT PAYMENTS

    cursor.execute("""
    SELECT
        member_id,
        amount
    FROM payments
    ORDER BY id DESC
    LIMIT 10
    """)

    recent_payments = cursor.fetchall()

    # ACTIVE / INACTIVE / OUTSTANDING

    active_members = 0
    inactive_members = 0
    total_outstanding = 0

    cursor.execute("""
        SELECT
            member_id,
            member_since
        FROM members
        WHERE status = 'Active'
        """)

    members = cursor.fetchall()

    # APPLICANTS COUNT

    cursor.execute("""
    SELECT COUNT(*)
    FROM members
    WHERE status='Applicant'
    """)

    result = cursor.fetchone()

    applicants = result[0] if result else 0

    current_date = datetime.now()

    for member in members:

        member_id = member[0]
        member_since = member[1]

        if not member_since:
            continue

        try:

            parts = member_since.split()

            month_name = parts[0]
            year = int(parts[1])

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

        missing_count = 0



        while (
            year < current_date.year
            or
            (
                year == current_date.year
                and month <= current_date.month
            )
        ):

            month_name = datetime(
                year,
                month,
                1
            ).strftime("%B")

            cursor.execute("""
            SELECT COUNT(*)
            FROM payments
            WHERE member_id = ?
            AND payment_type = 'Monthly Contribution'
            AND payment_month = ?
            AND payment_year = ?
            """, (
                member_id,
                month_name,
                str(year)
            ))

            result = cursor.fetchone()

            paid = (
                result[0]
                if result else 0
            )

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

        total_outstanding += (
            missing_count * 10000
        )

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

        birthday_list=[],

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
            status
        FROM members
        WHERE
            member_id LIKE ?
            OR full_name LIKE ?
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
            status
        FROM members
        ORDER BY full_name
        """)

    members = cursor.fetchall()

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

        cursor.execute("""
        SELECT member_id
        FROM members
        ORDER BY id DESC
        LIMIT 1
        """)

        last = cursor.fetchone()

        if last:

            try:

                last_no = int(
                    last[0].split("-")[-1]
                )

                next_no = last_no + 1

            except:

                next_no = 1

        else:

            next_no = 1

        current_year = datetime.now().year

        member_id = (
            f"FCCI-{current_year}-{next_no:06d}"
        )
        full_name = request.form["full_name"]
        contact = request.form["contact"]
        birthday = request.form["birthday"]
        email = request.form["email"]
        address = request.form["address"]
        photo = request.files["photo"]
        

        photo_filename = ""

        if photo:

            photo_filename = secure_filename(
                photo.filename
            )

            photo.save(
                os.path.join(
                    app.config["UPLOAD_FOLDER"],
                    photo_filename
                )
            )

        

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
            status,
            photo_path
        )
        VALUES
        (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            "Applicant",
            photo_filename
        ))

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

            photo_filename = secure_filename(
                photo.filename
            )

            photo.save(
                os.path.join(
                    app.config["UPLOAD_FOLDER"],
                    photo_filename
                )
            )

            cursor.execute("""
            UPDATE members
            SET photo_path = ?
            WHERE member_id = ?
            """, (
                photo_filename,
                member_id
            ))

        cursor.execute("""
        UPDATE members
        SET
            full_name = ?,
            contact = ?,
            birthday = ?,
            email = ?,
            address = ?
        WHERE member_id = ?
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

    cursor.execute("""
    SELECT *
    FROM members
    WHERE member_id = ?
    """, (member_id,))

    member = cursor.fetchone()

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
    WHERE member_id = ?
    """, (member_id,))

    # Attendance
    cursor.execute("""
    DELETE FROM attendance
    WHERE member_id = ?
    """, (member_id,))

    # Photos
    cursor.execute("""
    DELETE FROM member_photos
    WHERE member_id = ?
    """, (member_id,))

    # Withdrawals
    cursor.execute("""
    DELETE FROM withdrawals
    WHERE member_id = ?
    """, (member_id,))

    # ID Cards
    cursor.execute("""
    DELETE FROM id_cards
    WHERE member_id = ?
    """, (member_id,))

    # Main Member Record
    cursor.execute("""
    DELETE FROM members
    WHERE member_id = ?
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

    cursor.execute("""
    SELECT *
    FROM members
    WHERE member_id = ?
    """, (member_id,))

    member = cursor.fetchone()

    conn.close()

    return render_template(
        "view_member.html",
        member=member,
        username=session["username"]
    )

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
        WHERE member_id = ?
        """, (member_id,))

        exists = cursor.fetchone()[0]

        if exists == 0:

            conn.close()

            return "Member ID Not Found"

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
        (?, ?, ?, ?, ?, ?, ?)
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

            cursor.execute("""
            SELECT COUNT(*)
            FROM members
            WHERE member_id LIKE 'FCCI-%'
            """)

            count = cursor.fetchone()[0] + 1

            new_member_id = (
                f"FCCI-{datetime.now().year}-{count:06d}"
            )

            member_since = (
                f"{payment_month} {payment_year}"
            )

            cursor.execute("""
            UPDATE members
            SET
                member_id = ?,
                registration_fee = 20000,
                member_since = ?,
                status = 'Active'
            WHERE member_id = ?
            """, (
                new_member_id,
                member_since,
                member_id
            ))

        conn.commit()

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
        receipt_no,
        member_id,
        payment_type,
        amount,
        payment_date
    FROM payments
    ORDER BY id DESC
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

@app.route(
    "/member_id_card/<member_id>"
)
def member_id_card(member_id):

    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT *
    FROM members
    WHERE member_id = ?
    """, (member_id,))

    member = cursor.fetchone()

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

            receipt_filename = secure_filename(
                receipt_file.filename
            )

            receipt_file.save(
                os.path.join(
                    "static/receipts",
                    receipt_filename
                )
            )

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
        (?, ?, ?, ?, ?)
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
        (?, ?, ?, ?, ?)
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
"/donation_certificate/[int:donation_id](int:donation_id)"
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
    WHERE id = ?
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
    WHERE a.attendance_date = ?
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
        WHERE member_id = ?
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
        WHERE member_id = ?
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
            WHERE member_id = ?
            AND attendance_date = ?
            """, (
                member_id,
                today
            ))

            record = cursor.fetchone()

            if record:

                if not record[1]:

                    cursor.execute("""
                    UPDATE attendance
                    SET time_out = ?
                    WHERE id = ?
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
                VALUES (?, ?, ?, ?)
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
    WHERE attendance_date = ?
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
            WHERE member_id = ?
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
                WHERE member_id = ?
                AND attendance_date = ?
                """, (
                    member_id,
                    today
                ))

                record = cursor.fetchone()

                if record:

                    if not record[1]:

                        cursor.execute("""
                        UPDATE attendance
                        SET time_out = ?
                        WHERE id = ?
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
                    VALUES (?, ?, ?, ?)
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
    WHERE member_id = ?
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
    WHERE member_id = ?
    AND attendance_date = ?
    """, (
        member_id,
        today
    ))

    record = cursor.fetchone()

    if record:

        if not record[1]:

            cursor.execute("""
            UPDATE attendance
            SET time_out = ?
            WHERE id = ?
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
        VALUES (?, ?, ?, ?)
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
    WHERE member_id = ?
    AND attendance_date = ?
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

    cursor.execute("""
    SELECT *
    FROM members
    WHERE member_id = ?
    """, (member_id,))

    member = cursor.fetchone()

    if not member:

        conn.close()
        return "Member Not Found"

    cursor.execute("""
    SELECT
        payment_date,
        payment_type,
        amount
    FROM payments
    WHERE member_id = ?
    ORDER BY id DESC
    """, (member_id,))

    payments = cursor.fetchall()

    cursor.execute("""
    SELECT
        IFNULL(SUM(amount),0)
    FROM payments
    WHERE member_id = ?
    """, (member_id,))

    total_payment = cursor.fetchone()[0]
    
    cursor.execute("""
    SELECT attendance_date
    FROM attendance
    WHERE member_id = ?
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

        cursor.execute("""
        SELECT *
        FROM members
        WHERE member_id = ?
        """, (member_id,))

        member = cursor.fetchone()

        if member:

            cursor.execute("""
            SELECT
                IFNULL(SUM(amount),0)
            FROM payments
            WHERE member_id = ?
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
                (?, ?, ?, ?, ?, ?)
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
                WHERE member_id = ?
                """, (member_id,))

                cursor.execute("""
                DELETE FROM attendance
                WHERE member_id = ?
                """, (member_id,))

                cursor.execute("""
                DELETE FROM members
                WHERE member_id = ?
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
    SELECT *
    FROM members
    WHERE member_id = ?
    """, (member_id,))

    member = cursor.fetchone()

    if not member:
        conn.close()
        return "Member Not Found"
    
    cursor.execute("""
    SELECT
        IFNULL(SUM(amount),0)
    FROM payments
    WHERE member_id = ?
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

    photo = request.files["photo"]

    photo_filename = ""

    if photo and photo.filename:

        photo_filename = secure_filename(
            photo.filename
        )

        photo.save(
            os.path.join(
                app.config["UPLOAD_FOLDER"],
                photo_filename
            )
        )

    if len(member) > 11 and member[11]:

        photo_path = (
            "static/uploads/"
            + member[11]
        )

    if os.path.exists(
        photo_path
    ):

        member_photo = Image(
            photo_path,
            width=120,
            height=120
        )

        story.append(
            member_photo
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
    SELECT *
    FROM members
    WHERE member_id = ?
    """, (member_id,))

    member = cursor.fetchone()

    if not member:

        conn.close()

        return "Member Not Found"

    cursor.execute("""
    SELECT
        IFNULL(SUM(amount),0)
    FROM payments
    WHERE member_id = ?
    """, (member_id,))

    total_payment = cursor.fetchone()[0]

    cursor.execute("""
    SELECT attendance_date
    FROM attendance
    WHERE member_id = ?
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
    WHERE member_id = ?
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

    photo_path = ""

    if len(member) > 11 and member[11]:

        photo_path = (
            "static/uploads/"
            + member[11]
        )

    if os.path.exists(
        photo_path
    ):

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
            WHERE member_id=?
            AND payment_type='Monthly Contribution'
            AND payment_month=?
            AND payment_year=?
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
            WHERE member_id = ?
            AND payment_type='Monthly Contribution'
            AND payment_month = ?
            AND payment_year = ?
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
            WHERE member_id = ?
            AND payment_type='Monthly Contribution'
            AND payment_month = ?
            AND payment_year = ?
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
                WHERE member_id = ?
                AND payment_type='Registration Fee'
                """, (
                    applicant[0],
                ))

                has_payment = cursor.fetchone()[0]

                if has_payment == 0:

                    cursor.execute("""
                    DELETE FROM members
                    WHERE member_id = ?
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



@app.route(
"/approve_applicant/<member_id>"
)
def approve_applicant(member_id):


    if "username" not in session:
        return redirect("/login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    UPDATE members
    SET status='Active'
    WHERE member_id=?
    """, (
        member_id,
    ))

    conn.commit()
    conn.close()

    return redirect(
        "/registration_approval"
    )

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
    WHERE member_id=?
    """, (
        member_id,
    ))

    conn.commit()
    conn.close()

    return redirect(
        "/registration_approval"
    )

@app.route("/applicant_slip/<member_id>")
def applicant_slip(member_id):

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT *
    FROM members
    WHERE member_id=?
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

    if member[11]:

        photo_path = (
            "static/uploads/" +
            member[11]
        )

        if os.path.exists(photo_path):

            content.append(
                Image(
                    photo_path,
                    width=120,
                    height=120
                )
            )

            content.append(
                Spacer(1,10)
            )

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

    cursor.execute("""
    SELECT *
    FROM members
    WHERE status='Applicant'
    ORDER BY id DESC
    """)

    applicants = cursor.fetchall()

    conn.close()

    return render_template(
        "registration_approval.html",
        applicants=applicants
    )
    
@app.route("/logout")
def logout():

    session.clear()

    return redirect("/login")


if __name__ == "__main__":

    app.run(
        debug=True
    )