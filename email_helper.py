import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
 
load_dotenv()
 
GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
 
 
def send_email(to_email, subject, html_body):
    """
    Magpadala ng email gamit ang Gmail SMTP.
 
    Params:
        to_email   — email address ng tatanggap
        subject    — subject line ng email
        html_body  — HTML content ng email body
 
    Returns:
        True  — successful na naipadala
        False — may error (hindi nag-crash ang app, log lang)
    """
 
    if not to_email or "@" not in to_email:
        print(f"[EMAIL] Invalid email address: {to_email}")
        return False
 
    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        print("[EMAIL] Walang GMAIL_ADDRESS o GMAIL_APP_PASSWORD sa .env")
        return False
 
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = f"FCCI Filipino Community Center <{GMAIL_ADDRESS}>"
        msg["To"] = to_email
 
        html_part = MIMEText(html_body, "html")
        msg.attach(html_part)
 
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_ADDRESS, to_email, msg.as_string())
 
        print(f"[EMAIL] Successfully sent to {to_email}")
        return True
 
    except Exception as e:
        print(f"[EMAIL] Error sending to {to_email}: {e}")
        return False
 
 
def send_welcome_email(to_email, full_name, member_id):
    """
    Ipinapadala kapag na-approve ang isang applicant at naging
    Official FCCI Member.
    """
 
    subject = "🎉 Welcome to FCCI - You are now an Official Member!"
 
    html_body = f"""
    <div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;padding:20px;">
      <div style="background:linear-gradient(135deg,#00562a,#00a85c);padding:30px;border-radius:16px 16px 0 0;text-align:center;">
        <h1 style="color:#ffffff;margin:0;font-size:24px;">Welcome to FCCI!</h1>
      </div>
      <div style="background:#f4fbf7;padding:30px;border-radius:0 0 16px 16px;border:1px solid #e0f0e5;">
        <p style="font-size:15px;color:#0c2418;">Kumusta <b>{full_name}</b>,</p>
        <p style="font-size:14.5px;color:#3a5045;line-height:1.6;">
          Maligayang pagdating sa <b>Filipino Community Center International (FCCI)</b>!
          Ang iyong registration ay na-approve na, at ikaw ay opisyal nang miyembro ng aming komunidad.
        </p>
        <div style="background:#ffffff;border:1px solid #d5ece0;border-radius:12px;padding:16px 20px;margin:20px 0;">
          <p style="margin:0;font-size:12px;color:#5c8270;text-transform:uppercase;letter-spacing:0.5px;">Your Member ID</p>
          <p style="margin:4px 0 0;font-size:20px;font-weight:700;color:#00562a;">{member_id}</p>
        </div>
        <p style="font-size:14.5px;color:#3a5045;line-height:1.6;">
          Maaari ka nang mag-login sa member portal para tingnan ang iyong profile,
          makipag-ugnayan sa community feed, at marami pang iba.
        </p>
        <p style="font-size:13.5px;color:#5c8270;margin-top:24px;">
          United in Faith, Serving with Love.<br>
          — FCCI Team
        </p>
      </div>
    </div>
    """
 
    return send_email(to_email, subject, html_body)
 
 
def send_payment_confirmation_email(to_email, full_name, payment_type, amount, receipt_no, payment_date):
    """
    Ipinapadala kapag na-record ang isang payment (Registration Fee
    o Monthly Contribution).
    """
 
    subject = f"✅ Payment Received - {payment_type}"
 
    html_body = f"""
    <div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;padding:20px;">
      <div style="background:linear-gradient(135deg,#00562a,#00a85c);padding:30px;border-radius:16px 16px 0 0;text-align:center;">
        <h1 style="color:#ffffff;margin:0;font-size:24px;">Payment Confirmed</h1>
      </div>
      <div style="background:#f4fbf7;padding:30px;border-radius:0 0 16px 16px;border:1px solid #e0f0e5;">
        <p style="font-size:15px;color:#0c2418;">Kumusta <b>{full_name}</b>,</p>
        <p style="font-size:14.5px;color:#3a5045;line-height:1.6;">
          Natanggap namin ang iyong bayad. Salamat sa iyong patuloy na suporta sa FCCI!
        </p>
        <div style="background:#ffffff;border:1px solid #d5ece0;border-radius:12px;padding:18px 20px;margin:20px 0;">
          <table style="width:100%;font-size:13.5px;color:#0c2418;">
            <tr><td style="padding:6px 0;color:#5c8270;">Receipt No.</td><td style="padding:6px 0;text-align:right;font-weight:600;">{receipt_no}</td></tr>
            <tr><td style="padding:6px 0;color:#5c8270;">Payment Type</td><td style="padding:6px 0;text-align:right;font-weight:600;">{payment_type}</td></tr>
            <tr><td style="padding:6px 0;color:#5c8270;">Amount</td><td style="padding:6px 0;text-align:right;font-weight:700;color:#00562a;">₩{amount:,}</td></tr>
            <tr><td style="padding:6px 0;color:#5c8270;">Date</td><td style="padding:6px 0;text-align:right;font-weight:600;">{payment_date}</td></tr>
          </table>
        </div>
        <p style="font-size:13.5px;color:#5c8270;margin-top:24px;">
          United in Faith, Serving with Love.<br>
          — FCCI Team
        </p>
      </div>
    </div>
    """
 
    return send_email(to_email, subject, html_body)