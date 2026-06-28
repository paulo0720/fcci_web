import os
import threading
import urllib.request
import urllib.error
import json


def _send_email_now(to_email, subject, html_body):
    """
    Nagpapadala ng email gamit ang Resend API (HTTPS) —
    hindi SMTP, kaya hindi naka-block sa Render free tier.
    """

    RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "").strip()
    FROM_EMAIL = "FCCI Filipino Community Center <delivered@resend.dev>"

    if not to_email or "@" not in to_email:
        print(f"[EMAIL] Invalid email address: {to_email}")
        return False

    if not RESEND_API_KEY:
        print("[EMAIL] ERROR: Walang RESEND_API_KEY sa environment variables")
        return False

    try:
        payload = json.dumps({
            "from": FROM_EMAIL,
            "to": [to_email],
            "subject": subject,
            "html": html_body
        }).encode("utf-8")

        req = urllib.request.Request(
            "https://api.resend.com/emails",
            data=payload,
            headers={
                "Authorization": f"Bearer {RESEND_API_KEY}",
                "Content-Type": "application/json"
            },
            method="POST"
        )

        with urllib.request.urlopen(req, timeout=15) as response:
            result = json.loads(response.read().decode("utf-8"))
            print(f"[EMAIL] Successfully sent to {to_email} | ID: {result.get('id')}")
            return True

    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8")
        print(f"[EMAIL] HTTP {e.code} error: {error_body}")
        return False
    except Exception as e:
        print(f"[EMAIL] Unexpected error sending to {to_email}: {e}")
        return False


def send_email(to_email, subject, html_body):
    thread = threading.Thread(
        target=_send_email_now,
        args=(to_email, subject, html_body),
        daemon=True
    )
    thread.start()
    return True


def send_welcome_email(to_email, full_name, member_id):
    subject = "Maligayang pagdating sa FCCI - Ikaw ay Official Member na!"

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
          Maaari ka nang mag-login sa member portal para tingnan ang iyong profile at makipag-ugnayan sa community.
        </p>
        <p style="font-size:13.5px;color:#5c8270;margin-top:24px;">
          United in Faith, Serving with Love.<br>
          — FCCI Team
        </p>
      </div>
    </div>
    """

    return send_email(to_email, subject, html_body)


def send_payment_confirmation_email(
    to_email, full_name, payment_type,
    amount, receipt_no, payment_date
):
    subject = f"Payment Received - {payment_type}"

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
            <tr><td style="padding:6px 0;color:#5c8270;">Amount</td><td style="padding:6px 0;text-align:right;font-weight:700;color:#00562a;">&#8361;{amount:,}</td></tr>
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
