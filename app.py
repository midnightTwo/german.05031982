from flask import Flask, render_template, jsonify
import requests
import imaplib
import email
from email.header import decode_header
import re
import base64
import time

app = Flask(__name__)

# ─── Account config ───────────────────────────────────────────────
ACCOUNTS = [
    {
        "email": "galatiuk.19920703@outlook.com",
        "password": "3edUIRXeOB",
        "refresh_token": (
            "M.C555_BAY.0.U.-CiQ2wjZhd4CD2b6f0dfwdXEv7EBYij6e6!fG8ZaeVv0e"
            "PK8RP2JuTCzd7HC89eSOywNebQ37vLebDgM8!baCDqs5H8L2!vETh21gtUOZO"
            "rVCwpUX9SdpuH4c3TeoBj1NmPobbdUUErXzmhT9voMIoGQ*9Is*AHQe0HS4dz"
            "YF37bQ2U1rV12!rAXMlzy4Rd0UJQ4Efm2gO0jbnvpHQl4dkM1t3VnjfhQoXW"
            "0OFGntrxcnbEYtBrIXZncz7quHBsg!tOeQZMKV9LFOXm8Lpf5!C0FCnfpOdP"
            "9*FKbq*w750o9n49P3ZdchPhQNPjy9hv3silWKgec*ulYCmNSO3lRkQNYxj9T"
            "0lp5IgwiNlLmHhfbsCb5XmTv6ik7U5ciUFmghrQZJup57T5pLSxdC20zov0U$"
        ),
        "client_id": "9e5f94bc-e8a4-4e73-b8be-63364c29d753",
    },
]


def get_access_token(account):
    """Get access token via refresh token."""
    try:
        r = requests.post(
            "https://login.microsoftonline.com/common/oauth2/v2.0/token",
            data={
                "client_id": account["client_id"],
                "refresh_token": account["refresh_token"],
                "grant_type": "refresh_token",
                "scope": "https://outlook.office.com/IMAP.AccessAsUser.All offline_access",
            },
            timeout=15,
        )
        data = r.json()
        if "access_token" in data:
            return data["access_token"]
        return None
    except Exception as e:
        print(f"Token error: {e}")
        return None


def generate_oauth2_string(user, token):
    """Build XOAUTH2 SASL string."""
    auth_string = f"user={user}\x01auth=Bearer {token}\x01\x01"
    return base64.b64encode(auth_string.encode()).decode()


def decode_mime_words(s):
    """Decode MIME encoded-word tokens."""
    if not s:
        return ""
    decoded_parts = []
    for part, charset in decode_header(s):
        if isinstance(part, bytes):
            decoded_parts.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded_parts.append(part)
    return " ".join(decoded_parts)


def extract_body(msg):
    """Get plain-text body from an email message."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            if ctype == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    body += payload.decode(charset, errors="replace")
            elif ctype == "text/html" and not body:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    body += payload.decode(charset, errors="replace")
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            body = payload.decode(charset, errors="replace")
    return body


def extract_codes(text):
    """Try to find verification codes in text."""
    codes = re.findall(r'\b(\d{4,8})\b', text)
    return codes


def fetch_latest_emails(account, count=5):
    """Fetch the latest N emails via IMAP/XOAUTH2."""
    token = get_access_token(account)
    if not token:
        return {"error": "Failed to obtain access token"}

    try:
        mail = imaplib.IMAP4_SSL("outlook.office365.com", 993)
        auth_string = generate_oauth2_string(account["email"], token)
        mail.authenticate("XOAUTH2", lambda x: auth_string.encode())
        mail.select("INBOX")

        status, messages = mail.search(None, "ALL")
        if status != "OK":
            return {"error": "Failed to search mailbox"}

        msg_ids = messages[0].split()
        if not msg_ids:
            return {"emails": [], "message": "Inbox is empty"}

        latest_ids = msg_ids[-count:]
        latest_ids.reverse()

        results = []
        for mid in latest_ids:
            status, msg_data = mail.fetch(mid, "(RFC822)")
            if status != "OK":
                continue
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = decode_mime_words(msg.get("Subject", ""))
            sender = decode_mime_words(msg.get("From", ""))
            date_str = msg.get("Date", "")
            body = extract_body(msg)
            codes = extract_codes(subject + " " + body)

            results.append({
                "subject": subject,
                "from": sender,
                "date": date_str,
                "codes": codes,
                "body_preview": body[:300],
            })

        mail.logout()
        return {"emails": results}

    except Exception as e:
        return {"error": str(e)}


# ─── Routes ──────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html", accounts=ACCOUNTS)


@app.route("/api/emails/<int:account_idx>")
def api_emails(account_idx):
    if account_idx < 0 or account_idx >= len(ACCOUNTS):
        return jsonify({"error": "Invalid account index"}), 400
    data = fetch_latest_emails(ACCOUNTS[account_idx])
    return jsonify(data)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
