from flask import Flask, render_template, jsonify
import requests
import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
from datetime import datetime, timezone
import re
import base64
import time

app = Flask(__name__)

# ─── Account config ───────────────────────────────────────────────
ACCOUNTS = [
    {
        "email": "grud_1908@outlook.com",
        "password": "S21FOtHzxD",
        "refresh_token": (
            "M.C544_BAY.0.U.-CvIzhVW8lD9cLXMC8fuxSpUZhT8G87AZItEeXgigD4ylu"
            "*sd45MzGAp3lRgYhCWJwbfzgckNarC7K2MLKZsbkKvmo4jLwwIjw!aWwFkuUe*"
            "QjlYqQuaXW49LJCnX*cX18g9lWMvNBy6Rn*FkfrCzGiWQPzHL*TzHgNWb6ynj"
            "4H3COkRgsK8bnZ2gx!U0xgWHtAheu2q1bjFENlZ6MJSDoLDpK6SRRkp8n*P6Q"
            "IvkM!QuAYSW!SNnjdZWwIkpu9aOOk!LdwZANV2UaIFkUvyDprrKcsWIh1R3SfT"
            "2cySvCDiLew0!s0wzJPoqSAq2U47tWKlxvtTI8TYh8*oKrCohrSV1*uFLa8HM1"
            "n8f6VNi654rTCLutXyMF1ylGQ0HPT8!in9yfKXoX70lYIMs1QZUDvc$"
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
    """Build XOAUTH2 SASL string (raw bytes, not base64)."""
    auth_string = f"user={user}\x01auth=Bearer {token}\x01\x01"
    return auth_string.encode()


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


def extract_code(text):
    """Intelligently extract a verification code from email text."""
    # Remove URLs (they contain tons of random numbers)
    clean = re.sub(r'https?://\S+', '', text)
    # Remove HTML tags
    clean = re.sub(r'<[^>]+>', ' ', clean)
    # Remove common false-positive numbers (zip codes, street numbers, years)
    # by checking contextually

    # Strategy 1: Code near keywords like "code", "verify", "pin"
    keyword_patterns = [
        r'(?:code|код|pin|otp|passcode)[\s:;—\-]*[\r\n\s]*(\d{4,8})\b',
        r'(?:enter|entering)[\s\S]{0,50}?(\d{4,8})\b',
        r'(?:verification|security code|confirm|verify)[\s\S]{0,80}?(\d{4,8})\b',
    ]
    for pattern in keyword_patterns:
        matches = re.findall(pattern, clean, re.IGNORECASE)
        if matches:
            return matches[0]

    # Strategy 2: Standalone number on its own line (very likely a code)
    line_codes = re.findall(r'^\s*(\d{5,8})\s*$', clean, re.MULTILINE)
    if line_codes:
        return line_codes[0]

    # Strategy 3: Any 6-8 digit number, filtering false positives
    all_codes = re.findall(r'\b(\d{6,8})\b', clean)
    bad = {str(y) for y in range(1990, 2030)}
    filtered = [c for c in all_codes if c not in bad]
    if filtered:
        return filtered[0]

    # Strategy 4: 4-5 digit codes, filtering years
    short = re.findall(r'\b(\d{4,5})\b', clean)
    filtered = [c for c in short if c not in bad]
    if filtered:
        return filtered[0]

    return None


def fetch_latest_code(account):
    """Fetch the latest email that contains a verification code."""
    token = get_access_token(account)
    if not token:
        return {"error": "Failed to obtain access token"}

    try:
        mail = imaplib.IMAP4_SSL("outlook.office365.com", 993)
        auth_bytes = generate_oauth2_string(account["email"], token)
        mail.authenticate("XOAUTH2", lambda x: auth_bytes)
        mail.select("INBOX")

        status, messages = mail.search(None, "ALL")
        if status != "OK":
            return {"error": "Failed to search mailbox"}

        msg_ids = messages[0].split()
        if not msg_ids:
            return {"code": None, "message": "Inbox is empty"}

        # Check last 10 emails to find the most recent one with a code
        check_ids = msg_ids[-10:]
        check_ids.reverse()

        for mid in check_ids:
            status, msg_data = mail.fetch(mid, "(RFC822)")
            if status != "OK":
                continue
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = decode_mime_words(msg.get("Subject", ""))
            sender = decode_mime_words(msg.get("From", ""))
            date_str = msg.get("Date", "")
            body = extract_body(msg)
            code = extract_code(subject + " " + body)

            if code:
                # Check if email is within the 10-minute window
                try:
                    email_dt = parsedate_to_datetime(date_str)
                    if email_dt.tzinfo is None:
                        email_dt = email_dt.replace(tzinfo=timezone.utc)
                    now = datetime.now(timezone.utc)
                    age_seconds = (now - email_dt).total_seconds()
                    expires_in = max(0, 600 - int(age_seconds))  # 600s = 10 min
                    if expires_in <= 0:
                        continue  # Code expired, look for newer one
                except Exception:
                    expires_in = 600  # If can't parse date, show anyway

                mail.logout()
                return {
                    "code": code,
                    "from": sender,
                    "subject": subject,
                    "date": date_str,
                    "expires_in": expires_in,
                }

        mail.logout()
        return {"code": None, "message": "No codes found"}

    except Exception as e:
        return {"error": str(e)}


# ─── Routes ──────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html", accounts=ACCOUNTS)


@app.route("/api/code/<int:account_idx>")
def api_code(account_idx):
    if account_idx < 0 or account_idx >= len(ACCOUNTS):
        return jsonify({"error": "Invalid account index"}), 400
    data = fetch_latest_code(ACCOUNTS[account_idx])
    return jsonify(data)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
