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
        "email": "jaoqvad836@hotmail.com",
        "password": "W5eVxXOKSU",
        "refresh_token": (
            "M.C513_BAY.0.U.-CpkGHhALnQeAjoboxApX!HCGb2rOmFUyNkes8NKlB!f*u"
            "7id7KNvoqLOXRZLwQ6cswDmbfMH8jYfIitETM6cGYVEev6Plqr7a87k7ARuWId"
            "sJpir7DSiIYC1F6p7y!9NXv5OvGhilOLPH*ADikv29iTWFRDbkDofSJ2LCzymp"
            "IVhoDt7zKXB6d7P*DQYQJ*u9mUKpKS0y2XT!p8x29Ynlm6kiwu*TJ8!PXKyY"
            "JgyHyTrJ5HNFU!nZt1dHV*uer3xWFN339KBvevNFiksQmjGehcamwlBoatFo*p"
            "*ppGPxFjENsaT9am3vbsxVcb!O4Rap0QFJto3M5bdJRDMIxTam50pkgO2NUHq6"
            "ivdDkydzvby2IW7VUO*OhymppDmULO165gCzhY!AWuW06HVdyG8ed4beFf2QR3"
            "qsgO5KaUEzpZn"
        ),
        "client_id": "8b4ba9dd-3ea5-4e5f-86f1-ddba2230dcf2",
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


def extract_reset_links(text):
    """Extract password reset / account recovery links from email text."""
    # Find all URLs first
    urls = re.findall(r'https?://[^\s<>"\')]+', text)
    reset_keywords = [
        'password', 'reset', 'recover', 'restore', 'change-password',
        'resetpassword', 'account_verifications', 'verify', 'confirm',
        'security', 'signin', 'login', 'authenticate', 'authorization',
    ]
    reset_links = []
    for url in urls:
        url_lower = url.lower()
        if any(kw in url_lower for kw in reset_keywords):
            # Clean trailing punctuation
            url = url.rstrip('.,;:!?)}')
            if url not in reset_links:
                reset_links.append(url)
    return reset_links


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
            full_text = subject + " " + body
            code = extract_code(full_text)
            reset_links = extract_reset_links(body)

            if code or reset_links:
                # Check if email is within the 10-minute window
                try:
                    email_dt = parsedate_to_datetime(date_str)
                    if email_dt.tzinfo is None:
                        email_dt = email_dt.replace(tzinfo=timezone.utc)
                    now = datetime.now(timezone.utc)
                    age_seconds = (now - email_dt).total_seconds()
                    expires_in = max(0, 600 - int(age_seconds))  # 600s = 10 min
                    if expires_in <= 0:
                        continue  # Expired, look for newer one
                except Exception:
                    expires_in = 600

                mail.logout()
                return {
                    "code": code,
                    "links": reset_links if reset_links else [],
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
