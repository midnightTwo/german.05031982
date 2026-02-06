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
        "email": "golovchenko_19992909@outlook.com",
        "password": "OvwYmtT5id",
        "refresh_token": (
            "M.C544_BAY.0.U.-CqPp6mjB0Exm33ecKA5ofUP0GkpRb2U1CdcavikDUrdm"
            "s5n9IQDadRTviAqUiRmXNl2SYwMp69Dvk7r1XeQX0YMZgxaMNq4V1Au9YaLo"
            "Qv5DskDXBm7Z2B9qy*aPdQBfHo2ro41tWqmmkyvg1HTqJ!BA04cqz5E!P1Gz"
            "Dm*YoUhUSVOjm*ZrrFL!xER8U1CX5L3zgkG0eHiCh0tGvGw5*l1Jdi4cMczA"
            "bIFQLBrWSSJ4QlRKYrbDZRe*rk8m2qL36IEv66B53zuz5trEl9cWLCNkfYQGu"
            "MuLbf6A3yvQ4reVzHDigk0iyE5f3mprpa3bTWoET8huCMsvPMlbU4k3BXya0!"
            "zbfbbzYVF8EppINjorgFISmDF30Gg4Ld04h9EyRv4drXZTF2oWuHFR6as3RZAV"
            "kvJas!6FE1wp6Mt6imOh"
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


def extract_codes(text):
    """Try to find verification codes in text."""
    codes = re.findall(r'\b(\d{4,8})\b', text)
    return codes


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
            codes = extract_codes(subject + " " + body)

            if codes:
                mail.logout()
                return {
                    "code": codes[-1],
                    "from": sender,
                    "subject": subject,
                    "date": date_str,
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
