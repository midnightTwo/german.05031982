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
        "email": "avramenko_19982010@outlook.com",
        "password": "fN5egOO016",
        "refresh_token": (
            "M.C546_BAY.0.U.-Cg0l96WUAiSbclZcScnpnoHXY4Krv7e1649hGv4vmv8g"
            "W4EXmXbrOCKOXloBsedNLrfB5on54fRvtbUVQZWpnx2DuQn!6ldgXNDjpL*TSg"
            "P9MQnDOM9mL5KIyC*HNc0cMsu3W6zWbzCqhr3JV8rfzULQ*5soGScDR3M586*"
            "6A3O*HEA29N*Si4PAxCbMGjIZKBbw0RT4WAjoXdGdUeOUle4Gxr!paflPFdYk"
            "Kedhhu5k9FDw30gsSh2KR9wN4X0cx2mit4ZqbXDgO5848GOWXJQrwQF71PX*P"
            "3t!GqGbgdXd44QurUT2AvpD3fDHamlobZ5UYTjBE8jpF1DI0yOQFmAXXDKN14"
            "wyq8Lh!0P2Ao9ECRicizhemmP!R1WmL3oerJtHYKNUd3Kj003X5v!jvlIlMmz"
            "FaiVuPKfYVlRIGBmf"
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


def fetch_latest_emails(account, count=5):
    """Fetch the latest N emails via IMAP/XOAUTH2."""
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
