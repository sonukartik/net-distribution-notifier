import imaplib
import email
from email.header import decode_header
import re
import pandas as pd
import bs4
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import json
from datetime import datetime
from dotenv import load_dotenv

# --- Load environment variables ---
load_dotenv()
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SEEN_FILE = "seen_emails.json"

# --- Subjects to search ---
subjects_to_search = {
    "IndiGrid": ["IndiGrid Distribution Advice"],
    "Embassy REIT": ["Embassy REIT Distribution Advice", "Embassy Office Parks REIT", "Embassy REIT"],
    "Bharat Highways": ["Bharat Highways Invit"],
    "Capital Infra": ["INDUSINVIT"],
    "Nexus Trust REIT": ["Nexus Select Trust ReIT - Distribution Advice", "NEXUS SELECT TRUST", "Nexus Select Trust"]
}

# --- Helper Functions ---
def load_seen_ids():
    if os.path.exists(SEEN_FILE):
        with open(SEEN_FILE, "r") as f:
            return set(json.load(f))
    return set()

def save_seen_ids(seen_ids):
    with open(SEEN_FILE, "w") as f:
        json.dump(list(seen_ids), f)

def get_email_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                return part.get_payload(decode=True).decode(errors='ignore')
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                html = part.get_payload(decode=True).decode(errors='ignore')
                soup = bs4.BeautifulSoup(html, 'html.parser')
                return soup.get_text(separator='\n')
    else:
        return msg.get_payload(decode=True).decode(errors='ignore')
    return ""

def extract_total_net_distribution(body_text):
    pattern = re.compile(r'^E\s+Net Distribution.*?([?₹]?\s*[\d,]+\.\d+)\s*$', re.MULTILINE | re.IGNORECASE)
    matches = pattern.findall(body_text)
    if matches:
        last_value = matches[-1]
        return re.sub(r'[?₹\s]', '', last_value)
    for line in body_text.splitlines():
        if line.strip().startswith("E") and "Net Distribution" in line:
            numbers = re.findall(r'[?₹]?\s*[\d,]+\.\d+', line)
            if numbers:
                return re.sub(r'[?₹\s]', '', numbers[-1])
    return None

def send_email_with_table(data_entries):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "New Net Distribution Email(s) Received"
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = EMAIL_ADDRESS

    html = """<html><body>
    <h2 style="color: #2e6c80;">New Net Distribution Details:</h2>
    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; font-family: Arial;">
    <tr style="background-color: #f2f2f2;">
        <th style="background-color: #d4edda;">Company</th>
        <th style="background-color: #ffe5b4;">Subject (Link)</th>
        <th style="background-color: #d0e8f2;">Date</th>
    </tr>"""

    for entry in data_entries:
        formatted_date = entry['Date'].strftime('%d-%b-%Y %I:%M %p')
        link = f"https://mail.google.com/mail/u/0/#search/rfc822msgid:{entry['MessageID']}"
        html += f"""<tr>
            <td style="background-color: #d4edda;">{entry['Company']}</td>
            <td style="background-color: #ffe5b4;">
                <a href="{link}" target="_blank">{entry['Subject']}</a>
            </td>
            <td style="background-color: #d0e8f2;">{formatted_date}</td>
        </tr>"""

    html += "</table></body></html>"

    part = MIMEText(html, 'html')
    msg.attach(part)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, EMAIL_ADDRESS, msg.as_string())

# --- Main Execution ---
def main():
    data = []
    seen_ids = load_seen_ids()
    updated_seen_ids = set(seen_ids)

    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    mail.select('inbox')

    for company, variants in subjects_to_search.items():
        all_email_ids = []
        for subj in variants:
            status, messages = mail.search(None, f'(SUBJECT "{subj}")')
            if status != 'OK':
                continue
            all_email_ids.extend(messages[0].split())

        all_email_ids = list(set(all_email_ids))
        if not all_email_ids:
            continue

        emails_info = []
        for e_id in all_email_ids:
            _, msg_data = mail.fetch(e_id, '(RFC822)')
            msg = email.message_from_bytes(msg_data[0][1])
            dt = email.utils.parsedate_to_datetime(msg.get('Date'))
            body = get_email_body(msg)
            net_dist = extract_total_net_distribution(body)
            if net_dist:
                message_id = msg.get('Message-ID').strip('<>')
                if message_id in seen_ids:
                    continue
                emails_info.append((e_id, dt, msg, net_dist, message_id))

        if not emails_info:
            continue

        emails_info.sort(key=lambda x: x[1], reverse=True)
        _, latest_date, latest_msg, _, message_id = emails_info[0]
        raw_subj, enc = decode_header(latest_msg['Subject'])[0]
        subj_decoded = raw_subj.decode(enc or 'utf-8') if isinstance(raw_subj, bytes) else raw_subj

        data.append({
            'Company': company,
            'Subject': subj_decoded,
            'Date': latest_date,
            'MessageID': message_id
        })
        updated_seen_ids.add(message_id)

    mail.logout()

    if data:
        df = pd.DataFrame(data)
        df['Date'] = pd.to_datetime(df['Date'])
        df = df.sort_values('Date', ascending=False)
        send_email_with_table(df.to_dict(orient='records'))
        save_seen_ids(updated_seen_ids)
        print(f"✅ Found {len(data)} new records. Notification sent.")
    else:
        print("✅ No new distribution emails found.")

if __name__ == "__main__":
    main()
