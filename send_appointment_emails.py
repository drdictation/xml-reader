#!/usr/bin/env python3
"""
Send personalised appointment update emails to patients from a CSV file.

Usage:
    1. Create a .env or export environment variables:
       export EMAIL_ADDRESS="your@gmail.com"
       export EMAIL_PASSWORD="your-app-password"
    2. Prepare patients.csv with columns: name, email, appointment_date, appointment_time, mode
    3. Set DRY_RUN = True to preview emails without sending
    4. Run: python send_appointment_emails.py
"""

import csv
import os
import smtplib
import time
import re
from datetime import datetime
from email.mime.text import MIMEText

# ──────────────────────────────────────────────
# CONFIGURATION
# ──────────────────────────────────────────────
DRY_RUN = False  # Set to False when ready to send real emails

# Send emails based on this list
CSV_INPUT = "output/future-appointments_SEND.csv"

# Mark 'DONE' in this master file
CSV_TRACKING = "output/future-appointments_REAL.csv"

CSV_LOG = "email_log.csv"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

EMAIL_ADDRESS = "drchamarabasnayake@gmail.com"
EMAIL_PASSWORD = "pjub nlup ukaz rrou"

SEND_DELAY_SECONDS = 2  # delay between each email


# ──────────────────────────────────────────────
# EMAIL TEMPLATES (HTML)
# ──────────────────────────────────────────────

SUBJECT = "Your upcoming appointment – change of location"

BODY_HTML = """\
<html>
<body>
<p>Dear {name},</p>

<p>I am writing to let you know that I will be relocating my private consulting practice from 13 April 2026.</p>

<p><b><u>Your appointment date and time remain unchanged.</u></b></p>

<p>If you are attending in person, your upcoming appointment on <b>{appointment_date}</b> at <b>{appointment_time}</b> will be held at:</p>

<p>
    Focus Gastroenterology<br>
    Suite 201, Level 2<br>
    100 Victoria Parade<br>
    East Melbourne VIC 3002<br>
    (Opposite St Vincent's Public Hospital)
</p>

<p>
    Tel: 03 9650 7917<br>
    Email: office@focusgastro.com.au
</p>

<p>
    View location on Google Maps:<br>
    <a href="https://maps.google.com/?q=100+Victoria+Parade+East+Melbourne">https://maps.google.com/?q=100+Victoria+Parade+East+Melbourne</a>
</p>

<p>If your appointment is via telehealth, you can continue to use the same link:<br>
<a href="https://doxy.me/chamaracmg">https://doxy.me/chamaracmg</a></p>

<p>If you have any questions or need to make changes to your booking, please contact the rooms directly using the details above.
</p>

<p>I look forward to seeing you at your appointment.</p>

<p>Kind regards,<br><br>
Chamara<br>
A/Prof Chamara Basnayake</p>
</body>
</html>
"""


# ──────────────────────────────────────────────
# FUNCTIONS
# ──────────────────────────────────────────────

def format_date(date_str: str) -> str:
    """Convert common date formats (YYYY-MM-DD, D/M/YYYY) to DD - Month - Year."""
    if not date_str:
        return ""
    
    # Try multiple common formats
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime("%d - %B - %Y")
        except ValueError:
            continue
    
    # If all fail, return raw string
    return date_str


def get_first_name(full_name: str) -> str:
    """Extract first name by skipping common titles (Mrs, Ms, Mr, Dr, etc.)."""
    if not full_name:
        return "Patient"
    
    # Common titles to skip
    titles = {"mr", "mrs", "ms", "miss", "dr", "prof", "a/prof", "sir", "madam"}
    parts = full_name.split()
    
    for part in parts:
        clean_part = part.strip(".,").lower()
        if clean_part not in titles:
            return part  # First non-title part found
            
    return full_name.split()[0] if full_name.split() else "Patient"


def load_csv(filepath: str) -> list[dict]:
    """Load and validate patient rows from the input (SEND) CSV."""
    patients = []
    if not os.path.exists(filepath):
        print(f"  ✗  File '{filepath}' not found.")
        return []

    with open(filepath, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader, start=2):  # row 1 is header
            # Strip whitespace from all values
            row = {k.strip(): v.strip() for k, v in row.items() if k}

            # Skip truly empty rows
            if not any(row.values()):
                continue

            # Map to expected fields from output/future-appointments.csv
            full_name = row.get("patient_name", "")
            email = row.get("email", "")
            appointment_date_raw = row.get("appointment_date", "")
            appointment_time = row.get("appointment_time", "")

            # Validation
            if not email or "@" not in email:
                # Some emails might be in "Name <email>" format from the CSV preview
                # e.g. "William Au <iamwillau38@gmail.com>"
                if "<" in email and ">" in email:
                    email = email.split("<")[-1].split(">")[0]
                else:
                    print(f"  ⚠  Row {i}: Skipping '{full_name}' – invalid/missing email '{email}'")
                    continue

            patients.append({
                "full_name": full_name,  # Keep full name for logging
                "name": get_first_name(full_name),  # Use first name for greeting
                "email": email,
                "appointment_date_raw": appointment_date_raw, # Store raw for matching
                "appointment_date": format_date(appointment_date_raw),
                "appointment_time": appointment_time,
            })

    print(f"  ✓  Loaded {len(patients)} patient(s) from {filepath}\n")
    return patients


def load_master_csv(filepath: str) -> tuple[list[dict], list[str]]:
    """Load the tracking (REAL) CSV file into a list of dicts."""
    if not os.path.exists(filepath):
        print(f"  ✗  Master tracking file '{filepath}' not found.")
        return [], []
    
    with open(filepath, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames or []
        rows = list(reader)
    return rows, fieldnames


def save_master_csv(filepath: str, fieldnames: list[str], rows: list[dict]):
    """Save the updated master tracking file back to disk."""
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def _match_patient(row: dict, patient: dict) -> bool:
    """Check if a master CSV row matches a patient dict."""
    return (row.get("patient_name", "").strip() == patient["full_name"].strip() and
            row.get("email", "").strip() == patient["email"].strip() and
            row.get("appointment_date", "").strip() == patient["appointment_date_raw"].strip() and
            row.get("appointment_time", "").strip() == patient["appointment_time"].strip())


def is_already_done(master_rows: list[dict], patient: dict) -> bool:
    """Check if the patient is already marked as DONE in the master list."""
    for row in master_rows:
        if _match_patient(row, patient):
            if row.get("DONE", "").strip().lower() == "x":
                return True
    return False


def update_master_done(master_rows: list[dict], patient: dict) -> bool:
    """Find a match in master_rows and mark 'DONE' = 'x'. Returns True if matched."""
    found = False
    for row in master_rows:
        if _match_patient(row, patient):
            row["DONE"] = "x"
            found = True
    return found


def generate_email(patient: dict) -> tuple[str, str]:
    """Return (subject, body_html) for a given patient."""
    subject = SUBJECT
    body = BODY_HTML.format(
        name=patient["name"],
        appointment_date=patient["appointment_date"],
        appointment_time=patient["appointment_time"],
    )
    return subject, body


def send_email(
    smtp_conn: smtplib.SMTP | None,
    from_addr: str,
    to_addr: str,
    subject: str,
    body: str,
    dry_run: bool = True,
    bcc_addr: str | None = None,
) -> bool:
    """Send a single email. Returns True on success."""
    msg = MIMEText(body, "html", "utf-8")
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg["Subject"] = subject

    if dry_run:
        print("─" * 60)
        print(f"  TO:      {to_addr}")
        if bcc_addr:
            print(f"  BCC:     {bcc_addr}")
        print(f"  SUBJECT: {subject}")
        print(f"  BODY:\n{body}")
        print("─" * 60)
        return True

    recipients = [to_addr]
    if bcc_addr:
        recipients.append(bcc_addr)

    smtp_conn.sendmail(from_addr, recipients, msg.as_string())
    return True


def log_result(log_writer, full_name: str, email: str, status: str):
    """Append one row to the log CSV."""
    log_writer.writerow({
        "name": full_name,
        "email": email,
        "status": status,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })


def main():
    print("=" * 60)
    print("  APPOINTMENT EMAIL SENDER")
    print(f"  Mode: {'DRY RUN (no emails will be sent)' if DRY_RUN else 'LIVE – emails WILL be sent'}")
    print("=" * 60 + "\n")

    # ── Load patients (SEND list) ──
    patients = load_csv(CSV_INPUT)
    if not patients:
        return

    # ── Load master tracking (REAL list) ──
    master_rows, master_fields = load_master_csv(CSV_TRACKING)
    if not master_rows:
        print("  ✗  Proceeding without master tracking.\n")

    # ── Establish connection (only for live) ──
    smtp_conn = None
    if not DRY_RUN:
        print("  Connecting to Gmail SMTP…")
        try:
            smtp_conn = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            smtp_conn.starttls()
            smtp_conn.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            print("  ✓  SMTP connection established.\n")
        except Exception as e:
            print(f"  ✗  SMTP login failed: {e}")
            return

    sent_count = 0
    fail_count = 0
    skip_count = 0

    # Ensure log directory or file is ready (Appending to log)
    with open(CSV_LOG, "a", newline="", encoding="utf-8") as f:
        fieldnames = ["timestamp", "name", "email", "status"]
        log_writer = csv.DictWriter(f, fieldnames=fieldnames)
        if f.tell() == 0:
            log_writer.writeheader()

        for idx, patient in enumerate(patients, start=1):
            full_name = patient["full_name"]
            email = patient["email"]
            print(f"  [{idx}/{len(patients)}] {full_name} ({email})… ", end="", flush=True)

            # ── Skip if already sent ──
            if master_rows and is_already_done(master_rows, patient):
                print("⏭ already sent – skipping")
                skip_count += 1
                continue

            try:
                subject, body = generate_email(patient)
                send_email(
                    smtp_conn,
                    from_addr=EMAIL_ADDRESS,
                    to_addr=email,
                    subject=subject,
                    body=body,
                    dry_run=DRY_RUN,
                    bcc_addr=EMAIL_ADDRESS,
                )
                print("✓ success")
                log_result(log_writer, full_name, email, "sent" if not DRY_RUN else "dry_run")
                
                # Update Master Tracking file only if NOT dry run
                if master_rows:
                    matched = update_master_done(master_rows, patient)
                    if DRY_RUN:
                        if matched:
                            print("    → (Dry Run: would mark DONE in master file)")
                        else:
                            print("    → (Dry Run: no match found in master file)")
                    elif matched:
                        save_master_csv(CSV_TRACKING, master_fields, master_rows)

                sent_count += 1

            except Exception as e:
                print(f"✗ FAILED – {e}")
                log_result(log_writer, full_name, email, f"failed: {e}")
                fail_count += 1

            # Delay between sends
            if idx < len(patients):
                time.sleep(SEND_DELAY_SECONDS)

    # ── Cleanup ──
    if smtp_conn:
        smtp_conn.quit()

    # ── Summary ──
    print("\n" + "=" * 60)
    print(f"  DONE  —  Sent: {sent_count}  |  Skipped: {skip_count}  |  Failed: {fail_count}")
    print(f"  Log saved to: {CSV_LOG}")
    print("=" * 60)


if __name__ == "__main__":
    main()
