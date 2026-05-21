"""
IBOH 2026 Sponsorship Email Campaign Sender
============================================
Sends personalized HTML emails to each company in the sponsor list,
attaches the IBOH 2026 proposal PDF, embeds the signature banner image,
and CCs the advisors.

Requirements:
    pip install openpyxl

Usage:
    1. Fill in your SMTP credentials in the CONFIG section below.
    2. Run: python send_campaign.py
       - First run does a DRY RUN (no emails sent) so you can review.
       - Set DRY_RUN = False to actually send.
"""

import smtplib
import time
import logging
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
from pathlib import Path
import openpyxl

# ─────────────────────────────────────────────
#  CONFIG  –  edit these before running
# ─────────────────────────────────────────────
SMTP_HOST        = "smtp.office365.com"               # e.g. smtp.office365.com for Outlook
SMTP_PORT        = 587
SMTP_USER       = "EMAIL"
SMTP_PASSWORD   = "PW"            

FROM_NAME        = "NAME"
FROM_EMAIL       = "EMAIL"
YOUR_NAME        = "NAME"

CC_RECIPIENTS    = [
    "",
]

PROPOSAL_PATH    = ".pdf"        # PDF to attach
SPONSOR_LIST     = ".xlsx"            # list Excel file
SIGNATURE_BANNER = ".png"           # signature banner image
#ensure that all materials are saved in the same folder

DELAY_BETWEEN_EMAILS = 3    # seconds between sends (avoid rate limits)
DRY_RUN              = True # True = preview only | False = actually send
# ─────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

def load_sponsors(xlsx_path: str) -> list[dict]:
    """Read company names and email addresses from the Excel file."""
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    sponsors = []
    header_skipped = False
    for row in ws.iter_rows(values_only=True):
        if not header_skipped:
            header_skipped = True
            continue
        company, emails_raw = row[0], row[1]
        if not company or not emails_raw:
            continue
        # Some rows have multiple emails separated by whitespace
        emails = [e.strip() for e in str(emails_raw).split() if "@" in e]
        if emails:
            sponsors.append({"company": company.strip(), "emails": emails})
    return sponsors

def build_email(company: str, to_emails: list[str], proposal_path: str, banner_path: str) -> MIMEMultipart:
    """Construct the MIME message for one sponsor with an embedded banner signature."""

    # Outer container
    msg = MIMEMultipart("mixed")
    msg["From"]    = f"{FROM_NAME} <{FROM_EMAIL}>"
    msg["To"]      = ", ".join(to_emails)
    msg["Cc"]      = ", ".join(CC_RECIPIENTS)
    msg["Subject"] = "Sponsorship Opportunity for International Battle of Hackers (IBOH) 2026"

    # ── HTML body with inline banner ─────────────────────────────────────────
    # The banner is referenced via cid:signature_banner — attached below as
    # a related inline image so it renders directly in the email body.
    html_body = f"""\
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222; line-height: 1.6;">

<p>Dear Sir/Madam,</p>

<p>
I hope this email finds you well. My name is <strong>{YOUR_NAME}</strong>, Public Relations Director
representing <strong>Forensic and Cybersecurity Research Centre &ndash; Student Section (FSEC-SS)</strong>
at Asia Pacific University of Technology &amp; Innovation (APU). I am reaching out to invite
<strong>{company}</strong> to be an official sponsor of the
<strong>International Battle of Hackers (IBOH) 2026</strong>, a Capture the Flag (CTF) competition
that brings university students and institutions together from across the globe.
</p>

<p>
With great achievement, we hosted over <strong>543 individual participants</strong> and
<strong>190 teams</strong> alone in 2025 despite the other successful years. This year, we are
expecting to host over <strong>250 teams</strong> and <strong>750 participants</strong> that are
qualified and thrive to win globally.
</p>

<p><strong>Details of IBOH 2026:</strong></p>
<ul>
  <li><strong>Date:</strong> 31st of October 2026 (Saturday)</li>
  <li><strong>Location:</strong> Asia Pacific University of Technology &amp; Innovation (APU)</li>
  <li><strong>Mode:</strong> Hybrid (Physical &amp; Online)</li>
  <li><strong>Category 1:</strong> Jeopardy-Style CTF &ndash; Open to all students worldwide</li>
  <li><strong>Category 2:</strong> Attack &amp; Defense CTF &ndash; Open to local industry personnel</li>
</ul>

<p><strong>As an official sponsor, {company} would benefit from:</strong></p>
<ul>
  <li>Strong brand visibility across all event platforms, materials, and communications</li>
  <li>Direct access to a pool of technically skilled cybersecurity students for recruitment and talent pipeline development</li>
  <li>Opportunities to lead or co-host industry workshops, establishing thought leadership within the cybersecurity community</li>
  <li>Recognition as a key contributor to advancing cybersecurity education on an international stage</li>
</ul>

<p>
We have attached the official <strong>IBOH 2026 Proposal</strong> for your review and detailed
information. For further questions, please feel free to reach out to us at
<a href="mailto:fsecss@staffemail.apu.edu.my">fsecss@staffemail.apu.edu.my</a>.
We are also available to conduct physical or online discussions.
</p>

<p>
We look forward to the possibility of partnering with <strong>{company}</strong> and making
IBOH 2026 a landmark event for the cybersecurity community.
</p>

<p>Best Regards,</p>

<!-- Signature banner — rendered inline -->
<img src="cid:signature_banner" alt="{YOUR_NAME} | Public Relations Manager | FSEC-SS APU"
     style="max-width:600px; width:100%; border:0;" />

</body>
</html>
"""

    # Wrap HTML + inline image in a multipart/related block
    related = MIMEMultipart("related")
    related.attach(MIMEText(html_body, "html"))

    # Embed the banner as an inline image (CID reference)
    banner = Path(banner_path)
    if banner.exists():
        with open(banner, "rb") as f:
            img = MIMEImage(f.read(), _subtype="png")
        img.add_header("Content-ID", "<signature_banner>")
        img.add_header("Content-Disposition", "inline", filename=banner.name)
        related.attach(img)
    else:
        log.warning("Banner not found: %s — signature image will be missing.", banner_path)

    msg.attach(related)

    # ── Attach the proposal PDF ───────────────────────────────────────────────
    pdf = Path(proposal_path)
    if pdf.exists():
        with open(pdf, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{pdf.name}"')
        msg.attach(part)
    else:
        log.warning("Proposal PDF not found: %s — email will be sent without attachment.", proposal_path)

    return msg


def send_campaign():
    sponsors = load_sponsors(SPONSOR_LIST)
    log.info("Loaded %d sponsor entries from %s", len(sponsors), SPONSOR_LIST)

    if DRY_RUN:
        log.info("=" * 60)
        log.info("DRY RUN — no emails will be sent. Set DRY_RUN=False to send.")
        log.info("=" * 60)

    all_recipients: list[str] = []
    for s in sponsors:
        all_recipients.extend(s["emails"])

    if DRY_RUN:
        print("\n📋  Campaign preview\n" + "─" * 50)
        for s in sponsors:
            print(f"  → {s['company']:<40}  {', '.join(s['emails'])}")
        print(f"\n  CC on every email : {', '.join(CC_RECIPIENTS)}")
        print(f"  Attachment        : {PROPOSAL_PATH}")
        print(f"  Signature banner  : {SIGNATURE_BANNER}")
        print(f"\n  Total emails to send    : {len(sponsors)}")
        print(f"  Total unique recipients : {len(set(all_recipients))}")
        print("\nSet DRY_RUN = False in the CONFIG section to actually send.\n")
        return

    # ── Live send ─────────────────────────────────────────────────────────────
    sent, failed = 0, 0
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.ehlo()
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        log.info("SMTP login successful.")

        for s in sponsors:
            try:
                msg = build_email(s["company"], s["emails"], PROPOSAL_PATH, SIGNATURE_BANNER)
                all_to = s["emails"] + CC_RECIPIENTS
                server.sendmail(FROM_EMAIL, all_to, msg.as_string())
                log.info("✓ Sent  →  %-40s  (%s)", s["company"], ", ".join(s["emails"]))
                sent += 1
                time.sleep(DELAY_BETWEEN_EMAILS)
            except Exception as exc:
                log.error("✗ Failed → %-40s  %s", s["company"], exc)
                failed += 1

    log.info("─" * 60)
    log.info("Campaign complete. Sent: %d  |  Failed: %d  |  Total: %d", sent, failed, sent + failed)


if __name__ == "__main__":
    send_campaign()
