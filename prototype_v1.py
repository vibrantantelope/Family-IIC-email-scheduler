import pandas as pd
from datetime import datetime, timedelta
import yagmail
import os
import re
import logging
from typing import List, Dict, Tuple

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------

# Email Configuration
EMAIL_SENDER = "invest-in-character@trailblazer-ptac.org"
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")  # Securely fetch from environment variables
DONATION_LINK = "https://tinyurl.com/2025InvestInCharacter"

# File Paths
MEMBERS_FILE = "All-members-Trailblazer-Jan2024.xlsx"
PRESENTATION_DATES_FILE = "2025_presentation-dates.xlsx"
LOG_FILE = "email_scheduler_log.txt"

# Confirmation Recipient
CONFIRMATION_RECIPIENT = "john.kenny@trailblazer-ptac.org"

# -----------------------------------------------------------------------------

# Setup Logging
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s %(levelname)s: %(message)s", "%Y-%m-%d %H:%M:%S")
console.setFormatter(formatter)
logging.getLogger().addHandler(console)


def normalize_unit_number(unit_str: str) -> str:
    """
    Cleans up the unit number string to a standardized format, e.g., 'TROOP 0392(B)'.
    """
    s = unit_str.upper().strip()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'\s*\(\s*([BGF])\s*\)', r'(\1)', s)
    return s


def load_data() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Loads the presentations and members data from Excel files.
    """
    try:
        presentations = pd.read_excel(PRESENTATION_DATES_FILE)
        members = pd.read_excel(MEMBERS_FILE)
        logging.info("Successfully loaded presentations and members data.")
        return presentations, members
    except Exception as e:
        logging.error(f"Error loading data: {e}")
        raise


def prepare_presentations_df(presentations: pd.DataFrame) -> pd.DataFrame:
    """
    Prepares the presentations DataFrame by ensuring necessary columns and formats.
    """
    # Initialize columns if they don't exist
    for col in ['warm_up_sent', 'follow_up_sent']:
        if col not in presentations.columns:
            presentations[col] = ''

    presentations['warm_up_sent'] = presentations['warm_up_sent'].fillna('').astype(str)
    presentations['follow_up_sent'] = presentations['follow_up_sent'].fillna('').astype(str)

    # Convert dates to datetime
    presentations['presentation_date'] = pd.to_datetime(
        presentations['presentation_date'], errors='coerce'
    )

    # Generate Warm-Up and Follow-Up Dates if absent
    today = datetime.now().date()
    if 'Warm-Up Date' not in presentations.columns:
        presentations['Warm-Up Date'] = presentations['presentation_date'] - timedelta(days=3)
    if 'Follow-Up Date' not in presentations.columns:
        presentations['Follow-Up Date'] = presentations['presentation_date'] + timedelta(days=1)

    presentations['Warm-Up Date'] = pd.to_datetime(presentations['Warm-Up Date'], errors='coerce')
    presentations['Follow-Up Date'] = pd.to_datetime(presentations['Follow-Up Date'], errors='coerce')

    # Normalize unit numbers
    presentations['unit number'] = presentations['unit number'].fillna('').astype(str).apply(normalize_unit_number)

    logging.info("Prepared presentations DataFrame.")
    return presentations


def build_unit_emails_map(members: pd.DataFrame) -> Dict[str, List[str]]:
    """
    Builds a mapping from normalized unit numbers to their associated email addresses.
    """
    unit_emails_map = {}
    for _, row in members.iterrows():
        raw_unit_num = str(row.get('unit number', '')).strip()
        unit_num = normalize_unit_number(raw_unit_num)
        emails = str(row.get('email', '')).strip()

        if not unit_num:
            logging.warning(f"Skipping member with missing unit number: {row.to_dict()}")
            continue
        if not emails or emails.lower() == 'nan':
            continue

        email_list = [e.strip() for e in re.split(r'[;,]', emails) if e.strip() and e.lower() != 'nan']

        if unit_num not in unit_emails_map:
            unit_emails_map[unit_num] = set()
        unit_emails_map[unit_num].update(email_list)

    # Convert sets to sorted lists for consistency
    unit_emails_map = {unit: sorted(emails) for unit, emails in unit_emails_map.items()}
    logging.info(f"Built unit emails map with {len(unit_emails_map)} units.")
    return unit_emails_map


def identify_emails_to_send(presentations: pd.DataFrame, today: datetime.date) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Identifies which warm-up and follow-up emails are due to be sent today.
    """
    warm_up_due = presentations[
        (presentations['Warm-Up Date'].dt.date == today) &
        (presentations['warm_up_sent'].str.lower() != 'yes')
    ]

    follow_up_due = presentations[
        (presentations['Follow-Up Date'].dt.date == today) &
        (presentations['follow_up_sent'].str.lower() != 'yes')
    ]

    logging.info(f"Warm-Up Emails Due Today: {len(warm_up_due)}")
    logging.info(f"Follow-Up Emails Due Today: {len(follow_up_due)}")
    return warm_up_due, follow_up_due


def send_emails(
    yag: yagmail.SMTP,
    due_presentations: pd.DataFrame,
    unit_emails_map: Dict[str, List[str]],
    email_type: str
) -> List[Tuple[str, str, List[str]]]:
    """
    Sends warm-up or follow-up emails based on the email_type.
    """
    sent_emails_info = []
    for idx, row in due_presentations.iterrows():
        unit_num = row['unit number']
        recipients = unit_emails_map.get(unit_num, [])

        if not recipients:
            logging.warning(f"No emails found for unit {unit_num}. Skipping {email_type.lower()} email.")
            continue

        presentation_date = row['presentation_date']
        presentation_date_str = presentation_date.strftime('%B %d, %Y') if pd.notnull(presentation_date) else "Unknown Date"

        if email_type == "Warm-Up":
            subject = f"{unit_num}: Join Us to Invest in Character on {presentation_date_str}"
            email_body = build_warm_up_email(unit_num, presentation_date_str)
            status_col = 'warm_up_sent'
        elif email_type == "Follow-Up":
            subject = "Thank You for Supporting Scouting, {unit_num} Families!".format(unit_num=unit_num)
            email_body = build_follow_up_email(unit_num, presentation_date_str)
            status_col = 'follow_up_sent'
        else:
            logging.error(f"Invalid email type: {email_type}")
            continue

        try:
            yag.send(to=recipients, subject=subject, contents=email_body)
            due_presentations.at[idx, status_col] = 'yes'
            sent_emails_info.append((email_type, unit_num, recipients))
            logging.info(f"{email_type} email sent to unit {unit_num}. Recipients: {recipients}")
        except Exception as e:
            logging.error(f"Failed to send {email_type.lower()} email for unit {unit_num}: {e}")

    return sent_emails_info


def build_warm_up_email(unit_num: str, presentation_date_str: str) -> str:
    """Generates HTML content for the warm-up email."""
    return f"""
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; padding: 20px;">
        <p>Dear {unit_num} Families,</p>
        <p>
            Scouting is more than an activity—it’s a journey of discovery and growth that transforms lives...
            <!-- [Content Truncated for Brevity] -->
        </p>
        <p>
            You can make a secure donation now by visiting 
            <a href="{DONATION_LINK}" style="color: #1E90FF; text-decoration: none;">{DONATION_LINK}</a>. 
            Every contribution, no matter the size, creates ripples of impact that will last a lifetime.
        </p>
        <p>Thank you for being a vital part of our Scouting community and for helping us guide young people on this incredible journey of discovery, growth, and achievement.</p>
        <p>Gratefully,</p>
        <div>
            <p><strong>Ray Okigawa</strong><br>
            Family Invest in Character Chair<br>
            <a href="mailto:rokigawa@comcast.net" style="color: #1E90FF;">rokigawa@comcast.net</a></p>
            <p><strong>Bill Burgess</strong><br>
            Finance Committee Chair<br>
            <a href="mailto:wnburg48@msn.com" style="color: #1E90FF;">wnburg48@msn.com</a></p>
            <p><strong>John Kenny</strong><br>
            Senior District Executive<br>
            <a href="mailto:john.kenny@scouting.org" style="color: #1E90FF;">john.kenny@scouting.org</a></p>
        </div>
    </div>
    """.strip()


def build_follow_up_email(unit_num: str, presentation_date_str: str) -> str:
    """Generates HTML content for the follow-up email."""
    return f"""
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; padding: 20px;">
        <p>Dear {unit_num} Families,</p>
        <p>
            Thank you for welcoming us to your unit meeting for the <strong>Invest in Character Campaign</strong>. 
            It was truly inspiring to connect with families like yours who embody the spirit of Scouting...
            <!-- [Content Truncated for Brevity] -->
        </p>
        <p>
            Please consider making your donation today by visiting 
            <a href="{DONATION_LINK}" style="color: #1E90FF; text-decoration: none;">{DONATION_LINK}</a>. 
            Every dollar counts, and every act of kindness contributes to a brighter future for our Scouts.
        </p>
        <p>With heartfelt gratitude,</p>
        <div>
            <p><strong>Ray Okigawa</strong><br>
            Family Invest in Character Chair<br>
            <a href="mailto:rokigawa@comcast.net" style="color: #1E90FF;">rokigawa@comcast.net</a></p>
            <p><strong>Bill Burgess</strong><br>
            Finance Committee Chair<br>
            <a href="mailto:wnburg48@msn.com" style="color: #1E90FF;">wnburg48@msn.com</a></p>
            <p><strong>John Kenny</strong><br>
            Senior District Executive<br>
            <a href="mailto:john.kenny@scouting.org" style="color: #1E90FF;">john.kenny@scouting.org</a></p>
        </div>
    </div>
    """.strip()


def send_confirmation_email(
    yag: yagmail.SMTP,
    sent_emails_info: List[Tuple[str, str, List[str]]]
):
    """
    Sends a confirmation email summarizing all sent emails.
    """
    if not sent_emails_info:
        logging.info("No emails were sent today. Skipping confirmation email.")
        return

    summary_lines = [
        f"<li><strong>{email_type}</strong> for Unit {unit_num} &rarr; {', '.join(recipients)}</li>"
        for email_type, unit_num, recipients in sent_emails_info
    ]
    summary_list = "\n".join(summary_lines)

    subject = "Daily IIC Email Confirmation"
    body = f"""
    <h2>Invest in Character - Daily Confirmation</h2>
    <p>The following emails were sent today:</p>
    <ul>
        {summary_list}
    </ul>
    <p>Regards,<br>Your Automated IIC Email Script</p>
    """.strip()

    try:
        yag.send(
            to=CONFIRMATION_RECIPIENT,
            subject=subject,
            contents=body
        )
        logging.info(f"Confirmation email sent to {CONFIRMATION_RECIPIENT}")
    except Exception as e:
        logging.error(f"Failed to send confirmation email: {e}")


def save_presentations(presentations: pd.DataFrame):
    """
    Saves the updated presentations DataFrame back to the Excel file.
    """
    try:
        presentations.to_excel(PRESENTATION_DATES_FILE, index=False)
        logging.info("Updated presentations file saved.")
    except Exception as e:
        logging.error(f"Failed to save presentations file: {e}")


def main():
    """
    Main function orchestrating the email scheduling and sending process.
    """
    logging.info("=== Email Scheduler Run Started ===")
    today = datetime.now().date()

    try:
        # Load Data
        presentations, members = load_data()

        # Prepare Data
        presentations = prepare_presentations_df(presentations)
        unit_emails_map = build_unit_emails_map(members)

        # Identify Emails to Send
        warm_up_due, follow_up_due = identify_emails_to_send(presentations, today)

        # Initialize Yagmail Client
        try:
            yag = yagmail.SMTP(EMAIL_SENDER, EMAIL_PASSWORD)
            logging.info("Initialized Yagmail SMTP client.")
        except Exception as e:
            logging.error(f"Failed to initialize email client: {e}")
            return

        # Send Warm-Up Emails
        sent_emails_info = send_emails(yag, warm_up_due, unit_emails_map, "Warm-Up")

        # Send Follow-Up Emails
        sent_emails_info += send_emails(yag, follow_up_due, unit_emails_map, "Follow-Up")

        # Save Updates
        save_presentations(presentations)

        # Send Confirmation Email
        send_confirmation_email(yag, sent_emails_info)

    except Exception as e:
        logging.critical(f"Critical error during execution: {e}")

    logging.info("=== Email Scheduler Run Completed ===")


if __name__ == "__main__":
    main()
