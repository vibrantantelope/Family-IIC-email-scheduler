import pandas as pd 
from datetime import datetime, timedelta
import yagmail
import os
import re

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
EMAIL_SENDER = "invest-in-character@trailblazer-ptac.org"
EMAIL_PASSWORD = "YOUR APP PASSWORD" 
DONATION_LINK = "https://tinyurl.com/2025InvestInCharacter"

# File paths
MEMBERS_FILE = r"C:\Users\johnv\IIC-email-scheduler\All-members-Trailblazer-Jan2024.xlsx"
PRESENTATION_DATES_FILE = r"C:\Users\johnv\IIC-email-scheduler\2025_presentation-dates.xlsx"
LOG_FILE = r"email_scheduler_log.txt"

# Who gets the daily confirmation report
CONFIRMATION_RECIPIENT = "john.kenny@trailblazer-ptac.org"

# -----------------------------------------------------------------------------

def normalize_unit_number(unit_str: str) -> str:
    """
    Cleans up the unit number string so that, e.g.,
    'troop 0392(b)', 'TROOP 0392 (B)', ' Troop 0392 (b)'
    all become 'TROOP 0392(B)'.
    """
    s = unit_str.upper().strip()
    # Replace multiple spaces with a single space
    s = re.sub(r'\s+', ' ', s)
    # Remove any space before '(' and normalize the parentheses
    # E.g. "TROOP 0392 ( B )" -> "TROOP 0392(B)"
    s = re.sub(r'\s*\(\s*([BGF])\s*\)', r'(\1)', s)
    return s

def log_message(message):
    """Log a message to the console and a file."""
    with open(LOG_FILE, "a") as log:
        timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        log.write(f"{timestamp} {message}\n")
    print(message)

def main():
    """
    Main script logic:
      - Read presentations and members data
      - Normalize the 'unit number' fields
      - Determine which warm-up/follow-up emails need sending today
      - Send them, update the spreadsheet
      - Send a confirmation email about what was sent
    """
    # Start logging
    log_message("=== Email Scheduler Run Started ===")

    try:
        # 1. Read the spreadsheets
        presentations = pd.read_excel(PRESENTATION_DATES_FILE)
        members = pd.read_excel(MEMBERS_FILE)
        log_message("Loaded presentations and members data.")

        # Make sure warm_up_sent/follow_up_sent columns exist
        if 'warm_up_sent' not in presentations.columns:
            presentations['warm_up_sent'] = ''
        if 'follow_up_sent' not in presentations.columns:
            presentations['follow_up_sent'] = ''

        # Ensure these columns are properly converted to strings (avoiding .str errors)
        presentations['warm_up_sent'] = presentations['warm_up_sent'].fillna('').astype(str)
        presentations['follow_up_sent'] = presentations['follow_up_sent'].fillna('').astype(str)

        # Convert presentation_date to datetime
        presentations['presentation_date'] = pd.to_datetime(
            presentations['presentation_date'], errors='coerce'
        )

        # Generate warm-up/follow-up dates if not present
        if 'Warm-Up Date' not in presentations.columns:
            presentations['Warm-Up Date'] = presentations['presentation_date'] - timedelta(days=3)
        if 'Follow-Up Date' not in presentations.columns:
            presentations['Follow-Up Date'] = presentations['presentation_date'] + timedelta(days=1)

        presentations['Warm-Up Date'] = pd.to_datetime(presentations['Warm-Up Date'], errors='coerce')
        presentations['Follow-Up Date'] = pd.to_datetime(presentations['Follow-Up Date'], errors='coerce')

        # Normalize unit numbers in presentations
        presentations['unit number'] = (
            presentations['unit number']
            .fillna('')
            .astype(str)
            .apply(normalize_unit_number)
        )

        # Now identify which emails to send today
        today_date = datetime.now().date()
        log_message(f"Today's date: {today_date}")

        # Warm-Up Emails
        warm_up_due = presentations[
            (presentations['Warm-Up Date'].dt.date == today_date) &
            (presentations['warm_up_sent'].str.lower() != 'yes')
        ]
        log_message(f"Warm-Up Emails Due Today: {len(warm_up_due)}")

        # Follow-Up Emails
        follow_up_due = presentations[
            (presentations['Follow-Up Date'].dt.date == today_date) &
            (presentations['follow_up_sent'].str.lower() != 'yes')
        ]
        log_message(f"Follow-Up Emails Due Today: {len(follow_up_due)}")

        # Create a mapping from normalized unit number -> list of emails
        unit_emails_map = {}
        for idx, row in members.iterrows():
            # Safely convert the unit number to string and normalize
            raw_unit_num = str(row.get('unit number', '')).strip()
            unit_num = normalize_unit_number(raw_unit_num)

            emails = str(row.get('email', '')).strip()
            if not unit_num:
                log_message(f"Skipping row with missing unit number: {row}")
                continue
            if not emails or emails.lower() == 'nan':
                continue

            try:
                # Split multiple emails by comma or semicolon and remove "nan"
                email_list = [e.strip() for e in re.split(r'[;,]', emails) if e.strip() and e.lower() != 'nan']
            except Exception as ex:
                log_message(f"Error parsing emails for unit {unit_num}: {ex}")
                continue

            # Add emails to the map
            if unit_num not in unit_emails_map:
                unit_emails_map[unit_num] = set()
            unit_emails_map[unit_num].update(email_list)

        # Convert sets to lists
        unit_emails_map = {k: list(v) for k, v in unit_emails_map.items()}
        log_message(f"Final Unit Emails Map: {unit_emails_map}")

        # Log any mismatched units from presentations
        for unit in presentations['unit number'].unique():
            if unit and unit not in unit_emails_map:
                log_message(f"Unit {unit} not found in members file. Check for data mismatches.")

        # 4. Set up the mail client
        yag = yagmail.SMTP(EMAIL_SENDER, EMAIL_PASSWORD)

        sent_emails_info = []  # Track which emails are sent (for confirmation)

        # --- WARM-UP SEND ---
        for idx, row in warm_up_due.iterrows():
            unit_num = row['unit number']  # Already normalized above
            send_to_list = unit_emails_map.get(unit_num, [])
            if not send_to_list:
                log_message(f"No emails found for unit {unit_num}. Skipping warm-up.")
                continue

            if pd.notnull(row['presentation_date']):
                presentation_date_str = row['presentation_date'].strftime('%B %d, %Y')
            else:
                presentation_date_str = "Unknown Date"

            subject = f"{unit_num}: Join Us to Invest in Character on {presentation_date_str}"
            email_body = build_warm_up_email(unit_num, presentation_date_str)

            try:
                yag.send(to=send_to_list, subject=subject, contents=email_body)
                presentations.at[idx, 'warm_up_sent'] = 'yes'
                sent_emails_info.append(("Warm-Up", unit_num, send_to_list))
                log_message(f"Warm-up email sent to unit {unit_num}. Recipients: {send_to_list}")
            except Exception as e:
                log_message(f"Failed to send warm-up email for unit {unit_num}: {e}")

        # --- FOLLOW-UP SEND ---
        for idx, row in follow_up_due.iterrows():
            unit_num = row['unit number']  # Already normalized above
            send_to_list = unit_emails_map.get(unit_num, [])
            if not send_to_list:
                log_message(f"No emails found for unit {unit_num}. Skipping follow-up.")
                continue

            if pd.notnull(row['presentation_date']):
                presentation_date_str = row['presentation_date'].strftime('%B %d, %Y')
            else:
                presentation_date_str = "Unknown Date"

            subject = f"Thank You for Supporting Scouting, {unit_num} Families!"
            email_body = build_follow_up_email(unit_num, presentation_date_str)

            try:
                yag.send(to=send_to_list, subject=subject, contents=email_body)
                presentations.at[idx, 'follow_up_sent'] = 'yes'
                sent_emails_info.append(("Follow-Up", unit_num, send_to_list))
                log_message(f"Follow-up email sent to unit {unit_num}. Recipients: {send_to_list}")
            except Exception as e:
                log_message(f"Failed to send follow-up email for unit {unit_num}: {e}")

        # 5. Save the updated presentations file
        presentations.to_excel(PRESENTATION_DATES_FILE, index=False)
        log_message("Updated presentations file saved.")

        # 6. Send confirmation email
        if sent_emails_info:
            send_confirmation_email(yag, sent_emails_info)
            log_message("Confirmation email sent.")
        else:
            log_message("No emails were sent today.")
    except Exception as e:
        log_message(f"Critical error during execution: {e}")

    log_message("=== Email Scheduler Run Completed ===")

def build_warm_up_email(unit_num, presentation_date_str):
    """Return HTML content for the warm-up email with improved formatting."""
    return f"""
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; padding: 20px;">
        <p style="margin-bottom: 15px;">Dear {unit_num} Families,</p>

        <p style="margin-bottom: 15px;">
            Scouting is more than an activity—it’s a journey of discovery and growth that transforms lives. From the excitement of outdoor adventures to the quiet moments of learning integrity and perseverance, Scouting equips young people with skills that shape their character and define their future.
        </p>

        <p style="margin-bottom: 15px;">
            This year, we are proud to present the <strong>Invest in Character Campaign</strong> during your unit meeting on 
            <strong>{presentation_date_str}</strong>. This initiative is crucial to ensuring every child has access to the life-changing experiences Scouting provides—regardless of financial circumstances. With your help, we can remove barriers, open doors, and inspire hope in every Scout who walks this path.
        </p>

        <p style="margin-bottom: 15px;">
            Your contribution to this campaign supports vital programs, training for our dedicated volunteers, and opportunities that spark imagination and ignite leadership in our Scouts. It’s more than a donation—it’s an investment in the leaders of tomorrow, the changemakers, and the dreamers who will shape a brighter future for us all.
        </p>

        <p style="margin-bottom: 15px;">
            You can make a secure donation now by visiting 
            <a href="{DONATION_LINK}" style="color: #1E90FF; text-decoration: none;">{DONATION_LINK}</a>. 
            Every contribution, no matter the size, creates ripples of impact that will last a lifetime.
        </p>

        <p style="margin-bottom: 15px;">
            Thank you for being a vital part of our Scouting community and for helping us guide young people on this incredible journey of discovery, growth, and achievement.
        </p>

        <p style="margin-bottom: 20px;">Gratefully,</p>

        <div style="margin-bottom: 15px;">
            <p style="margin: 5px 0;"><strong>Ray Okigawa</strong><br>
            Family Invest in Character Chair<br>
            <a href="mailto:rokigawa@comcast.net" style="color: #1E90FF; text-decoration: none;">rokigawa@comcast.net</a></p>

            <p style="margin: 5px 0;"><strong>Bill Burgess</strong><br>
            Finance Committee Chair<br>
            <a href="mailto:wnburg48@msn.com" style="color: #1E90FF; text-decoration: none;">wnburg48@msn.com</a></p>

            <p style="margin: 5px 0;"><strong>John Kenny</strong><br>
            Senior District Executive<br>
            <a href="mailto:john.kenny@scouting.org" style="color: #1E90FF; text-decoration: none;">john.kenny@scouting.org</a></p>
        </div>
    </div>
    """.strip()


def build_follow_up_email(unit_num, presentation_date_str):
    """Return HTML content for the follow-up email with improved styling."""
    return f"""
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; padding: 20px;">
        <p style="margin-bottom: 15px;">Dear {unit_num} Families,</p>

        <p style="margin-bottom: 15px;">
            Thank you for welcoming us to your unit meeting for the <strong>Invest in Character Campaign</strong>. 
            It was truly inspiring to connect with families like yours who embody the spirit of Scouting. 
            Your enthusiasm and dedication to this program are what make our Scouting community so strong.
        </p>

        <p style="margin-bottom: 15px;">
            As we reflect on the incredible impact Scouting has on young lives, we are reminded that this journey 
            would not be possible without the generosity of supporters like you. Your contribution ensures that 
            every Scout, regardless of their background, can experience the adventure, leadership opportunities, 
            and lifelong friendships that Scouting offers.
        </p>

        <p style="margin-bottom: 15px;">
            If you haven’t yet had the opportunity to contribute, there’s still time to join us in making a difference. 
            Your gift to the <strong>Invest in Character Campaign</strong> helps fund critical programs, provide 
            financial assistance to families in need, and support the dedicated volunteers who make Scouting exceptional.
        </p>

        <p style="margin-bottom: 15px;">
            Imagine the joy of a Scout catching their first fish, overcoming challenges on the trail, or leading a 
            project that gives back to the community. These moments are made possible by your generosity. Together, 
            we can empower young people to grow into leaders who will leave a lasting mark on the world.
        </p>

        <p style="margin-bottom: 15px;">
            Please consider making your donation today by visiting 
            <a href="{DONATION_LINK}" style="color: #1E90FF; text-decoration: none;">{DONATION_LINK}</a>. 
            Every dollar counts, and every act of kindness contributes to a brighter future for our Scouts.
        </p>

        <p style="margin-bottom: 20px;">
            Thank you for your unwavering support and for helping us prepare young people to lead with courage, 
            confidence, and character.
        </p>

        <p style="margin-bottom: 20px;">With heartfelt gratitude,</p>

        <div style="margin-bottom: 15px;">
            <p style="margin: 5px 0;"><strong>Ray Okigawa</strong><br>
            Family Invest in Character Chair<br>
            <a href="mailto:rokigawa@comcast.net" style="color: #1E90FF; text-decoration: none;">rokigawa@comcast.net</a></p>

            <p style="margin: 5px 0;"><strong>Bill Burgess</strong><br>
            Finance Committee Chair<br>
            <a href="mailto:wnburg48@msn.com" style="color: #1E90FF; text-decoration: none;">wnburg48@msn.com</a></p>

            <p style="margin: 5px 0;"><strong>John Kenny</strong><br>
            Senior District Executive<br>
            <a href="mailto:john.kenny@scouting.org" style="color: #1E90FF; text-decoration: none;">john.kenny@scouting.org</a></p>
        </div>
    </div>
    """.strip()


def send_confirmation_email(yag, sent_emails_info):
    """
    Sends a single confirmation email to CONFIRMATION_RECIPIENT
    summarizing all the unit emails sent.
    """
    lines = []
    for email_type, unit_num, recipients in sent_emails_info:
        recipients_formatted = ', '.join(recipients)
        lines.append(f"<li><strong>{email_type}</strong> for Unit {unit_num} &rarr; {recipients_formatted}</li>")
    summary_list = "".join(lines)

    subject = "Daily IIC Email Confirmation"
    body = f"""
    <h2>Invest in Character - Daily Confirmation</h2>
    <p>The following emails were sent today:</p>
    <ul>{summary_list}</ul>
    <p>Regards,<br>Your Automated IIC Email Script</p>
    """
    yag.send(
        to=CONFIRMATION_RECIPIENT,
        subject=subject,
        contents=body
    )
    print(f"Confirmation email sent to {CONFIRMATION_RECIPIENT}")

if __name__ == "__main__":
    main()
