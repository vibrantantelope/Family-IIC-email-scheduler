import pandas as pd
from datetime import datetime, timedelta
import yagmail

# ----------------------------------------------------------------------------- 
# CONFIGURATION
# ----------------------------------------------------------------------------- 
EMAIL_SENDER = "invest-in-character@trailblazer-ptac.org"
EMAIL_PASSWORD = "gicc njub rabc autd"
DONATION_LINK = "https://tinyurl.com/2025InvestInCharacter"

# File paths
members_file = r"C:\Users\JohnKenny\OneDrive - Pathway to Adventure Council\Desktop\IIC-Emails\All-members-Trailblazer-Jan2024.xlsx"
presentation_dates_file = r"C:\Users\JohnKenny\OneDrive - Pathway to Adventure Council\Desktop\IIC-Emails\2025_presentation-dates.xlsx"

# ----------------------------------------------------------------------------- 
# 1. READ THE SPREADSHEETS
# ----------------------------------------------------------------------------- 
members = pd.read_excel(members_file)
presentations = pd.read_excel(presentation_dates_file)

# Ensure warm_up_sent and follow_up_sent columns exist
if 'warm_up_sent' not in presentations.columns:
    presentations['warm_up_sent'] = ''
if 'follow_up_sent' not in presentations.columns:
    presentations['follow_up_sent'] = ''

# Convert 'presentation_date' to datetime
presentations['presentation_date'] = pd.to_datetime(presentations['presentation_date'])

# ----------------------------------------------------------------------------- 
# 2. GENERATE WARM-UP AND FOLLOW-UP SCHEDULE
# ----------------------------------------------------------------------------- 
today = pd.Timestamp(datetime.now().date())  # Ensure `today` matches `datetime64` type

# Add warm-up and follow-up dates
presentations['Warm-Up Date'] = presentations['presentation_date'] - timedelta(days=3)
presentations['Follow-Up Date'] = presentations['presentation_date'] + timedelta(days=1)

# Filter for unsent emails
unsent_warm_up = presentations[
    (presentations['Warm-Up Date'] == today) & (presentations['warm_up_sent'] != 'yes')
].sort_values('Warm-Up Date').head(1)  # Get the next warm-up email

unsent_follow_up = presentations[
    (presentations['Follow-Up Date'] == today) & (presentations['follow_up_sent'] != 'yes')
].sort_values('Follow-Up Date').head(1)  # Get the next follow-up email

# ----------------------------------------------------------------------------- 
# 3. SEND TEST EMAIL TO YOURSELF
# ----------------------------------------------------------------------------- 
yag = yagmail.SMTP(EMAIL_SENDER, EMAIL_PASSWORD)
test_email_body = "<h2>Simulated Emails for Testing</h2>"

# Simulate warm-up email
if not unsent_warm_up.empty:
    row = unsent_warm_up.iloc[0]
    unit_number = row['unit number']
    presentation_date = row['presentation_date'].date()
    unit_name = row['Unit Name'] if 'Unit Name' in row and pd.notna(row['Unit Name']) else f"Unit {unit_number}"

    warm_up_body = f"""
    <p>Dear {unit_name} Families,</p>
    <p>Scouting changes lives. It builds character, teaches leadership, and helps young people become the best version of themselves.</p>
    <p>We’re looking forward to joining you at one of your unit meetings for an <strong>Invest in Character Campaign</strong> presentation on <strong>{presentation_date}</strong>.</p>
    <p>You can donate securely at <a href="{DONATION_LINK}">{DONATION_LINK}</a>. Every gift—no matter the size—makes an impact.</p>
    <p>Thank you for your commitment to Scouting and for helping us prepare young people for life.</p>
    """
    test_email_body += f"""
    <h3>Warm-Up Email</h3>
    <p><strong>Recipient:</strong> Members of {unit_name}</p>
    <p><strong>Send Date:</strong> {row['Warm-Up Date'].date()}</p>
    <p><strong>Content:</strong></p>
    {warm_up_body}
    """

# Simulate follow-up email
if not unsent_follow_up.empty:
    row = unsent_follow_up.iloc[0]
    unit_number = row['unit number']
    presentation_date = row['presentation_date'].date()
    unit_name = row['Unit Name'] if 'Unit Name' in row and pd.notna(row['Unit Name']) else f"Unit {unit_number}"

    follow_up_body = f"""
    <p>Dear {unit_name} Families,</p>
    <p>Thank you for hosting an <strong>Invest in Character Campaign</strong> presentation on <strong>{presentation_date}</strong>.</p>
    <p>If you haven’t already, please consider making a difference by donating securely at <a href="{DONATION_LINK}">{DONATION_LINK}</a>. Every gift—no matter the size—makes an impact.</p>
    <p>Your support helps keep Scouting strong and ensures we can continue building tomorrow’s leaders.</p>
    <p>Thank you for your commitment to Scouting and for helping us prepare young people for life.</p>
    """
    test_email_body += f"""
    <h3>Follow-Up Email</h3>
    <p><strong>Recipient:</strong> Members of {unit_name}</p>
    <p><strong>Send Date:</strong> {row['Follow-Up Date'].date()}</p>
    <p><strong>Content:</strong></p>
    {follow_up_body}
    """

# Send the test email
yag.send(
    to="jkenny2334@gmail.com",
    subject="Simulated Emails for Testing: Invest in Character",
    contents=test_email_body
)

print("Test email sent to jkenny2334@gmail.com.")
