import pandas as pd
from datetime import datetime, timedelta

# ----------------------------------------------------------------------------- 
# CONFIGURATION
# ----------------------------------------------------------------------------- 
DONATION_LINK = "https://tinyurl.com/2025InvestInCharacter"

# File paths (Use real paths, or adjust as needed)
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
today = datetime.now().date()

# Add warm-up and follow-up dates
presentations['Warm-Up Date'] = presentations['presentation_date'] - timedelta(days=3)
presentations['Follow-Up Date'] = presentations['presentation_date'] + timedelta(days=1)

# Filter for future presentations only
future_presentations = presentations[presentations['presentation_date'].dt.date >= today]

# Filter for unsent emails
unsent_warm_up = future_presentations[
    (future_presentations['Warm-Up Date'] == today) & (future_presentations['warm_up_sent'] != 'yes')
]
unsent_follow_up = future_presentations[
    (future_presentations['Follow-Up Date'] == today) & (future_presentations['follow_up_sent'] != 'yes')
]

# ----------------------------------------------------------------------------- 
# 3. SIMULATE EMAILS
# ----------------------------------------------------------------------------- 
print("=== EMAIL SIMULATION FOR TODAY ===")
email_summary_today = []  # Collect all today's simulated emails
email_summary_future = []  # Collect all future emails for review

# Simulate warm-up emails for today
for _, row in unsent_warm_up.iterrows():
    unit_number = row['unit number']
    presentation_date = row['presentation_date'].date()
    warm_up_date = row['Warm-Up Date'].date()

    # Determine unit name
    unit_name = row['Unit Name'] if 'Unit Name' in row and pd.notna(row['Unit Name']) else f"Unit {unit_number}"

    # Get unit members
    unit_members = members[members['unit number'].str.strip().str.lower() == unit_number.strip().lower()]
    print(f"\nSimulating Warm-Up Email for {unit_name} (Date: {warm_up_date}):")

    for _, member in unit_members.iterrows():
        recipient_email = member['Email']
        if pd.notna(recipient_email):
            mail_body = f"""
            <p>Dear {unit_name} Families,</p>
            <p>Scouting changes lives. It builds character, teaches leadership, and helps young people become the best version of themselves.</p>
            <p>We’re looking forward to joining you at one of your unit meetings for an <strong>Invest in Character Campaign</strong> presentation on <strong>{presentation_date}</strong>.</p>
            <p>You can donate securely at <a href="{DONATION_LINK}">{DONATION_LINK}</a>. Every gift—no matter the size—makes an impact.</p>
            <p>Thank you for your commitment to Scouting and for helping us prepare young people for life.</p>
            """
            email_summary_today.append({
                "Recipient Email": recipient_email,
                "Subject": f"Reminder: Invest in Character Presentation for {unit_name}",
                "Body": mail_body
            })
            print(f"  - {recipient_email}")

# Simulate follow-up emails for today
for _, row in unsent_follow_up.iterrows():
    unit_number = row['unit number']
    presentation_date = row['presentation_date'].date()
    follow_up_date = row['Follow-Up Date'].date()

    unit_name = row['Unit Name'] if 'Unit Name' in row and pd.notna(row['Unit Name']) else f"Unit {unit_number}"

    unit_members = members[members['unit number'].str.strip().str.lower() == unit_number.strip().lower()]
    print(f"\nSimulating Follow-Up Email for {unit_name} (Date: {follow_up_date}):")

    for _, member in unit_members.iterrows():
        recipient_email = member['Email']
        if pd.notna(recipient_email):
            follow_up_body = f"""
            <p>Dear {unit_name} Families,</p>
            <p>Thank you for hosting an <strong>Invest in Character Campaign</strong> presentation on <strong>{presentation_date}</strong>.</p>
            <p>If you haven’t already, please consider making a difference by donating securely at <a href="{DONATION_LINK}">{DONATION_LINK}</a>. Every gift—no matter the size—makes an impact.</p>
            <p>Your support helps keep Scouting strong and ensures we can continue building tomorrow’s leaders.</p>
            <p>Thank you for your commitment to Scouting and for helping us prepare young people for life.</p>
            """
            email_summary_today.append({
                "Recipient Email": recipient_email,
                "Subject": f"Thank You: Invest in Character Presentation for {unit_name}",
                "Body": follow_up_body
            })
            print(f"  - {recipient_email}")

# Simulate all future emails
for _, row in future_presentations.iterrows():
    unit_number = row['unit number']
    presentation_date = row['presentation_date'].date()
    warm_up_date = row['Warm-Up Date'].date()
    follow_up_date = row['Follow-Up Date'].date()

    unit_name = row['Unit Name'] if 'Unit Name' in row and pd.notna(row['Unit Name']) else f"Unit {unit_number}"

    unit_members = members[members['unit number'].str.strip().str.lower() == unit_number.strip().lower()]
    print(f"\nSimulating Future Emails for {unit_name}:")

    # Simulate warm-up email
    for _, member in unit_members.iterrows():
        recipient_email = member['Email']
        if pd.notna(recipient_email):
            email_summary_future.append({
                "Recipient Email": recipient_email,
                "Subject": f"Reminder: Invest in Character Presentation for {unit_name}",
                "Body": f"Simulated Warm-Up Email for {presentation_date}"
            })

    # Simulate follow-up email
    for _, member in unit_members.iterrows():
        recipient_email = member['Email']
        if pd.notna(recipient_email):
            email_summary_future.append({
                "Recipient Email": recipient_email,
                "Subject": f"Thank You: Invest in Character Presentation for {unit_name}",
                "Body": f"Simulated Follow-Up Email for {presentation_date}"
            })

# ----------------------------------------------------------------------------- 
# 4. MARK AS SENT
# ----------------------------------------------------------------------------- 
# Mark warm-up emails as sent
presentations.loc[unsent_warm_up.index, 'warm_up_sent'] = 'yes'

# Mark follow-up emails as sent
presentations.loc[unsent_follow_up.index, 'follow_up_sent'] = 'yes'

# Save updated spreadsheet
presentations.to_excel(presentation_dates_file, index=False)

# ----------------------------------------------------------------------------- 
# 5. DISPLAY SIMULATED EMAILS
# ----------------------------------------------------------------------------- 
print("\n=== SIMULATED EMAILS FOR TODAY ===")
for email in email_summary_today:
    print(f"\nTo: {email['Recipient Email']}")
    print(f"Subject: {email['Subject']}")
    print(f"Body: {email['Body']}")

print("\n=== SIMULATED EMAILS FOR ALL FUTURE DATES ===")
for email in email_summary_future:
    print(f"\nTo: {email['Recipient Email']}")
    print(f"Subject: {email['Subject']}")
    print(f"Body: {email['Body']}")

# ----------------------------------------------------------------------------- 
# 6. DONE
# ----------------------------------------------------------------------------- 
print("\nSimulation complete. No emails were actually sent.")
