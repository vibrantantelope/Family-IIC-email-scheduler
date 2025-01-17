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

# Add warm-up and follow-up presentation_date to the presentations DataFrame
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
# 3. SIMULATE EMAILS FOR ALL FUTURE DATES
# -----------------------------------------------------------------------------
print("=== EMAIL SCHEDULE FOR ALL FUTURE PRESENTATIONS ===")

for _, row in future_presentations.iterrows():
    unit_number = row['unit number']
    presentation_date = row['presentation_date'].date()
    warm_up_date = row['Warm-Up Date'].date()
    follow_up_date = row['Follow-Up Date'].date()

    # If you have "Unit Name," use it; else fallback
    if 'Unit Name' in row and pd.notna(row['Unit Name']):
        unit_name = row['Unit Name']
    else:
        unit_name = f"Unit {unit_number}"

    # Grab members in that unit
    unit_members = members[members['unit number'].str.strip().str.lower() == unit_number.strip().lower()]


    print(f"\nPresentation for {unit_name} (Unit {unit_number})")
    print(f"  - Presentation Date: {presentation_date}")
    print(f"  - Warm-Up Email: {warm_up_date}")
    print(f"  - Follow-Up Email: {follow_up_date}")

    # Warm-Up Email Recipients
    print(f"  Warm-Up Email Recipients ({len(unit_members)} members):")
    for _, member in unit_members.iterrows():
        recipient_email = member['Email']
        if recipient_email:
            print(f"    - {recipient_email}")

    # Follow-Up Email Recipients
    print(f"  Follow-Up Email Recipients ({len(unit_members)} members):")
    for _, member in unit_members.iterrows():
        recipient_email = member['Email']
        if recipient_email:
            print(f"    - {recipient_email}")

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
# 4. DONE
# -----------------------------------------------------------------------------
print("\nSimulation complete. No emails were actually sent.")
