# ----------------------------------------------------------------------------- 
# 3. SIMULATE EMAILS
# ----------------------------------------------------------------------------- 
print("=== EMAIL SIMULATION FOR TODAY ===")
email_summary = []  # List to collect all email simulations

# Process warm-up emails
for _, row in unsent_warm_up.iterrows():
    unit_number = row['unit number']
    presentation_date = row['presentation_date'].date()
    warm_up_date = row['Warm-Up Date'].date()

    unit_name = row['Unit Name'] if 'Unit Name' in row and pd.notna(row['Unit Name']) else f"Unit {unit_number}"

    unit_members = members[members['unit number'].str.strip().str.lower() == unit_number.strip().lower()]
    print(f"\nSimulating Warm-Up Email for {unit_name} (Date: {warm_up_date}):")

    for _, member in unit_members.iterrows():
        recipient_email = member['Email']
        if pd.notna(recipient_email):
            # Warm-Up Email Body
            mail_body = f"""
            <p>Dear {unit_name} Families,</p>
            <p>Scouting changes lives. It builds character, teaches leadership, and helps young people become the best version of themselves.</p>
            <p>We’re looking forward to joining you at one of your unit meetings for an <strong>Invest in Character Campaign</strong> presentation on <strong>{presentation_date}</strong>.</p>
            <p>You can donate securely at <a href="{DONATION_LINK}">{DONATION_LINK}</a>. Every gift—no matter the size—makes an impact.</p>
            <p>Thank you for your commitment to Scouting and for helping us prepare young people for life.</p>
            """
            email_summary.append({
                "Recipient Email": recipient_email,
                "Subject": f"Reminder: Invest in Character Presentation for {unit_name}",
                "Body": mail_body
            })
            print(f"  - {recipient_email}")

# Process follow-up emails
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
            # Follow-Up Email Body
            follow_up_body = f"""
            <p>Dear {unit_name} Families,</p>
            <p>Thank you for hosting an <strong>Invest in Character Campaign</strong> presentation on <strong>{presentation_date}</strong>.</p>
            <p>If you haven’t already, please consider making a difference by donating securely at <a href="{DONATION_LINK}">{DONATION_LINK}</a>. Every gift—no matter the size—makes an impact.</p>
            <p>Your support helps keep Scouting strong and ensures we can continue building tomorrow’s leaders.</p>
            <p>Thank you for your commitment to Scouting and for helping us prepare young people for life.</p>
            """
            email_summary.append({
                "Recipient Email": recipient_email,
                "Subject": f"Thank You: Invest in Character Presentation for {unit_name}",
                "Body": follow_up_body
            })
            print(f"  - {recipient_email}")

# ----------------------------------------------------------------------------- 
# 4. OUTPUT SIMULATED EMAILS
# ----------------------------------------------------------------------------- 
print("\n=== SIMULATED EMAILS ===")
for email in email_summary:
    print(f"\nTo: {email['Recipient Email']}")
    print(f"Subject: {email['Subject']}")
    print(f"Body: {email['Body']}")
