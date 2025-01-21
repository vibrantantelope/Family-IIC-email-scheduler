Invest in Character Email Scheduler
This project automates the process of sending Invest in Character (IIC) Campaign emails to unit families. It ensures that warm-up emails are sent 3 days before a presentation, and follow-up emails are sent the day after, based on a schedule in an Excel file. The program is designed to save time, ensure consistency, and increase engagement with Scouting families.

Features
Automated Email Sending

Sends warm-up emails to units 3 days before their scheduled presentation.
Sends follow-up emails to units the day after their presentation.
Spreadsheet Integration

Pulls data from:
2025_presentation-dates.xlsx: Contains unit presentation dates and email statuses.
All-members-Trailblazer-Jan2024.xlsx: Contains unit numbers and family email addresses.
Daily Summary Report

Sends a confirmation email summarizing the emails sent each day.
Logs all actions in email_scheduler_log.txt for easy troubleshooting and auditing.
Error Handling

Skips units with missing or invalid email addresses and logs the issue.
Updates the spreadsheet to ensure emails are only sent once.
File Structure
livev1.py: Main Python script for running the email scheduler.
2025_presentation-dates.xlsx: Spreadsheet containing unit presentation dates and email statuses.
All-members-Trailblazer-Jan2024.xlsx: Spreadsheet containing unit email addresses.
email_scheduler_log.txt: Log file for tracking actions, errors, and email history.
run_iic_emails.bat: Batch file for automating the script execution with Task Scheduler.
requirements.txt: List of Python dependencies.
venv/: Virtual environment for running the project.
Setup Instructions
1. Install Dependencies
Create a virtual environment:
bash
Copy code
python -m venv venv
Activate the virtual environment:
Windows:
bash
Copy code
venv\Scripts\activate
Mac/Linux:
bash
Copy code
source venv/bin/activate
Install dependencies:
bash
Copy code
pip install -r requirements.txt
2. Configure the Script
Update the following variables in livev1.py:

EMAIL_SENDER: Your sender email address (e.g., invest-in-character@trailblazer-ptac.org).
EMAIL_PASSWORD: The password or app-specific password for your email account.
DONATION_LINK: The link for the Invest in Character donation page.
Ensure the file paths for 2025_presentation-dates.xlsx and All-members-Trailblazer-Jan2024.xlsx are correct.

3. Test the Script
Run the script manually to ensure it works as expected:

bash
Copy code
python livev1.py
4. Automate with Task Scheduler
Use the run_iic_emails.bat batch file to schedule the script.
Set up a daily trigger in Task Scheduler to run the batch file at your desired time.
Logs and Troubleshooting
Logs: Check email_scheduler_log.txt for details on sent emails, skipped units, and any errors.
Debugging: Ensure the unit number fields in both spreadsheets match exactly (e.g., no extra spaces or mismatched formats).
Future Improvements
Tailor emails to each unit's interests (e.g., legacy, adventure, community impact).
Enhance the reporting system to include email open rates and engagement metrics.
Why This Matters
This tool saves hours of manual work, ensures consistent communication with families, and helps maximize contributions to the IIC Campaign. By automating repetitive tasks, we can focus on building relationships and making a bigger impact in our Scouting community.
