import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import os
import time

# Constants
MAIL_LIMIT = 70  # Max emails per sender per day
TIME_GAP = 300  # 5 minutes in seconds
MAIN_FILE = 'company_data1.xlsx'  # Main Excel file with recipients
MAILER_FILE = 'Mailer.csv'  # Temp file for recipients
SENDER_FILE = 'Sender.csv'  # Sender file with credentials
LOG_FILE = 'email_logs.csv'  # File to log email activities
DAILY_COUNT_FILE = 'daily_count_logs.csv'  # Daily count logs file

EMAIL_SUBJECT = "99Papers - Find perfect documents for your business now - Reg"
HTML_CONTENT_FILE = '99papers-content.html'
ALIAS_TEXT = """<p>Hello Team,&nbsp;</p>
<p>Best wishes on the successful incorporation of your company! and Welcome to the business ecosystem.</p>
<p>You must have received the company documents such as AoA, MoA and Certificate of Incorporation from the Ministry of Corporate Affairs. However, many other documents are essential for a company to operate without facing any legal trouble. As overwhelming as it is to start a new company, the burden of the sorting out the required formalities and documentation often adds to the worries, leading to missed deadlines and high penalties. Essential Legal &amp; HR Documents bundle by&nbsp;<span class="il">99papers</span>&nbsp;is a set of pre-prepared editable documents required at different stages of the company, which can be used directly only by affixing your company letterhead to it.&nbsp;</p>
<p>At&nbsp;<a href="https://www.99papers.in/">99papers.in</a>, we take pride in offering a vast array of esteemed documents tailored to meet the diverse requirements of your organisation. Whether you are navigating through lawful intricacies, streamlining HR processes, or seeking templates for various business activities, we have got you covered.</p>
<p>The bundle contains documents fit for a wide range of processes &amp; applications such as -&nbsp;</p>
<ul>
<li>COVID-19 related documents&nbsp; | Statutory Documents&nbsp;| Non-Disclosure &amp; Copyright</li>
<li>Hiring | &nbsp;Performance Review | Employee Policies</li>
<li>Recruitment | HR Forms | HR Letters</li>
<li>Co-founders' agreement | Tenancy Agreement | NDA | Freelancer agreement&nbsp;</li>
</ul>
<p>Some of the important documents you'll find in the bundle are listed below&nbsp;</p>
<table border="1">
<tbody>
<tr>
<td>Legal Documents &amp; Templates&nbsp;</td>
<td>HR Documents and templates</td>
</tr>
<tr>
<td>1. NDA(Non-disclosure agreement)</td>
<td>- Employee Information form</td>
</tr>
<tr>
<td>2. New employee offer letter</td>
<td>- Leave request form</td>
</tr>
<tr>
<td>3. Conflict of interest</td>
<td>- Performance evaluation form</td>
</tr>
<tr>
<td>4. Confidentiality agreement</td>
<td>- Employee Incident report form</td>
</tr>
<tr>
<td>5. Employee appraisal form</td>
<td>- Return to Work form</td>
</tr>
<tr>
<td>6. Company loan agreement</td>
<td>- Covid-19 Self Screening form</td>
</tr>
<tr>
<td>7. Invention agreement</td>
<td>- Employee Reference Request&nbsp;</td>
</tr>
<tr>
<td>8. Contractor agreement</td>
<td>- Employee Complaint form</td>
</tr>
<tr>
<td>9. Advisor agreement</td>
<td>- Employee Nomination form</td>
</tr>
<tr>
<td>10. Important agreements like partnership deed, dealership agreement etc.</td>
<td>- Labour application form</td>
</tr>
<tr>
<td>+ 4000 other important documents &amp; templates&nbsp;</td>
<td>&nbsp;</td>
</tr>
</tbody>
</table>
<p>&nbsp;</p>
<p><strong>How does it help?</strong></p>
<p>Legal documents are not something you can prepare without experience and We can't stress enough how important it is that the documents you use are compliant with the latest laws, guidelines and statutes.&nbsp; All documents provided by&nbsp;<span class="il">99papers</span>&nbsp;are updated every quarter and if there are any changes in the documents you are using, you'll receive a notification regarding the same through e-mail.&nbsp;</p>
<p>&nbsp;</p>
<p><strong>Cost - Rs. 1299, including taxes.</strong></p>
<p>&nbsp;</p>
<p>Visit our website with the link here -&nbsp;<a href="https://99papers.in/" target="_blank" rel="noopener">Visit Website</a>.</p>
<p>You can get further information and the link to payment of the bundle using the link here -&nbsp;<a href="https://docs.99papers.in/99Papers%20-%20Sample%20Documents.pdf" target="_blank" rel="noopener">SAMPLE DOCUMENTS</a>.</p>
<p>If you have any questions, feel free to reply to this mail or contact us at <a href="mailto:support@99papers.in">support@99papers.in</a> or +91 6379934788 with your query.&nbsp;</p>
<p>Regards,</p>
<p>Team&nbsp;<span class="il">99Papers, IN</span></p>
"""  # Replace with your full alias text content

def load_email_content():
    """Load HTML email content or fallback to alias text."""
    if os.path.exists(HTML_CONTENT_FILE):
        with open(HTML_CONTENT_FILE, 'r', encoding='utf-8') as file:
            return file.read()
    return ALIAS_TEXT

def append_to_csv(file, data):
    """Append a row to a CSV file, creating it if necessary."""
    df = pd.DataFrame([data])
    if not os.path.exists(file):
        df.to_csv(file, mode='a', header=True, index=False)
    else:
        df.to_csv(file, mode='a', header=False, index=False)

def log_email_activity(sender_email, recipient_email):
    """Log email activity to email_logs.csv."""
    send_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_entry = {
        'sender_email': sender_email,
        'recipient_email': recipient_email,
        'send_time': send_time
    }
    append_to_csv(LOG_FILE, log_entry)

def log_daily_count(daily_count):
    """Logs the daily count of sent emails per sender."""
    today_date = datetime.now().strftime('%Y-%m-%d')

    # Check if the daily count file exists. If it does, read it.
    if os.path.exists(DAILY_COUNT_FILE):
        daily_count_df = pd.read_csv(DAILY_COUNT_FILE)
    else:
        # If the file doesn't exist, create a new DataFrame.
        daily_count_df = pd.DataFrame(columns=['date', 'sender_email', 'mail_count'])

    # Loop through the daily_count dictionary and update the counts.
    for sender_email, count in daily_count.items():
        # Check if this sender and date already exist in the DataFrame.
        existing_row = daily_count_df[(daily_count_df['date'] == today_date) & 
                                      (daily_count_df['sender_email'] == sender_email)]
        if not existing_row.empty:
            # If the sender's entry for today exists, increment the mail_count.
            daily_count_df.loc[(daily_count_df['date'] == today_date) & 
                                (daily_count_df['sender_email'] == sender_email), 'mail_count'] += 1
        else:
            # If the sender's entry for today doesn't exist, create a new one.
            new_row = pd.DataFrame({'date': [today_date], 'sender_email': [sender_email], 'mail_count': [1]})
            daily_count_df = pd.concat([daily_count_df, new_row], ignore_index=True)

    # Save the updated daily_count_df back to the CSV file.
    daily_count_df.to_csv(DAILY_COUNT_FILE, index=False)

def safe_read_excel(file):
    """Safely read an Excel file, returning an empty DataFrame if missing."""
    if os.path.exists(file):
        return pd.read_excel(file)
    return pd.DataFrame()

def safe_read_csv(file):
    """Safely read a CSV file, returning an empty DataFrame if missing."""
    if os.path.exists(file):
        return pd.read_csv(file)
    return pd.DataFrame()

def update_status_in_main_file(email):
    """Update EmailMarketing1Status to 'Y' in the main Excel file for the given email."""
    main_df = pd.read_excel(MAIN_FILE)
    main_df.loc[main_df['EMAIL'] == email, 'EmailMarketing1Status'] = 'Y'
    main_df.to_excel(MAIN_FILE, index=False)

def send_email(sender_email, sender_password, recipient_email, email_content):
    """Send an email using the given sender credentials."""
    try:
        print(f"Sending email from {sender_email} to {recipient_email}...")
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = EMAIL_SUBJECT
        msg.attach(MIMEText(email_content, 'html'))

        smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        smtp.login(sender_email, sender_password)
        smtp.sendmail(sender_email, recipient_email, msg.as_string())
        smtp.quit()

        print(f"Email sent to {recipient_email}.")
        log_email_activity(sender_email, recipient_email)
        return True
    except Exception as e:
        print(f"Failed to send email to {recipient_email}. Error: {e}")
        return False

def get_next_sender(senders_df, daily_count, last_send_time):
    """Get the next eligible sender based on limits and time gaps."""
    current_time = time.time()
    for _, sender in senders_df.iterrows():
        sender_email = sender['Email']
        if daily_count[sender_email] >= MAIL_LIMIT:
            continue
        if last_send_time[sender_email] and (current_time - last_send_time[sender_email] < TIME_GAP):
            continue
        return sender_email, sender['Password']
    return None, None

def main():
    email_content = load_email_content()

    # Load and filter recipients
    main_df = safe_read_excel(MAIN_FILE)
    if main_df.empty:
        print("Main Excel file is empty or missing.")
        return

    filtered_df = main_df[main_df['EmailMarketing1Status'] != 'Y']
    filtered_df[['EMAIL']].to_csv(MAILER_FILE, index=False)

    senders_df = safe_read_csv(SENDER_FILE)
    recipients_df = safe_read_csv(MAILER_FILE)

    if senders_df.empty or recipients_df.empty:
        print("Senders or recipients list is empty.")
        return

    daily_count = {sender['Email']: 0 for _, sender in senders_df.iterrows()}
    last_send_time = {sender['Email']: None for _, sender in senders_df.iterrows()}

    for _, recipient in recipients_df.iterrows():
        recipient_email = recipient['EMAIL']

        sender_email, sender_password = get_next_sender(senders_df, daily_count, last_send_time)
        if not sender_email:
            print("No eligible sender found.")
            break

        if send_email(sender_email, sender_password, recipient_email, email_content):
            daily_count[sender_email] += 1
            last_send_time[sender_email] = time.time()
            update_status_in_main_file(recipient_email)

    log_daily_count(daily_count)

if __name__ == "__main__":
    main()
