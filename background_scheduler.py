import schedule
import time
from sandfordReport import generate_report, send_report
from configparser import ConfigParser
from datetime import datetime, timedelta
import base64

config = ConfigParser()
config.read('config.ini')

# Decrypt the necessary fields
smtp_password = base64.b64decode(config.get('SMTP', 'password').encode()).decode()
time_to_send = config.get('SMTP', 'time')

def send_email():
    now = datetime.now()
    start_date_time = datetime(now.year, now.month, now.day - 1, 0, 0, 0)  # Set time to midnight
    end_date_time = datetime(now.year, now.month, now.day - 1, 23, 59, 59)  # Set time to 23:59:59

    start_date_time_str = start_date_time.strftime('%Y-%m-%d %H:%M:%S')
    end_date_time_str = end_date_time.strftime('%Y-%m-%d %H:%M:%S')

    # Generate the report and get the DataFrame
    df = generate_report(start_date_time_str, "00:00:00", end_date_time_str, "23:59:59", save_to_file=False)

    # Send the report by email
    send_report(df, start_date_time_str, end_date_time_str)


# Schedule the report generation and email sending
schedule.every().day.at(time_to_send).do(send_email)

# Keep the script running
while True:
    schedule.run_pending()
    time.sleep(1)
