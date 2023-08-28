import threading
import tkinter as tk
import pandas as pd
import pyodbc
from tkinter import filedialog
from configparser import ConfigParser
from tkinter import ttk
from tkcalendar import DateEntry  # Make sure to install this library
import schedule
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import re
from datetime import datetime, timedelta
from queue import Queue
import time
import tkinter as tk
from tkinter import ttk, filedialog
from io import StringIO
import base64
import smtplib


server_entry = None
database_entry = None
username_entry = None
password_entry = None
smtp_server_entry = None
smtp_username_entry = None
smtp_password_entry = None
smtp_from_entry = None
to_email_entry = None
time_entry = None

# Define config globally
config = ConfigParser()
# Function to save both SQL Server and SMTP details to config.ini file
def save_config(config_window):
    config['DATABASE'] = {
        'server': base64.b64encode(server_entry.get().encode()).decode(),
        'database': base64.b64encode(database_entry.get().encode()).decode(),
        'username': base64.b64encode(username_entry.get().encode()).decode(),
        'password': base64.b64encode(password_entry.get().encode()).decode()
    }

    config['SMTP'] = {
        'server': base64.b64encode(smtp_server_entry.get().encode()).decode(),
        'username': base64.b64encode(smtp_username_entry.get().encode()).decode(),
        'password': base64.b64encode(smtp_password_entry.get().encode()).decode(),
        'from': base64.b64encode(smtp_from_entry.get().encode()).decode(),
        'to': base64.b64encode(to_email_entry.get().encode()).decode(),
        'time': time_entry.get()
    }

    with open('config.ini', 'w') as configfile:
        config.write(configfile)

    status_label.config(text='Configuration saved successfully!', fg='green')
    config_window.destroy()





# Function to schedule the report generation and email sending
time_format = re.compile("^([0-1]?[0-9]|2[0-3]):[0-5][0-9](:[0-5][0-9])?$")

def schedule_report():
    # Use the global config object
    global config
    config.read('config.ini')
    time_to_send = config.get('SMTP', 'time')
    if not time_format.match(time_to_send):
        status_queue.put(('Invalid time format! Please enter time in HH:MM or HH:MM:SS format.', 'red'))
        return

    # Compute start and end date strings
    now = datetime.now()
    yesterday = now - timedelta(days=1)
    start_date_time = datetime(yesterday.year, yesterday.month, yesterday.day)
    end_date_time = datetime(yesterday.year, yesterday.month, yesterday.day, 23, 59, 59)
    start_date_time_str = start_date_time.strftime('%Y-%m-%d %H:%M:%S')
    end_date_time_str = end_date_time.strftime('%Y-%m-%d %H:%M:%S')

    # Compute start and end date strings for the arguments
    start_date_str = start_date_time.strftime('%Y-%m-%d')
    end_date_str = end_date_time.strftime('%Y-%m-%d')
    start_time_str = start_date_time.strftime('%H:%M:%S')
    end_time_str = end_date_time.strftime('%H:%M:%S')

    # Schedule the task
    schedule.every().day.at(time_to_send).do(lambda: send_report(generate_report(start_date_str, start_time_str, end_date_str, end_time_str, start_date_time_str, end_date_time_str, save_to_file=False), start_date_time_str, end_date_time_str))
# Create a queue for status updates
status_queue = Queue()

# Define window as a global variable
window = None

# def update_status_label():
#     # Access the global window variable
#     global window
#     # Check if there are any new status updates in the queue
#     if not status_queue.empty():
#         # Get the latest status update and configure the status label
#         text, color = status_queue.get()
#         status_label.config(text=text, fg=color)
#     # Schedule this function to be called again after 100ms
#     window.after(100, update_status_label)


# Function to generate the report
def generate_report(start_date, start_time, end_date, end_time, start_date_time_str, end_date_time_str, save_to_file=True):
    print(f"Value of save_to_file: {save_to_file}")
    try:
        config = ConfigParser()
        config.read('config.ini')
        server = base64.b64decode(config.get('DATABASE', 'server').encode()).decode()
        database = base64.b64decode(config.get('DATABASE', 'database').encode()).decode()
        username = base64.b64decode(config.get('DATABASE', 'username').encode()).decode()
        password = base64.b64decode(config.get('DATABASE', 'password').encode()).decode()


        if username and password:
            connection_string = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
        else:
            connection_string = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes"

        connection = pyodbc.connect(connection_string)

        start_date_time = pd.to_datetime(start_date + ' ' + start_time)
        end_date_time = pd.to_datetime(end_date + ' ' + end_time)

        start_date_time_str = start_date_time.strftime('%Y-%m-%d %H:%M:%S')
        end_date_time_str = end_date_time.strftime('%Y-%m-%d %H:%M:%S')


        # Modify the query to include the date range filter
        query = f"""
            SELECT 
    TH.Customer AS "A/C Number",
    TL.Line AS "Line Number",
    C.LastName AS "Customer Name",
    TH.Receipt AS "Invoice Number",
    CONVERT(DATE, TH.Logged) AS "Invoice Date",
    I.SKU,
    I.Description,
    TL.Quantity,
    I.Unit AS UOM,
    TL.Price AS "Unit Price",
    TL.SubAfterTax AS Amount,
    SUM(TL.SubAfterTax) * 0.15 AS GST,
    SUM(TL.SubAfterTax) + SUM(TL.SubAfterTax) * 0.15 AS Gross
FROM 
    TransHeaders TH
JOIN 
    Customers C ON TH.Customer = C.Code
JOIN 
    TransLines TL ON TH.TransNo = TL.TransNo AND TH.Branch = TL.Branch AND TH.Station = TL.Station
JOIN 
    Items I ON TL.UPC = I.UPC
WHERE 
    TH.Logged BETWEEN '{start_date_time_str}' AND '{end_date_time_str}'
GROUP BY 
    TH.Customer,
    TH.Receipt,
    TL.Line,
    C.LastName,
    CONVERT(DATE, TH.Logged),
    I.SKU,
    I.Description,
    TL.Quantity,
    I.Unit,
    TL.Price,
    TL.SubAfterTax
ORDER BY 
    TH.Customer,
    C.LastName,
    TH.Receipt,
    TL.Line


        """
        df = pd.read_sql_query(query, connection)

        #df['Invoice Number'] = "'" + df['Invoice Number'] + "'"

        # Create a new DataFrame for the modified report
        modified_df = pd.DataFrame(columns=df.columns)

        if save_to_file:
            file_path = filedialog.asksaveasfilename(defaultextension='.csv')
            if file_path:
                df.to_csv(file_path, index=False)
                status_label.config(text='Report generated and saved successfully!', fg='green')
            else:
                status_label.config(text='Report generation cancelled.', fg='red')

        return df

        connection.close()

    except Exception as e:
        status_queue.put((f'Error: {str(e)}', 'red'))

def send_report(df, start_date_time_str, end_date_time_str):
    try:
        # Create a StringIO object and save the DataFrame to it
        csv_buffer = StringIO()
        df.to_csv(csv_buffer, index=False)

        # Send the report as an email attachment
        msg = MIMEMultipart()

        # Read SMTP details from config.ini file
        config = ConfigParser()
        config.read('config.ini')
        smtp_server = base64.b64decode(config.get('SMTP', 'server').encode()).decode()
        smtp_username = base64.b64decode(config.get('SMTP', 'username').encode()).decode()
        smtp_password = base64.b64decode(config.get('SMTP', 'password').encode()).decode()
        smtp_from = base64.b64decode(config.get('SMTP', 'from').encode()).decode()
        to_email = base64.b64decode(config.get('SMTP', 'to').encode()).decode().split(',')


        msg['From'] = smtp_from
        msg['To'] = to_email
        msg['Subject'] = f'Daily Report from {start_date_time_str} to {end_date_time_str}'

        body = 'Please find the daily report attached.'
        msg.attach(MIMEText(body, 'plain'))

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(csv_buffer.getvalue())
        encoders.encode_base64(part)

        part.add_header('Content-Disposition', f'attachment; filename= Daily_Report_{start_date_time_str}_to_{end_date_time_str}.csv')

        msg.attach(part)

        with smtplib.SMTP_SSL(smtp_server, 465) as server:
            server.login(smtp_username, smtp_password)
            for to_address in to_email:
                to_address = to_address.strip()
                
                msg = MIMEMultipart()
                msg['From'] = smtp_from
                msg['To'] = to_address
                msg['Subject'] = f'Daily Report from {start_date_time_str} to {end_date_time_str}'
                msg.attach(MIMEText(body, 'plain'))
                msg.attach(part)

                server.send_message(msg)




    except Exception as e:
        status_queue.put((f'Error: {str(e)}', 'red'))

def start_scheduler():
    while True:
        schedule.run_pending()
        time.sleep(1)

# Function to schedule the report and start the scheduler
def schedule_and_start():
    schedule_report()
    scheduler_thread = threading.Thread(target=start_scheduler)
    scheduler_thread.start()

if __name__ == "__main__":
    schedule_and_start()


def open_config_window():
    global server_entry, database_entry, username_entry, password_entry
    global smtp_server_entry, smtp_username_entry, smtp_password_entry
    global smtp_from_entry, to_email_entry, time_entry

    config_window = tk.Toplevel(window)
    config_window.title('Update Configuration')

    # All your Labels and Entries come here

    # Server Label and Entry
    server_label = tk.Label(config_window, text='Server:')
    server_label.pack()
    server_entry = tk.Entry(config_window)
    server_entry.pack()

    # Database Label and Entry
    database_label = tk.Label(config_window, text='Database:')
    database_label.pack()
    database_entry = tk.Entry(config_window)
    database_entry.pack()

    # Username Label and Entry
    username_label = tk.Label(config_window, text='Username (leave blank for Windows Authentication):')
    username_label.pack()
    username_entry = tk.Entry(config_window)
    username_entry.pack()

    # Password Label and Entry
    password_label = tk.Label(config_window, text='Password (leave blank for Windows Authentication):')
    password_label.pack()
    password_entry = tk.Entry(config_window, show='*')
    password_entry.pack()

    # SMTP Server Label and Entry
    smtp_server_label = tk.Label(config_window, text='SMTP Server:')
    smtp_server_label.pack()
    smtp_server_entry = tk.Entry(config_window)
    smtp_server_entry.pack()

    # SMTP Username Label and Entry
    smtp_username_label = tk.Label(config_window, text='SMTP Username:')
    smtp_username_label.pack()
    smtp_username_entry = tk.Entry(config_window)
    smtp_username_entry.pack()

    # SMTP Password Label and Entry
    smtp_password_label = tk.Label(config_window, text='SMTP Password:')
    smtp_password_label.pack()
    smtp_password_entry = tk.Entry(config_window, show='*')
    smtp_password_entry.pack()

    # 'From' Email Address Label and Entry
    smtp_from_label = tk.Label(config_window, text="'From' Email Address:")
    smtp_from_label.pack()
    smtp_from_entry = tk.Entry(config_window)
    smtp_from_entry.pack()

    # 'To' Email Address Label and Entry
    to_email_label = tk.Label(config_window, text="'To' Email Address:")
    to_email_label.pack()
    to_email_entry = tk.Entry(config_window)
    to_email_entry.pack()

    # Time to Send Report Label and Entry
    time_entry_label = tk.Label(config_window, text="Time to Send Report (HH:MM):")
    time_entry_label.pack()
    time_entry = tk.Entry(config_window)
    time_entry.pack()

    # Save SMTP Config Button
    # Pass the config_window to the save_config function
    save_config_button = tk.Button(config_window, text='Save Config', command=lambda: save_config(config_window))
    save_config_button.pack()



# Create the main application window
window = tk.Tk()
window.title('Sanford Limited Report')
window.geometry('400x400')

# # Start the status label update function
# update_status_label()



start_date_label = ttk.Label(window, text="Start date:")
start_date_label.pack()
start_date_entry = DateEntry(window)
start_date_entry.pack()

start_time_label = tk.Label(window, text="Start time (HH:MM):")
start_time_label.pack()
start_time_entry = ttk.Entry(window)
start_time_entry.pack()

end_date_label = ttk.Label(window, text="End date:")
end_date_label.pack()
end_date_entry = DateEntry(window)
end_date_entry.pack()

end_time_label = tk.Label(window, text="End time (HH:MM):")
end_time_label.pack()
end_time_entry = ttk.Entry(window)
end_time_entry.pack()

generate_report_button = tk.Button(window, text='Generate Report', command=lambda: generate_report(
    start_date_entry.get(), 
    start_time_entry.get(), 
    end_date_entry.get(), 
    end_time_entry.get(), 
    datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 
    (datetime.now() - timedelta(seconds=1)).strftime('%Y-%m-%d %H:%M:%S')
))
generate_report_button.pack()

# 'Update Configuration' Button
update_config_button = tk.Button(window, text='Update Configuration', command=open_config_window)
update_config_button.pack()


# Create status label
status_label = ttk.Label(window)
status_label.pack()

# # Schedule Email Button
# schedule_email_button = tk.Button(window, text='Schedule Email', command=schedule_and_start)
# schedule_email_button.pack()

# Status Label
status_label = tk.Label(window, text='')
status_label.pack()

# Start the GUI event loop
window.mainloop()
