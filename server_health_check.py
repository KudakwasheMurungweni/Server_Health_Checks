from dotenv import load_dotenv
import os
import requests
from bs4 import BeautifulSoup
import openpyxl
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import schedule
import time

# Load environment variables from the .env file
load_dotenv()

# Retrieve sensitive information from environment variables
smtp_server = 'smtp.office365.com'
smtp_port = 587  # Common port for TLS
smtp_user = os.getenv('SMTP_USER')
smtp_password = os.getenv('SMTP_PASSWORD')
from_email = smtp_user
to_email = 'w.mapurisa@delta.co.zw'

# Define server details and credentials
servers = [
    {'chassis': 'LIVE PUREFLEX', 'ip': '10.210.210.2', 'data_url': 'https://10.210.210.2/#/home'},
    {'chassis': 'LIVE PUREFLEX', 'ip': '10.210.210.1', 'data_url': 'https://10.210.210.1/#/home'},
    {'chassis': 'Bay 1 X240 FSM', 'ip': '10.2.10.120', 'data_url': 'https://10.2.10.120/#/home'},
    {'chassis': 'BA2 X240', 'ip': '10.2.10.123', 'data_url': 'https://10.2.10.123/#/home'},
    {'chassis': 'BAY 3-4 X440', 'ip': '10.2.10.163', 'data_url': 'https://10.2.10.163/#/home'},
    {'chassis': 'BAY 5 X220', 'ip': '10.2.10.121', 'data_url': 'https://10.2.10.121/#/home'},
    {'chassis': 'BAY 6 X220', 'ip': '10.2.10.131', 'data_url': 'https://10.2.10.131/#/home'},
    {'chassis': 'BAY 7 – 8 X440', 'ip': '10.2.10.162', 'data_url': 'https://10.2.10.162/#/home'},
    {'chassis': 'BAY 9 X240', 'ip': '10.2.10.161', 'data_url': 'https://10.2.10.161/#/home'},
    {'chassis': 'BAY 10 X240', 'ip': '10.2.10.178', 'data_url': 'https://10.2.10.178/#/home'},
    {'chassis': 'BAY 11 X240', 'ip': '10.2.10.179', 'data_url': 'https://10.2.10.179/#/home'},
    {'chassis': 'BAY 12 X240', 'ip': '10.2.10.147', 'data_url': 'https://10.2.10.147/#/home'},
    {'chassis': 'Node 1 MGMT(New Dense Server - Node 1 ESXi - 10.2.0.21 -', 'ip': '10.2.0.13', 'data_url': 'https://10.2.0.13/#/home'},
    {'chassis': 'Node 2 MGMT(New Dense Server - Node 2 ESXi -', 'ip': '10.2.0.11', 'data_url': 'https://10.2.0.11/#/home'},
    {'chassis': 'New SR850 ESXi Host Client', 'ip': '10.2.0.25', 'data_url': 'https://10.2.0.25/#/home'},
    {'chassis': 'New SAN switch 1', 'ip': '10.2.0.201', 'data_url': 'https://10.2.0.201/#/home'},
    {'chassis': 'New SAN switch 2', 'ip': '10.2.0.202', 'data_url': 'https://10.2.0.202/#/home'},
    {'chassis': 'Bay 1 EN2092 Ethernet Switch', 'ip': '10.2.10.116', 'data_url': 'https://10.2.10.116/#/home'},
    {'chassis': 'Bay 2 EN2092 Ethernet Switch', 'ip': '172.30.30.14', 'data_url': 'https://172.30.30.14/#/home'},
    {'chassis': 'Bay 3 FC3171 San Switch', 'ip': '10.2.10.118', 'data_url': 'https://10.2.10.118/#/home'},
    {'chassis': 'Bay 4 FC3171 San Switch', 'ip': '10.2.10.119', 'data_url': 'https://10.2.10.119/#/home'},
    {'chassis': 'LIVE LENOVO PUREFLEX', 'ip': '10.210.210.1', 'data_url': 'https://10.210.210.1/#/home'},
    {'chassis': 'Bay 1-2 X440', 'ip': '10.2.10.218', 'data_url': 'https://10.2.10.218/#/home'},
    {'chassis': 'Bay 3-4 X440', 'ip': '10.2.10.220', 'data_url': 'https://10.2.10.220/#/home'},
    {'chassis': 'Bay 5 x240', 'ip': '10.2.10.221', 'data_url': 'https://10.2.10.221/#/home'},
    {'chassis': 'Bay 6 x240', 'ip': '10.2.10.188', 'data_url': 'https://10.2.10.188/#/home'},
    {'chassis': 'Bay 7 x240', 'ip': '10.2.10.222', 'data_url': 'https://10.2.10.222/#/home'},
    {'chassis': 'Bay 8 X240', 'ip': '10.2.10.227', 'data_url': 'https://10.2.10.227/#/home'},
    {'chassis': 'Bay 9 x240', 'ip': '10.2.10.236', 'data_url': 'https://10.2.10.236/#/home'},
    {'chassis': 'Bay 10 x240', 'ip': '10.2.10.237', 'data_url': 'https://10.2.10.237/#/home'},
    {'chassis': 'Bay 11 x240', 'ip': '10.2.10.242', 'data_url': 'https://10.2.10.242/#/home'},
    {'chassis': 'Bay 12 x240', 'ip': '10.2.10.226', 'data_url': 'https://10.2.10.226/#/home'},
    {'chassis': 'Bay 13-14 x440', 'ip': '10.2.10.221', 'data_url': 'https://10.2.10.221/#/home'},
    {'chassis': 'Bay 2 EN4091 Ethernet Switch', 'ip': '10.2.10.216', 'data_url': 'https://10.2.10.216/#/home'},
    {'chassis': 'Bay 3 FC3171 San Switch', 'ip': '10.2.10.217', 'data_url': 'https://10.2.10.217/#/home'},
    {'chassis': 'Bay 4 FC3171 San Switch', 'ip': '10.2.10.218', 'data_url': 'https://10.2.10.218/#/home'},
]


def check_server_health_and_send_email():
    # Create or open an Excel file to write the data
    excel_filename = "server_health_status.xlsx"
    
    try:
        # Try to open an existing workbook
        workbook = openpyxl.load_workbook(excel_filename)
        sheet = workbook.active
    except FileNotFoundError:
        # If the file doesn't exist, create a new one
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # Create headers
        headers = ["Chassis Name", "Component IP", "Status", "Comment", "By Who"]
        sheet.append(headers)

    # Iterate through each server configuration
    for server in servers:
        session = requests.Session()
        login_url = f'http://{server["ip"]}/login'  # Adjust based on actual login URL structure
        response = session.post(login_url, data={'username': 'your_username', 'password': 'your_password'})

        if response.status_code == 200:
            response = session.get(server['data_url'])
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')

                # Define the components to check
                components = ['CPU', 'Memory', 'PCI', 'Fan', 'System Board']
                data_to_export = []

                # Check each component's status
                for component in components:
                    status_element = soup.find('div', {'id': component.lower()})
                    
                    if status_element and 'green_tick_class' in status_element['class']:
                        status = "Healthy"
                        comment = ""
                    else:
                        status = "Not Healthy"
                        comment = f"{component} not healthy"  # Adjust this as needed

                    # Append data to list
                    data_to_export.append([
                        server['chassis'],
                        server['ip'],
                        status,
                        comment,
                        "Your Name"  # Replace with your actual name or information
                    ])

                # Append the data to the Excel sheet
                for row in data_to_export:
                    sheet.append(row)

                # Save the workbook
                workbook.save(excel_filename)
                print(f"Data successfully exported to {excel_filename}")

            else:
                print(f"Failed to retrieve data from {server['data_url']}. HTTP Status code: {response.status_code}")
        else:
            print(f"Login failed for {server['ip']}. HTTP Status code: {response.status_code}")

    # Send the email
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = 'Daily Server Health Report'

    # Attach the Excel file
    part = MIMEBase('application', 'octet-stream')
    with open(excel_filename, 'rb') as f:
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={excel_filename}')
    msg.attach(part)

    # Send the email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Upgrade the connection to a secure encrypted SSL/TLS connection
            server.login(smtp_user, smtp_password)
            server.sendmail(from_email, to_email, msg.as_string())
        print(f"Email sent successfully to {to_email}")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Schedule the task to run daily at 8 AM
schedule.every().day.at("08:00").do(check_server_health_and_send_email)

while True:
    schedule.run_pending()
    time.sleep(60)  # Wait a minute before checking for pending tasks
