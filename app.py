import os
import datetime
import openpyxl
from flask import Flask, request, jsonify
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.encoders import encode_base64
import threading
import uuid

app = Flask(__name__)

# Configurations
#BASE_URL = "afss-staff-login.onrender.com"
BASE_URL = "afss-staff-login.onrender.com"
ADMIN_EMAIL = "cobwebb784@gmail.com"
EXCEL_FILE = "user_logs.xlsx"
OFFICE_MACS = ["14-AB-C5-43-F0-08"]  # Replace with allowed MAC addresses
OFFICE_NETWORK = "192.168.137.1."  # Replace with your office network prefix

# Ensure the Excel file exists
if not os.path.exists(EXCEL_FILE):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Date", "Time", "Name", "Role", "MAC Address"])
    workbook.save(EXCEL_FILE)


def get_user_mac_address():
    """Retrieve the MAC address of the device."""
    return ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff) for elements in range(0, 8 * 6, 8)][::-1])


def send_email_with_excel():
    """Send the Excel file to the admin email."""
    msg = MIMEMultipart()
    msg['From'] = "team.afssltd@gmail.com"
    msg['To'] = ADMIN_EMAIL
    msg['Subject'] = "User Login Logs"

    # Attach file
    with open(EXCEL_FILE, "rb") as file:
        attachment = MIMEBase("application", "octet-stream")
        attachment.set_payload(file.read())
        encode_base64(attachment)
        attachment.add_header("Content-Disposition", f"attachment; filename={EXCEL_FILE}")
        msg.attach(attachment)

    # Email setup
    with smtplib.SMTP("smtp.gmail.com", 587) as server:  # Replace with your SMTP details
        server.starttls()
        server.login("team.afssltd@gmail.com", "lajk rvva ytaa qsxe")  # Replace with your email credentials
        server.send_message(msg)


def schedule_email():
    """Schedule email sending every 59 minutes."""
    threading.Timer(3540, schedule_email).start()  # 59 minutes = 3540 seconds
    send_email_with_excel()


@app.route('/login', methods=['POST'])
def login_user():
    url = request.json.get("url")
    mac_address = get_user_mac_address()
    user_ip = request.remote_addr

    # Validate the URL and Date
    today_date = datetime.datetime.now().strftime("%m-%d")
    if f"{BASE_URL}/{today_date}" != url:
        return jsonify({"error": "Access Denied: Invalid URL for today."}), 403

    # Check MAC address and office network
    if mac_address not in OFFICE_MACS or not user_ip.startswith(OFFICE_NETWORK):
        return jsonify({"error": "Access Denied: Not from the office network or unauthorized MAC."}), 403

    # Save login details to Excel
    name = "John Doe"  # Replace with actual logic to retrieve user details
    role = "Staff"
    now = datetime.datetime.now()
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active
    sheet.append([now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S"), name, role, mac_address])
    workbook.save(EXCEL_FILE)

    return jsonify({"message": "Login Successful", "name": name, "role": role})


# Start scheduled email
schedule_email()

if __name__ == '__main__':
    app.run(debug=True)
