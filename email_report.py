import smtplib
from email.message import EmailMessage

# Sender and Receiver Email (Both are the same in your case)
sender_email = "dharshi2.24@gmail.com"
receiver_email = "dharshi2.24@gmail.com"
sender_password = "fwoj rxkl mfbn lddf"

# Create Email Message
msg = EmailMessage()
msg["Subject"] = "Sales Report"
msg["From"] = sender_email
msg["To"] = receiver_email
msg.set_content("Please find the attached sales report.")

# Attach File (Sales Report)
with open("C:\\Excel_Projects\\Sales_Report_Automation\\output\\sales_report.xlsx", "rb") as file:
    msg.add_attachment(file.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="sales_report.xlsx")

# Send Email using SMTP (Port 587 with TLS)
try:
    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.starttls()  # Upgrade to secure connection
        smtp.login(sender_email, sender_password)
        smtp.send_message(msg)
        print("✅ Email Sent Successfully!")
except Exception as e:
    print(f"❌ Error Sending Email: {e}")
