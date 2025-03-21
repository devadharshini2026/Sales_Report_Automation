Sales Report Automation

📌 Project Overview

Sales Report Automation is a Python-based project designed to automate the process of generating and emailing sales reports. This project processes sales data, generates an Excel report with charts, and sends it via email, saving time and improving efficiency.

🚀 Features

Automated Data Processing: Reads and analyzes sales data from an Excel file.

Report Generation: Creates a well-structured Excel report with charts using OpenPyXL.

Email Automation: Sends the generated report via email using SMTP.

Error Handling & Logging: Ensures smooth execution and logs errors for debugging.

🛠️ Technologies Used

Python 3

OpenPyXL (for Excel report generation)

Matplotlib (for data visualization)

smtplib & email (for email automation)

📂 Project Structure

Sales_Report_Automation/
│-- sales_data.xlsx            # Raw sales data
│-- sales_report.py            # Script for generating sales reports
│-- email_report.py            # Script for sending reports via email
│-- output/                    # Stores generated reports and charts
│   ├── sales_report.xlsx
│   ├── sales_chart.png
│-- README.md                  # Project documentation

🔧 Installation & Setup

Clone the repository:

git clone https://github.com/devadharshini2026/Sales_Report_Automation.git
cd Sales_Report_Automation

Install dependencies:

pip install openpyxl matplotlib

Set up email configuration (inside email_report.py):

Replace sender email and password.

Ensure SMTP settings are configured properly.

▶️ How to Run

Generate Sales Report:

python sales_report.py

Send Report via Email:

python email_report.py

Run both scripts together:

python sales_report.py && python email_report.py

📧 Email Configuration

Ensure Less Secure Apps is enabled in your Gmail settings (or use an App Password).

Use smtplib to configure SMTP settings for sending emails.

🏆 Contribution

Feel free to fork this repository and submit pull requests for improvements! 😊

🔹 Developed by Dharshini | GitHub: @devadharshini2026
