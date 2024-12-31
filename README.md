# Email-Automation

import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Email credentials (use placeholders for security)
EMAIL_ADDRESS = "your_email@example.com"  # Replace with your email
EMAIL_PASSWORD = "your_app_password"  # Replace with your app password
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Path to the master Excel file (use a generic name)
MASTER_FILE = "supplier_details.xlsx"  # Ensure the file is in the same directory as this script
OUTPUT_FOLDER = "supplier_files"

# Create output folder if it doesn't exist
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# Function to send email with attachment
def send_email(to_email, supplier_name, total_amount, file_path):
    try:
        # Create the email
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = to_email
        msg['Subject'] = f"Report for {supplier_name}"

        # Email body (generic, without company-specific details)
        body = f"""
Hi {supplier_name},

Please find attached your report for the current period.

Kindly review the details and confirm at your earliest convenience.

Best regards,
Your Name
Your Role - Your Company

📞 Your Contact Number
✉️ {EMAIL_ADDRESS}
🌐 www.yourcompanywebsite.com
        """
        msg.attach(MIMEText(body, 'plain'))

        # Attach the file
        with open(file_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(file_path)}")
            msg.attach(part)

        # Send the email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)

        print(f"Email sent successfully to {supplier_name} ({to_email}).")

    except Exception as e:
        print(f"Error sending email to {supplier_name} ({to_email}): {e}")

# Main function to process the master file and send emails
def process_and_email():
    try:
        # Load the master Excel file
        data = pd.read_excel(MASTER_FILE)
    except Exception as e:
        print(f"Error reading the master file: {e}")
        return

    # Check for required columns (use generic column names)
    required_columns = ['supplier_name', 'to_email', 'total_amount']
    for col in required_columns:
        if col not in data.columns:
            print(f"Error: Missing required column: {col}")
            return

    # Ensure the 'total_amount' column is numeric
    data['total_amount'] = pd.to_numeric(data['total_amount'], errors='coerce')

    # Drop rows with invalid total_amount values
    data = data.dropna(subset=['total_amount'])

    # Calculate total amount per supplier
    total_amount_per_supplier = data.groupby('supplier_name', as_index=False).agg({'total_amount': 'sum', 'to_email': 'first'})

    # Send emails to each supplier
    for _, row in total_amount_per_supplier.iterrows():
        supplier_name = row['supplier_name']
        to_email = row['to_email']
        total_amount = row['total_amount']

        # Filter data for the supplier
        supplier_data = data[data['supplier_name'] == supplier_name]
        file_name = f"{supplier_name.replace(' ', '_')}_report.xlsx"
        file_path = os.path.join(OUTPUT_FOLDER, file_name)

        # Save supplier-specific file
        supplier_data.to_excel(file_path, index=False)

        # Send email with the file attached
        send_email(to_email, supplier_name, total_amount, file_path)

# Run the script
if __name__ == "__main__":
    process_and_email()
