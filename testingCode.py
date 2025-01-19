import os
import re
import win32com.client
from datetime import datetime
import pdfkit

path_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)


def html_to_pdf(html_content, pdf_file):
    """Convert HTML content to a PDF file."""
    try:
        pdfkit.from_file(html_content, pdf_file, configuration=config)
    except Exception as e:
        print(f"Failed to convert HTML to PDF: {e}")


def sanitize_filename(filename):
    """Replace invalid filename characters with underscores."""
    return re.sub(r'[<>:"/\\|?*]', "_", filename)


def get_email_date(message):
    # Assuming 'message.Date' is a datetime object, adjust if it's in a different format
    email_date = message.ReceivedTime
    if isinstance(email_date, str):  # If Date is a string, parse it
        email_date = datetime.strptime(email_date, "%a, %d %b %Y %H:%M:%S %z")

    # Format date as YYYY-MM-DD
    return email_date.strftime("%Y-%m-%d")


# Output folder for saving emails
output_folder = "C:/Emails"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Initialize Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 refers to the Inbox folder

# Sort items by received time in descending order (latest first)
items = inbox.Items
items.Sort("[ReceivedTime]", True)  # True for descending order

message = items.GetFirst()
for _ in range(1):  # Skip the first 4 emails to reach the 5th
    message = items.GetNext()

if message:
    try:
        print(f"Testing the latest email: {message.Subject}")
        # Sanitize subject for filename
        print(f"message :{ message.SenderName}")
        email_date = get_email_date(message)
        sender_name = sanitize_filename(message.SenderName or "Unknown")[:18]
        subject = sanitize_filename(message.Subject or "Untitled")[:42]
        filename = f"{email_date}-{sender_name} - {subject}"
        html_filename = os.path.join(output_folder, f"{filename}.html")

        # Save the email as an HTML file
        message.SaveAs(html_filename, 5)  # 5 is the format code for HTML
        print(f"Saved the latest email as: {html_filename}")
        print("Converting to PDF...")
        pdf_filename = os.path.join(output_folder, f"{filename}.pdf")

        # Convert the HTML file to PDF
        html_to_pdf(html_filename, pdf_filename)

        print("Saved as PDF!")
    except Exception as e:
        print(
            f"Failed to save the latest email ({message.Subject or 'No Subject'}): {e}"
        )
else:
    print("No emails found.")
