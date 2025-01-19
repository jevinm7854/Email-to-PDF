import os
import re
import datetime


def create_output_folder(output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)


def sanitize_filename(filename):
    """
    Replace invalid filename characters with underscores.
    """
    # Remove or replace unsupported characters
    filename = re.sub(
        r'[<>:"/\\|?*\x00-\x1F]', "_", filename
    )  # Replace invalid characters with '_'
    filename = re.sub(r"[^\x00-\x7F]+", "_", filename)  # Replace non-ASCII characters
    return filename.strip()


def get_email_date(message):
    # Assuming 'message.Date' is a datetime object, adjust if it's in a different format
    email_date = message.ReceivedTime
    if isinstance(email_date, str):  # If Date is a string, parse it
        email_date = datetime.strptime(email_date, "%a, %d %b %Y %H:%M:%S %z")

    # Format date as YYYY-MM-DD
    # return email_date.strftime("%Y-%m-%d")
    return email_date.strftime("%d-%m-%Y")
