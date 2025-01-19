from file_utils import create_output_folder
from email_process import Email_process
import logging

if __name__ == "__main__":

    output_folder = "D:/Emails"
    skip_email_from_top = 0
    number_of_emails_to_process = 100
    omit_senders = [
        "Adobe Acrobat" "Asian Paints",
        "Flipkart",
        "Goibibo",
        "Just Dial",
        "HDFC Bank InstaAlerts",
        "ICICI Prudential Life Insurance",
        "ixigo",
        "LinkedIn",
        "LinkedIn Job Alerts",
        "Lodha",
        "MyGov",
        "mygate",
        "Netflix",
        "Newsletters@yourstory.com",
        "Pepperfry",
        "Saraswat Bank",
        "Shoppers Stop",
        "Wakefit",
    ]

    path_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

    # Configuration for logger
    logging.basicConfig(
        filename="email_to_pdf.log", format="%(asctime)s %(message)s", filemode="w"
    )

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # Creates output folder if not exists
    create_output_folder(output_folder)

    email_obj = Email_process(path_wkhtmltopdf, logger)

    # Sets up win32com outlook object and returns the sorted items
    items = email_obj.setup_process_email()

    logger.info("Email processing started...")
    logger.info("Number of email to process: %s", number_of_emails_to_process)

    # skip latest 4 emails and then process the next 2 emails to pdf
    email_obj.process_email(
        items,
        output_folder,
        skip_email_from_top=skip_email_from_top,
        omit_senders=omit_senders,
        email_count=number_of_emails_to_process,
    )

    logger.info("Email processing completed. Exited Program successfully.")
