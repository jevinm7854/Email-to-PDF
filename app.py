from file_utils import create_output_folder
from email_process import Email_process
import logging

if __name__ == "__main__":

    output_folder = "C:/Emails"
    skip_email_from_top = 4
    number_of_emails_to_process = 2

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

    # skip latest 4 emails and then process the next 2 emails to pdf
    email_obj.process_email(
        items,
        output_folder,
        skip_email_from_top=skip_email_from_top,
        email_count=number_of_emails_to_process,
    )

    logger.info("Email processing completed. Exited Program successfully.")
