import pdfkit
import win32com.client
from datetime import datetime
from file_utils import sanitize_filename, get_email_date
import os
import shutil


class Email_process:
    def __init__(self, path_wkhtmltopdf, logger):
        self.path_wkhtmltopdf = path_wkhtmltopdf
        self.config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
        self.logger = logger

    def html_to_pdf(self, html_content, pdf_file):
        """Convert HTML content to a PDF file."""
        try:
            pdfkit.from_file(html_content, pdf_file, configuration=self.config)
            self.logger.info(f"PDF saved successfully at {pdf_file}")
        except Exception as e:
            self.logger.error(f"Failed to convert HTML to PDF: {e}")

    def skip_emails(self, items, number_of_emails=0):
        message = items.GetFirst()
        for _ in range(number_of_emails):
            message = items.GetNext()
        self.logger.info(f"Skipped {number_of_emails} emails.")
        return message

    def setup_process_email(self):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace(
                "MAPI"
            )
            inbox = outlook.GetDefaultFolder(6)  # 6 refers to the Inbox folder

            # Sort items by received time in descending order (latest first)
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)  # True for descending order
            self.logger.info("Emails sorted by received time in descending order.")
            return items
        except Exception as e:
            self.logger.error(f"Failed to setup process email: {e}")
            return None

    def process_email(self, items, output_folder, skip_email_from_top, email_count=1):
        count = 0
        message = self.skip_emails(items, skip_email_from_top)
        while message and count < email_count:
            if not message:
                self.logger.warning("No emails found.")
                return

            try:
                self.logger.info(
                    f"Working on the email: {message.Subject} dated {message.ReceivedTime}"
                )
                # Sanitize subject for filename

                email_date = get_email_date(message)
                sender_name = sanitize_filename(message.SenderName or "Unknown")[:18]
                subject = sanitize_filename(message.Subject or "Untitled")[:42]

                filename = f"{email_date}-{sender_name} - {subject}"

                html_filename = os.path.join(output_folder, f"{filename}.html")

                # Save the email as an HTML file
                message.SaveAs(html_filename, 5)  # 5 is the format code for HTML
                self.logger.info(f"Saved the email in html as: {html_filename}")
                self.logger.info("Converting to PDF...")
                pdf_filename = os.path.join(output_folder, f"{filename}.pdf")

                # Convert the HTML file to PDF
                self.html_to_pdf(html_filename, pdf_filename)

                self.logger.info("Saved as PDF!")

                os.remove(html_filename)  # Remove the HTML file after conversion
                rm_dir_filename = os.path.join(output_folder, f"{filename}_files")
                shutil.rmtree(rm_dir_filename)
                self.logger.info(
                    f"Removed the HTML file and associated folder: {html_filename}"
                )
                print(f"Completed email number {count} ")
                self.logger.info("----------------------------------------------------")
            except Exception as e:
                self.logger.error(
                    f"Failed to save the email ({message.Subject or 'No Subject'}. Date {message.ReceivedTime}): {e}"
                )
            message = items.GetNext()
            count += 1
