import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from selenium.webdriver import ActionChains, Keys
from selenium.webdriver.common.alert import Alert
import time




from selenium.webdriver.support.select import Select

SMTP_SERVER = 'smtp.gmail.com'
EMAIL_ADDRESS = 'devgenium@gmail.com'
PASSWORD = 'xdezfdxibnyivhqo'
recipient_email = None

def getRecipient_email(user_email):
    global recipient_email
    recipient_email = user_email
    return recipient_email

def send_notification_email(recipient_email, subject, body):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_ADDRESS, PASSWORD)
        server.sendmail(EMAIL_ADDRESS, recipient_email, msg.as_string())
        server.quit()

        print("Notification email sent successfully")
    except Exception as e:
        print("Failed to send notification email:", str(e))
def create_incomeSatetment(driver, report_range, date_ranges):
    try:
        reports = driver.find_element(by="xpath", value="//*[@id='reportsa']/p/i")
        reports.click()
        income_statement = driver.find_element(by="xpath",
                                               value="//*[@id='reportsmodule']/ul/li[1]/a/p")
        income_statement.click()

        if report_range == "This Fiscal Year":
            year_range = driver.find_element(by="xpath",
                                             value="//*[@id='changeDate_options_div']/select")
            year_range.click()
            time.sleep(2)
            drop = Select(year_range)
            time.sleep(2)
            input_year = 'This Fiscal Year'
            drop.select_by_visible_text(input_year)
            export = driver.find_element(by="xpath", value="//*[@id='printButton'][2]")
            export.click()
        elif report_range == "Previous Fiscal Year":
            year_range = driver.find_element(by="xpath",
                                             value="//*[@id='changeDate_options_div']/select")
            year_range.click()
            time.sleep(2)
            drop = Select(year_range)
            time.sleep(2)
            input_year = 'Previous Fiscal Year'
            drop.select_by_visible_text(input_year)
            export = driver.find_element(by="xpath", value="//*[@id='printButton'][2]")
            export.click()
        elif report_range == "Custom":
            year_range = driver.find_element(by="xpath",
                                             value="//*[@id='changeDate_options_div']/select")
            year_range.click()
            time.sleep(2)
            drop = Select(year_range)
            time.sleep(2)
            input_year = 'Custom'
            drop.select_by_visible_text(input_year)
            from_date = driver.find_element(by="xpath", value="//*[@id='FromDateInput']")
            from_date.send_keys(
                f"{str(date_ranges['from_month']).zfill(2)}/{str(date_ranges['from_day']).zfill(2)}/{str(date_ranges['from_year'])}")
            action = ActionChains(driver)
            action.key_down(Keys.ARROW_RIGHT).perform()
            time.sleep(2)
            to_date = driver.find_element(by="xpath", value="//*[@id='ToDateInput']")
            to_date.send_keys(
                f"{str(date_ranges['to_month']).zfill(2)}/{str(date_ranges['to_day']).zfill(2)}/{str(date_ranges['to_year'])}")
            action = ActionChains(driver)
            action.key_down(Keys.ARROW_RIGHT).perform()
            time.sleep(2)
            #export = driver.find_element(by="xpath", value="//*[@id='printButton'][2]")//*[@id="printButton"]
            export = driver.find_element(by="xpath", value="//*[@id='printButton'][1]")
            export.click()
            printButton = driver.find_element(by="xpath", value="/html/body/print-preview-app//print-preview-sidebar//print-preview-button-strip//div/cr-button[2]")
            printButton().click()

        time.sleep(7)

        download_directory = "C:\\Users\\accountiumvm\\Downloads\\python-3.8.0-embed-amd64"
        file_list = os.listdir(download_directory)

        if file_list:
            newest_file = max(file_list, key=lambda x: os.path.getctime(
                os.path.join(download_directory, x)))
            downloaded_file_path = os.path.join(download_directory, newest_file)
            print("Downloaded file path:", downloaded_file_path)
        else:
            print("No files found in the download directory.")

        sender_email = "devgenium@gmail.com"
        receiver_email = recipient_email  # Make sure recipient_email is defined somewhere
        if report_range == "This Fiscal Year":
            subject = "Income Statement For This Fiscal Year"
        elif report_range == "Previous Fiscal Year":
            subject = "Income Statement For Previous Fiscal Year"
        elif report_range == "Custom":
            subject = "Custom Income Statement"
        else:
            subject = "Income Statement"
        body = "Please find the attached file."

        #create new email object
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject

        # Add message body
        message.attach(MIMEText(body, "plain"))

        # Add attachments
        with open(downloaded_file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={newest_file}")
            message.attach(part)

        # sendmail
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, PASSWORD)
            #server.sendmail(EMAIL_ADDRESS, recipient_email, message.as_string())
            server.quit()

        # Delete downloaded files
        try:
            os.remove(downloaded_file_path)
            print(f"File '{downloaded_file_path}' deleted successfully.")
        except OSError as e:
            print(f"Error deleting file: {e}")
    except Exception as e:
        print("An error occurred:", str(e))
        email_subject = "Formatting issues"
        email_body = f"The format of your email request does not match, please try the following email formats:\r\ncreate income statement for previous year" \
                     f"\r\ncreate income statement for this year"
        #send_notification_email(recipient_email, email_subject, email_body)
        return
