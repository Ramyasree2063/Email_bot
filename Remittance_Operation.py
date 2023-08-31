from selenium.webdriver import ChromeOptions, ActionChains, Keys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium.webdriver.common.alert import Alert
import time
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
def create_remittance(driver, month, year):
    try:
        payroll = driver.find_element(by="xpath", value="//*[@id='payrolla']/p/i")
        driver.execute_script("arguments[0].click();", payroll)

        remittance = driver.find_element(by="xpath", value="//*[@id='payrollmodule']/ul/li[3]/a/p")
        driver.execute_script("arguments[0].click();", remittance)

        add_new = driver.find_element(by="xpath", value="//*[@id='createNewRemittanceButton']")
        driver.execute_script("arguments[0].click();", add_new)

        remittance_report_calendar = driver.find_element(by="xpath", value="//*[@id='Date']")
        remittance_report_calendar.send_keys(month)
        action = ActionChains(driver)
        action.key_down(Keys.ARROW_RIGHT).perform()
        remittance_report_calendar.send_keys(year)
        create_button = driver.find_element(by="xpath", value="//*[@id='createRemittanceReport']")
        create_button.click()
        time.sleep(5)
        alert = Alert(driver)
        alert.accept()
        email_dropdown = driver.find_element(by="xpath", value="//*[@id='dropbtn']")
        driver.execute_script("arguments[0].click();", email_dropdown)
        time.sleep(5)
        email = driver.find_element(by="xpath", value="//*[@id='dropdown1']/li[2]/a")
        driver.execute_script("arguments[0].click();", email)
        time.sleep(5)
        alert = Alert(driver)
        alert.accept()
    except Exception as e:
        print("An error occurred:", str(e))
        email_subject = "Formatting issues"
        email_body = f"The format of your email request does not match, please try the following email formats:\r\ncreate a payroll for the following employees on 10th August 2023:" \
                     f"\r\n\r\nyour employee name1 has 20 worked hours\r\n\r\nyour employee name2 has 40 worked hours\r\nSemimonthly has 10 worked hours\r\n" \
                     f"\r\nyour employee name3 has 20 worked hours\r\n5 holiday hours\r\n2 overtime hours\r\n2 Bonus hours"
        send_notification_email(recipient_email, email_subject, email_body)
        return
