import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from selenium.webdriver import ActionChains, Keys
from selenium.webdriver.common.alert import Alert
import time
from selenium.webdriver.support.ui import Select

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

def create_payroll(driver, year, month, day, employee_details):
    try:

        payroll = driver.find_element(by="xpath", value="//*[@id='payrolla']/p/i").click()
        run_payroll = driver.find_element(by="xpath",
                                          value="//*[@id='payrollmodule']/ul/li[2]/a/p").click()
        time.sleep(2)

        add_new = driver.find_element(by="xpath",
                                      value="//*[@id='createNewRunPayrollButton']").click()
        individual = driver.find_element(by="xpath", value="//*[@id='IndividualTab']").click()
        time.sleep(5)
        select_employee = driver.find_element(by="xpath",
                                              value="//*[@id='EmployeesIndividual']")
        select_employee.click()
        drop = Select(select_employee)
        time.sleep(3)
        website_names_list = []
        for option in drop.options:
            name_attribute = option.get_attribute('name')
            if name_attribute:
                website_names_list.append(name_attribute)
        select_employee.click()
        name_hours_list = []
        for employee_name, hours_or_salary, holiday_hours, overtime_hours, bonus_hours in employee_details:
            name_hours_list.append(
                (employee_name, hours_or_salary, holiday_hours, overtime_hours, bonus_hours))
        time.sleep(5)
        print(name_hours_list)
        if not name_hours_list:
            email_subject = "Formatting issues"
            email_body = f"The format of your email request does not match, please try the following email formats:\r\ncreate a payroll for the following employees on 10th August 2023:" \
                         f"\r\n\r\nyour employee name1 has 20 worked hours\r\n\r\nyour employee name2 has 40 worked hours\r\nSemimonthly has 10 worked hours\r\n" \
                         f"\r\nyour employee name3 has 20 worked hours\r\n5 holiday hours\r\n2 overtime hours\r\n2 Bonus hours"
            send_notification_email(recipient_email, email_subject, email_body)
            return

        display_names_list = []
        matched_names_list = []

        for website_name in website_names_list:
            matching_element = None
            website_name_normalized = website_name.strip().lower()
            # Split website name into parts
            website_name_parts = website_name_normalized.split()
            website_first_name = website_name_parts[0]
            website_last_name = website_name_parts[-1]
            for employee_name, worked_hours, holiday_hours, overtime_hours, bonus_hours in name_hours_list:
                employee_name_normalized = employee_name.strip().lower()
                # Check for exact name match or first name match or last name match
                if (
                        employee_name_normalized == website_name_normalized or employee_name_normalized == website_first_name or employee_name_normalized == website_last_name):
                    matching_element = employee_name
                    matched_names_list.append((website_name, worked_hours, holiday_hours,
                                               overtime_hours, bonus_hours))
                    break
            if matching_element is None:
                display_names_list.append(website_name)
        print(matched_names_list)

        for full_name, worked_hours, holiday_hours, overtime_hours, bonus_hours in matched_names_list:
            year = driver.find_element(by="xpath", value="//*[@id='Date']").send_keys(
                year)
            action = ActionChains(driver)
            action.key_down(Keys.ARROW_RIGHT).perform()
            month = driver.find_element(by="xpath", value="//*[@id='Date']").send_keys(
                month)
            action = ActionChains(driver)
            action.key_down(Keys.ARROW_RIGHT).perform()
            date = driver.find_element(by="xpath", value="//*[@id='Date']").send_keys(
                day)
            time.sleep(5)
            select_employee = driver.find_element(by="xpath",
                                                  value="//*[@id='EmployeesIndividual']")
            select_employee.click()
            drop = Select(select_employee)
            time.sleep(3)
            drop.select_by_visible_text(full_name)
            time.sleep(5)
            for option in drop.options:
                if option.text == full_name:
                    id = option.get_attribute('id')
                    option.click()
            worked = driver.find_element(by="xpath", value="//*[@id='employee_" + str(
                id) + "']/div[1]/input")
            worked.clear()
            worked.send_keys(worked_hours)
            holiday = driver.find_element(by="xpath", value="//*[@id='employee_" + str(
                id) + "']/div[2]/input")
            holiday.clear()
            holiday.send_keys(holiday_hours)
            overtime = driver.find_element(by="xpath", value="//*[@id='employee_" + str(
                id) + "']/div[3]/input")
            overtime.clear()
            overtime.send_keys(overtime_hours)
            bonus = driver.find_element(by="xpath", value="//*[@id='employee_" + str(
                id) + "']/div[4]/input")
            bonus.clear()
            bonus.send_keys(bonus_hours)

            time.sleep(3)
            create = driver.find_element(by="xpath",
                                         value="/html/body/div[1]/div[2]/main/div/div/div[2]/button[1]")
            driver.execute_script("arguments[0].click();", create)
            time.sleep(7)
            alert = Alert(driver)
            alert.accept()
            time.sleep(5)
            alert = Alert(driver)
            alert.accept()
            add_new = driver.find_element(by="xpath",
                                          value="//*[@id='createNewRunPayrollButton']").click()
            individual = driver.find_element(by="xpath",
                                             value="//*[@id='IndividualTab']").click()
            time.sleep(5)
        not_found_employees = []
        for employee_name, _, _, _, _ in employee_details:
            name_found = False
            employee_name_normalized = employee_name.strip().lower()
            for website_name in website_names_list:
                website_name_normalized = website_name.strip().lower()

                # Split website name into parts
                website_name_parts = website_name_normalized.split()
                website_first_name = website_name_parts[0]
                website_last_name = website_name_parts[-1]
                if (
                        employee_name_normalized == website_name_normalized or employee_name_normalized == website_first_name or employee_name_normalized == website_last_name):
                    name_found = True
                    break
            if not name_found:
                not_found_employees.append(employee_name)
        print(not_found_employees)
        if not_found_employees:
            # Send email to notify about missing employees
            email_subject = "Missing Employees Notification"
            email_body = f"The following employees were not found in our system: {', '.join(not_found_employees)}"
            f"\r\nYou can try this format to request  :\r\ncreate a payroll for the following employees on 10th August 2023:" \
            f"\r\n\r\nyour employee name1 has 20 worked hours\r\n\r\nyour employee name2 has 40 worked hours\r\nSemimonthly has 10 worked hours\r\n" \
            f"\r\nyour employee name3 has 20 worked hours\r\n5 holiday hours\r\n2 overtime hours\r\n2 Bonus hours"
            print(recipient_email, email_subject, email_body)
            send_notification_email(recipient_email, email_subject, email_body)
    except Exception as e:
        print("An error occurred:", str(e))
        return


