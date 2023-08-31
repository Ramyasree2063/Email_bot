import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

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


def create_invoice(driver, extracted_details):

    try:
        sales = driver.find_element(by="xpath", value="//*[@id='salesa']/p/i").click()
        invoices = driver.find_element(by="xpath",
                                       value="//*[@id='salesmodule']/ul/li[2]/a").click()
        for customer_name, summary, invoice_details in extracted_details:
            customer_name = customer_name
            items = invoice_details
            summary = summary
            print("Customer Name:", customer_name)
            print("Summary:", summary)
            print("Items:")
            for item in items:
                item_name, item_quantity, item_price = item
                print("  - Item:", item_name)
                print("    Price:", item_price)
                print("    Quantity:", item_quantity)
            add_new = driver.find_element(by="xpath",
                                          value="//*[@id='invoiceArchiveButton']").click()
            time.sleep(2)
            add_customer = driver.find_element(by="xpath",
                                               value="//*[@id='addCustomer']").click()
            time.sleep(2)
            try:
                no_of_customers = driver.find_element(by="xpath", value="//*[@id='showCustomer']")
                no_of_customers.click()
                time.sleep(2)
                drop = Select(no_of_customers)
                time.sleep(2)
                drop.select_by_visible_text("100")
                time.sleep(2)
                name = driver.find_element(by="xpath",
                                           value="//*[@id='CustomerTable']//td[text()='" + str(
                                               customer_name) + "']").click()
                time.sleep(2)
            except Exception as e:
                print("An error occurred:", str(e))
                email_subject = "Can not find customer name"
                email_body = f"The following customer was not found in your customers' list: {customer_name}\r\nPlease confirm the input of the customer's name. Or you need to create this customer information on our website\r\nhttps://www.myaccountium.ca/account/login "
                send_notification_email(recipient_email, email_subject, email_body)
                return

            non_matched_items = []
            for item_name, item_quantity, item_price in items:
                # print(item_name)
                add_item = driver.find_element(by="xpath", value="//*[@id='addItem']").click()
                time.sleep(2)
                search = driver.find_element(by="xpath", value="//*[@id='site-search']")
                search.click()
                search.clear()
                time.sleep(2)
                search.send_keys(item_name)
                table_elements = driver.find_elements(by="xpath",
                                                      value="//*[@id='myTable']/tr/td[1]")
                table_texts = [element.get_attribute('textContent') for element in
                               table_elements]
                matching_items = [name for name in table_texts if item_name in name]
                # print(matching_items)
                if len(matching_items) == 1:
                    time.sleep(2)
                    checkbox = driver.find_element(by="xpath",
                                                   value="//*[@id='myTable']/tr/td/input")
                    checkbox.click()
                    add_button = driver.find_element(by="xpath",
                                                     value="//*[text()='Add']").click()
                else:
                    non_matched_items.append((item_name, item_quantity, item_price))
                    cancel_button = driver.find_element(by="xpath",
                                                        value="//*[text()='Add']/following-sibling::button").click()

            # add_button = driver.find_element(by="xpath", value="//*[@id='addItemInvoiceButton']")
            # add_button.click()

            time.sleep(5)

            for index, (item_name, item_quantity, item_price) in enumerate(items):
                matching_items = [name for name in table_texts if item_name in name]
                if len(matching_items) == 1:
                    items[index] = (matching_items[0], item_quantity, item_price)

            # Check non-matched items and add them individually

            for item_name, item_quantity, item_price in non_matched_items:
                quick_add_button = driver.find_element(by="xpath", value="//*[text()='Quick Add']")
                quick_add_button.click()

                name_text = driver.find_element(by="xpath", value="//*[@id='nameProduct']")
                name_text.clear()
                name_text.send_keys(item_name)

                price_text = driver.find_element(by="xpath", value="//*[@id='priceProduct']")
                price_text.clear()
                price_text.send_keys(str(item_price))

                save_button = driver.find_element(by="xpath", value="//*[@id='saveProduct']")
                save_button.click()

            time.sleep(5)

            # Update quantities for all items

            for index, (item, item_quantity, price) in enumerate(items):
                k = index + 1
                quantity = driver.find_element(by="xpath", value="//div[text()='" + str(
                    item) + "']/../../following-sibling::tr[1]/descendant::input")
                quantity.click()
                quantity.clear()
                quantity.send_keys(item_quantity)
                time.sleep(2)

            summary_button = driver.find_element(by="xpath",
                                                 value="//*[@id='headerCard']/div[1]/button").click()
            time.sleep(2)
            summary_input = driver.find_element(by="xpath", value="//*[@id='summary']")

            time.sleep(2)
            driver.execute_script("arguments[0].value = arguments[1];", summary_input, summary)
            time.sleep(5)

            save_continue_button = driver.find_element(by="xpath",
                                                       value="//*[@id='saveContinue']")
            save_continue_button.click()
            time.sleep(5)

            only_save_button = driver.find_element(by="xpath",
                                                   value="//*[@id='mainContainer']/form/div[1]/div/a[2]")
            only_save_button.click()

            time.sleep(5)

            # Handle any alerts or pop-ups
            alert = Alert(driver)
            alert.accept()

            time.sleep(7)

            # Handle any alerts or pop-ups

            alert = Alert(driver)
            alert.accept()
            time.sleep(5)


    except Exception as e:
        print("An error occurred:", str(e))
        return

