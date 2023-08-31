import imaplib
from email.parser import BytesParser
from email.utils import parseaddr
import json
import os
import re
from datetime import datetime
from selenium.webdriver import EdgeOptions, Edge

import Remittance_Operation
import Invoice_Operation
import Payroll_Operation
import T4_Operation
import IncomeStatement_Operation
import BalanceSheet_Operation

import requests as requests

# Existing keywords
existing_keywords = ['remittance', 'create', 'send', 'prepare', 'invoice', 'payroll', 't4', 'balance', 'sheet', 'income', 'statement']
#Add new parameters
date_list = []
employee_details = []
items = []
customer_name = None
driver = None
keywords = []
def extract_date(email_content):
    # Define a dictionary to map month names to numbers
    month_dict = {"Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
                  "May": "05", "Jun": "06", "Jul": "07", "Aug": "08",
                  "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12"}

    # Extract the date using regular expression
    date_pattern = r'\b(\d{1,2})(?:st|nd|rd|th)?\s(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s(\d{4})\b'
    date_match = re.search(date_pattern, email_content)
    if date_match:
        day, month, year = date_match.groups()
        # Use the month_dict to get the month number
        month = month_dict[month]
        # Use datetime to parse the date
        date = datetime.strptime(f'{day} {month} {year}', '%d %m %Y')
    else:
        date = None
    return date
#for remittance
def extract_month_year(email_content):
    month_regex = r'(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)'
    year_regex = r'\b\d{4}\b'
    date_regex = r'\b(\d{1,2})\s(?:st|nd|rd|th)?\s(?:January|February|March|April|May|June|July|August|September|October|November|December)\s\d{4}\b'

    month_matches = re.findall(month_regex, email_content, re.IGNORECASE)
    year_matches = re.findall(year_regex, email_content, re.IGNORECASE)

    date = [month_matches, year_matches]
    return date


#for Invoice
def extract_details_from_email(email_body):
    invoices = email_body.split('TOTAL')

    extracted_details = []

    for invoice in invoices:
        customer_name, summary, invoice_details = extract_invoice_details(invoice)
        if customer_name and invoice_details:
            extracted_details.append((customer_name, summary, invoice_details))

    print(extracted_details)
    return extracted_details


def extract_invoice_details(invoice_text):
    # Check if the invoice is in tabular format
    if "QTY IN LF" in invoice_text:
        return extract_details_from_tabular(invoice_text)
    else:
        return extract_details_from_text(invoice_text)


def extract_details_from_tabular(invoice_text):
    pattern = r'^\s*(.+?)\s+(\d+)\s+(\d+\.\d+)\s+\$'
    matches = re.findall(pattern, invoice_text, re.MULTILINE)
    invoice_details = [(item.strip(), qty, price) for item, qty, price in matches if item.strip()]
    customer_name = extract_customer_name(invoice_text)
    summary = extract_summary(invoice_text)
    return customer_name, summary, invoice_details


def extract_details_from_text(invoice_text):
    item_details_pattern = r"Add\s+([^\n]+)\s+price\s+\$\s*([\d,.]+)\s+qty\s+(\d+)"
    matches = re.findall(item_details_pattern, invoice_text, re.IGNORECASE)
    invoice_details = []

    for item_match in matches:
        item_name = item_match[0].strip()
        item_price = float(item_match[1].replace(',', ''))
        item_quantity = int(item_match[2])
        invoice_details.append((item_name, item_quantity, item_price))

    customer_name = extract_customer_name(invoice_text)
    summary = extract_summary(invoice_text)
    return customer_name, summary, invoice_details


def extract_customer_name(invoice_text):
    lines = invoice_text.splitlines()
    customer_name = None
    non_empty_line_count = 0

    for line in lines:
        if line.strip():
            non_empty_line_count += 1
            if non_empty_line_count == 2:
                customer_name = line.strip('*').strip()  # Remove asterisks if present
                break
            elif non_empty_line_count == 1 and "invoice for" in line.lower():
                customer_name = line.split("invoice for", 1)[1].strip('*').strip()
                break

    return customer_name


def extract_summary(invoice_text):
    phone_number_pattern = r"Office:\s*\(\d+\)\s*\d+-\d+"
    phone_number_match = re.search(phone_number_pattern, invoice_text)

    if phone_number_match:
        phone_number_index = phone_number_match.end()
        table_index = invoice_text.index("ITEM")
        summary = invoice_text[phone_number_index:table_index].lstrip('*').replace('\r\n', ' ').strip()
        return summary
    else:
        return None


#for Income Statement
def extract_date_and_report_range(email_content):
    # Extract the report range and date range using regular expressions
    pattern = r'(?i)(Custom|from) (\d{4})/(\d{2})/(\d{2}) to (\d{4})/(\d{2})/(\d{2})'
    pattern1 = r'(?i)(to|upto) (\d{4})/(\d{2})/(\d{2})'
    match = re.search(pattern, email_content, re.IGNORECASE)
    match1 = re.search(pattern1, email_content, re.IGNORECASE)

    if match:
        report_range = "Custom"
        from_year, from_month, from_day, to_year, to_month, to_day = match.groups()[1:]
        date_ranges = {
            'from_year': int(from_year),
            'from_month': int(from_month),
            'from_day': int(from_day),
            'to_year': int(to_year),
            'to_month': int(to_month),
            'to_day': int(to_day)
        }
    elif match1:
        report_range = "Custom"
        to_year, to_month, to_day = match1.groups()[1:]
        date_ranges = {
            'from_year': None,
            'from_month': None,
            'from_day': None,
            'to_year': int(to_year),
            'to_month': int(to_month),
            'to_day': int(to_day)
        }
    elif any(keyword.lower() in email_content.lower() for keyword in ["this year", "current year", "this fiscal year"]):
        report_range = "This Fiscal Year"
        date_ranges = None
    elif any(keyword.lower() in email_content.lower() for keyword in
             ["previous year", "year before", "previous fiscal year"]):
        report_range = "Previous Fiscal Year"
        date_ranges = None
    else:
        report_range = None
        date_ranges = None

    return report_range, date_ranges

# For T4
def extract_date_and_names(email_content):
    # Extract the date using regular expression
    date_pattern = r'(?:\bthis year\b|\bprevious year\b|\d{4})'
    matches = re.findall(date_pattern, email_content, re.IGNORECASE)

    if matches:
        date_str = matches[0]

        if date_str.lower() == "this year":
            date = datetime.now().year
        elif date_str.lower() == "previous year":
            date = datetime.now().year - 1
        else:
            date = int(date_str)
    else:
        date = None

    # Extract employee names or check if "all" is present in the email content
    names_pattern = r'(?:(?<=\n\n)|(?<=\n))(?!.*\b(?:this year|previous year|\d{4}|all)\b)(.*)'
    names_match = re.search(names_pattern, email_content, re.IGNORECASE | re.DOTALL)
    if names_match:
        names = re.findall(r'[A-Z][a-z]*(?: [A-Z][a-z]*)?', names_match.group(1))
    else:
        names = []

    # Check if "all" is present in the email content
    if re.search(r'\ball\b', email_content, re.IGNORECASE):
        names.append('all')

    return date, names

def extract_employee_details(email_body):
    employee_details = []
    lines = email_body.strip().split('\n')
    current_employee = []

    for line in lines:
        line = line.strip()
        if line:
            match = re.search(r'^(.*?) has', line)
            if match:
                if current_employee:
                    employee_details.append(current_employee)

                employee_name = match.group(1)
                current_employee = [employee_name, 0, 0, 0, 0]  # [name, worked, holiday, overtime, bonus]

                # Extract worked hours from the same line
                worked_hours_match = re.search(r'(\d+) worked hours', line)
                if worked_hours_match:
                    current_employee[1] = int(worked_hours_match.group(1))
            elif current_employee and 'holiday hours' in line:
                current_employee[2] = int(re.search(r'(\d+)', line).group(1))
            elif current_employee and 'overtime hours' in line:
                current_employee[3] = int(re.search(r'(\d+)', line).group(1))
            elif current_employee and 'Bonus hours' in line:
                current_employee[4] = int(re.search(r'(\d+)', line).group(1))

    if current_employee:
        employee_details.append(current_employee)
    print(employee_details)
    return employee_details

#get information for invoice
def extract_details_for_invoice(email_body):
    # Extract customer name
    global customer_name
    customer_name_pattern = r"for\s+(.+?)\r"
    customer_name_match = re.search(customer_name_pattern, email_body, re.IGNORECASE)
    customer_name = customer_name_match.group(1).strip() if customer_name_match else None
    print(customer_name)
    # Extract item names, prices, and quantities
    item_details_pattern = r"Add\s+([^\n]+)\s+price\s+\$\s*([\d.]+)\s+qty\s+(\d+)"
    item_details_matches = re.findall(item_details_pattern, email_body, re.IGNORECASE)


    for item_match in item_details_matches:
        item_name = item_match[0].strip()
        item_price = float(item_match[1])
        item_quantity = int(item_match[2])
        items.append((item_name, item_price, item_quantity))

    return customer_name, items

def extract_T4(email_content):
    # Extract the date using regular expression
    date_pattern = r'(?:\bthis year\b|\bprevious year\b|\d{4})'
    matches = re.findall(date_pattern, email_content, re.IGNORECASE)

    if matches:
        date_str = matches[0]

        if date_str.lower() == "this year":
            date = datetime.now().year
        elif date_str.lower() == "previous year":
            date = datetime.now().year - 1
        else:
            date = int(date_str)
    else:
        date = None

    # Extract employee names or check if "all" is present in the email content
    names_pattern = r'(?:(?<=\n\n)|(?<=\n))(?!.*\b(?:this year|previous year|\d{4}|all)\b)(.*)'
    names_match = re.search(names_pattern, email_content, re.IGNORECASE | re.DOTALL)
    if names_match:
        names = re.findall(r'[A-Z][a-z]*(?: [A-Z][a-z]*)?', names_match.group(1))
    else:
        names = []

    # Check if "all" is present in the email content
    if re.search(r'\ball\b', email_content, re.IGNORECASE):
        names.append('all')

    return date, names

def delete_json_file(file_path):
    try:
        os.remove(file_path)
        print(f"Deleted file: {file_path}")
    except OSError as e:
        print(f"Error deleting file: {file_path} - {e}")
def reset_globals():
    global date_list, employee_details, items, customer_name, driver, keywords
    date_list = []
    employee_details = []
    items = []
    customer_name = None
    driver = None
    keywords = []


def login_to_accountium(userAccount, userPassword):

    edge_options = EdgeOptions()
    edge_options.use_chromium = True
    #edge_options.add_argument("--headless")
    driver = Edge(options=edge_options)
    driver.maximize_window()
    driver.set_page_load_timeout(20)

    # Get URL
    driver.get("https://www.myaccountium.ca/account/loginselenium")
    driver.implicitly_wait(10)
    driver.maximize_window()
    driver.set_page_load_timeout(20)

    # Login
    email_element = driver.find_element(by="xpath", value="//*[@id='email']")
    email_element.click()
    email_element.send_keys(userAccount)
    password_element = driver.find_element(by="xpath", value="//*[@id='password']")
    password_element.click()
    password_element.send_keys(userPassword)
    login = driver.find_element(by="xpath", value="/html/body/div/div/div/form/button")
    driver.execute_script("arguments[0].click();", login)
    return driver


def logout(driver):
    driver.quit()
# set system email
IMAP_SERVER = 'imap.gmail.com'
SMTP_SERVER = 'smtp.gmail.com'
EMAIL_ADDRESS = 'devgenium@gmail.com'
PASSWORD = 'xdezfdxibnyivhqo'
# connect IMAP service
imap_server = imaplib.IMAP4_SSL(IMAP_SERVER)
imap_server.login(EMAIL_ADDRESS, PASSWORD)
imap_server.select('INBOX')

# get unread list
response, email_ids = imap_server.search(None, 'UNSEEN')
email_ids = email_ids[0].split()

# Iterate over each unread email
for email_id in email_ids:
    response, email_data = imap_server.fetch(email_id, '(RFC822)')
    msg = BytesParser().parsebytes(email_data[0][1])
    # extract sender email
    sender_name, sender_email = parseaddr(msg['From'])
    print(sender_name, sender_email)

    # API endpoint URL
    email_api = "https://www.myaccountium.ca/api/settings/checkEmailBotReceiver"
    api_url = "https://www.myaccountium.ca/api/getpasswordhash"

    # Create the authorization header dictionary
    authorization_header = {
        "role": "bot"
    }

    # Convert the dictionary to JSON
    authorization_header_json = json.dumps(authorization_header)

    # Make the API request with the authorization header
    headers = {
        "Authorization": authorization_header_json
    }
    headers2 = {
        'Content-Type': 'application/json'
    }
    email_data = json.dumps(sender_email)
    email_response = requests.post(email_api, data=email_data, headers=headers2)
    print("Response Status Code:", email_response.status_code)
    print("Response Content:", email_response.content)
    response_data = email_response.json()
    if response_data:
        login_email = response_data[0].get('loginEmail')
    else:
        print("Response data is empty")

    # Check if the sender's email exists in the client credentials dictionary
    if email_response.status_code == 200 and response_data and isinstance(response_data, list) and len(
            response_data) > 0 and login_email:

        username = login_email
        response = requests.post(api_url, json={"Email": username}, headers=headers, verify=False)
        data = response.json()

        if response.status_code == 200 and data["success"]:
            password_hash = data["data"]
        else:
            # Handle error case when password retrieval fails
            print("Failed to retrieve password from the API")
            print("Error message:", data.get("message"))

        # Append "0x" to the beginning of the password_hash
        password_hash_with_prefix = "0x" + password_hash[:16]
        print(password_hash_with_prefix)

        if sender_email == login_email:
            client = [sender_email, password_hash_with_prefix]
            userAccount = client[0]
            print(userAccount)
            userPassword = client[1]
            print(userPassword)

            # Extract email body
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == 'text/plain':
                        body = part.get_payload(decode=True).decode('utf-8')
                        break
            else:
                body = msg.get_payload(decode=True).decode('utf-8')

            # Perform POS tagging and extract keywords
            words = body.split()
            # Iterate over the words
            for word in words:
                # Convert the word and the keywords to lowercase for case-insensitive matching
                if word.lower() in (keyword.lower() for keyword in existing_keywords):
                    keywords.append(word.lower())

            # Extract month, day and year from email body using regular expressions
            date_details = extract_date(body) or None
            print(date_details)
            # If date_details is not empty, extract the month, day and year
            if date_details:
                # Extracted month
                month = date_details.month
                # Extracted year
                year = date_details.year
                # Extracted day
                day = date_details.day
            else:
                # If date_details is empty, use some default values
                print('No date details found in the email.')

            # Check if the keywords match with existing keywords :remittance
            if 'remittance' in keywords:
                driver = login_to_accountium(userAccount, userPassword)
                Remittance_Operation.getRecipient_email(userAccount)
                remittance_date = extract_month_year(body)
                print(remittance_date)
                month = remittance_date[0]
                year = remittance_date[1]
                Remittance_Operation.create_remittance(driver, month, year)
                logout(driver)
            # Check if the keywords match with existing keywords :payroll
            if 'payroll' in keywords:
                driver = login_to_accountium(userAccount, userPassword)
                employee_details = extract_employee_details(body)
                Payroll_Operation.getRecipient_email(userAccount)
                print(employee_details)
                Payroll_Operation.create_payroll(driver, year, month, day, employee_details)
                logout(driver)
            # Check if the keywords match with existing keywords :T4
            if 't4' in keywords:
                date_details, employee_details = extract_date_and_names(body)
                driver = login_to_accountium(userAccount, userPassword)
                T4_Operation.getRecipient_email(userAccount)
                print(date_details, employee_details)
                T4_Operation.create_T4(driver, date_details, employee_details)
                logout(driver)
            # Check if the keywords match with existing keywords :Invoice
            if 'invoice' in keywords:
                driver = login_to_accountium(userAccount, userPassword)
                Invoice_Operation.getRecipient_email(userAccount)
                extracted_details = extract_details_from_email(body)
                print(customer_name, items)
                Invoice_Operation.create_invoice(driver, extracted_details)
                logout(driver)
            # Check if the keywords match with existing keywords : Income statement
            if 'income ' and 'statement' in keywords:
                driver = login_to_accountium(userAccount, userPassword)
                report_range, date_ranges = extract_date_and_report_range(body)
                IncomeStatement_Operation.getRecipient_email(userAccount)
                IncomeStatement_Operation.create_incomeSatetment(driver, report_range, date_ranges)

            # Check if the keywords match with existing keywords : balance sheet
            if 'balance' and 'sheet' in keywords:
                driver = login_to_accountium(userAccount, userPassword)
                report_range, date_ranges = extract_date_and_report_range(body)
                BalanceSheet_Operation.getRecipient_email(userAccount)
                BalanceSheet_Operation.create_balanceSheet(driver, report_range, date_ranges)

            # delete the processed file
            reset_globals()
        # Mark email as read
        imap_server.store(email_id, '+FLAGS', '\\Seen')
# close connection
imap_server.close()
imap_server.logout()

