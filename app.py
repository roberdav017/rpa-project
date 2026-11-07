import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows



# Question 1:

while True:
    directory = select_file('Please select a file in the directory where ProductData.xlsx, MasterInvoice.pdf and Orders 1-4 (as Order#.xlsx) are saved.')
    directory = directory.split('/')
    del directory[-1]
    directory = '\\'.join(directory)
    directory = directory + '\\'
    directory_list = os.listdir(directory)

    if (('ProductData.xlsx' in directory_list) and 'MasterInvoice.pdf' in directory_list):
        if (((('Order1.xlsx' in directory_list) and 'Order2.xlsx' in directory_list) and 'Order3.xlsx' in directory_list) and 'Order4.xlsx' in directory_list):
            break
        else:
            show_message('Error',"Please select a file in the directory with all 4 Orders.")
    else:
        show_message('Error','Please select a file in the direcotry with ProductData.xlsx and MasterInvoice.pdf.')



# Question 2: 

descriptions = pd.read_excel(directory + 'ProductData.xlsx',sheet_name = 'Descriptions')
product_descriptions = {}

for index,row in descriptions.iterrows():
    unit_number = row['Unit number']
    description = row['Description']
    product_descriptions[unit_number] = description

prices = pd.read_excel('ProductData.xlsx', sheet_name = 'Prices')
product_prices = {}

for index,row in prices.iterrows():
    unit_number = row['Unit number']
    price_per_unit = row['Price per unit']
    product_prices[unit_number] = price_per_unit



# Question 3:

def get_order_data(invoice_num, filename, directory):
    template = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']
    
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    order_date = selected_sheet['C9'].value
    order_date = order_date.strftime('%m/%d/%Y')

    street_address = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zip_code = selected_sheet['C19'].value

    update_pdf = {}
    update_pdf['InvoiceNum'] = str(invoice_num)
    update_pdf['Date'] = order_date
    update_pdf['CustomerName'] = name
    update_pdf['CustomerCompany'] = selected_sheet['C15'].value
    update_pdf['CustomerAddress1'] = street_address
    update_pdf['CustomerAddress2'] = f'{city}, {state} {zip_code}'
    total = 0

    unit_number0 = selected_sheet['B24'].value
    if pd.isnull(unit_number0) == False:
        description0 = product_descriptions[unit_number0]
        price0 = product_prices[unit_number0]
        quantity0 = selected_sheet['C24'].value
        amount0 = quantity0 * price0
        total = total + amount0

        update_pdf['Description0'] = description0
        update_pdf['Quantity0'] = f'{quantity0}'
        update_pdf['Price0'] = f'${price0:,.2f}'
        update_pdf['Amount0'] = f'${amount0:,.2f}'

    unit_number1 = selected_sheet['B25'].value
    if pd.isnull(unit_number1) == False:
        description1 = product_descriptions[unit_number1]
        price1 = product_prices[unit_number1]
        quantity1 = selected_sheet['C25'].value
        amount1 = quantity1 * price1
        total = total + amount1

        update_pdf['Description1'] = description1
        update_pdf['Quantity1'] = f'{quantity1}'
        update_pdf['Price1'] = f'${price1:,.2f}'
        update_pdf['Amount1'] = f'${amount1:,.2f}'
    
    unit_number2 = selected_sheet['B26'].value
    if pd.isnull(unit_number2) == False:
        description2 = product_descriptions[unit_number2]
        price2 = product_prices[unit_number2]
        quantity2 = selected_sheet['C26'].value
        amount2 = quantity2 * price2
        total = total + amount2

        update_pdf['Description2'] = description2
        update_pdf['Quantity2'] = f'{quantity2}'
        update_pdf['Price2'] = f'${price2:,.2f}'
        update_pdf['Amount2'] = f'${amount2:,.2f}'

    unit_number3 = selected_sheet['B27'].value
    if pd.isnull(unit_number3) == False:
        description3 = product_descriptions[unit_number3]
        price3 = product_prices[unit_number3]
        quantity3 = selected_sheet['C27'].value
        amount3 = quantity3 * price3
        total = total + amount3

        update_pdf['Description3'] = description3
        update_pdf['Quantity3'] = f'{quantity3}'
        update_pdf['Price3'] = f'${price3:,.2f}'
        update_pdf['Amount3'] = f'${amount3:,.2f}'
    
    unit_number4 = selected_sheet['B28'].value
    if pd.isnull(unit_number4) == False:
        description4 = product_descriptions[unit_number4]
        price4 = product_prices[unit_number4]
        quantity4 = selected_sheet['C28'].value
        amount4 = quantity4 * price4
        total = total + amount4

        update_pdf['Description4'] = description4
        update_pdf['Quantity4'] = f'{quantity4}'
        update_pdf['Price4'] = f'${price4:,.2f}'
        update_pdf['Amount4'] = f'${amount4:,.2f}'

    update_pdf['Total'] = f'${total:,.2f}'
    return update_pdf, name, email, order_date




# Question 4:

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)



# Question 5:

def send_email(invoice_num, name, email, date, directory):
    invoice_num = str(invoice_num)
    name = name
    email = email
    date = date
    directory = directory

    # Enter your gmail user name and password

    gmail_user = "david.roberts1127@gmail.com"
    gmail_password_path = "\\Python3\\Class - In Class Activities\\Unit 3 - Robotic Process Automation\\Day 2 - Invoices, PDFs & Emails\\gmail_password.txt"
    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice{invoice_num}'

    # Body of the email

    body = f"""
    Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.
        
    Python Tech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    message = add_attachment(directory+f'Invoice {invoice_num}.pdf', message)
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()



# Question 6:

filenames = os.listdir(directory)

for filename in filenames:
    invoice_num = 8514
    if ('Order' in filename and 'xlsx' in filename):
        num = filename
        num = int(num.split('.')[0][5:])
        invoice_num = invoice_num + num
        
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)