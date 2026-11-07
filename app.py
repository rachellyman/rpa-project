# Insert your python code here

#import packages
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

# Q1
while True:    
    directory = single_input("Please enter the path to your folder.")
    directory = directory.replace('"','')
    if directory.endswith('\\') == True:
        break
    else:
        directory = directory +'\\'
        break
filenames = os.listdir(directory)
while True:
    if 'ProductData.xlsx' and 'MasterInvoice.pdf' and 'Order1.xlsx' and 'Order2.xlsx' and 'Order3.xlsx'and 'Order4.xlsx' in filenames:
        break
    else:
        show_message("Error", "The file you entered is not correct. Please try again.")


print(directory)


# Q2
product_descriptions = {}
product_prices = {}


ProductData = 'ProductData.xlsx'
full_path = os.path.join(directory, ProductData)


df = pd.read_excel(full_path, sheet_name='Descriptions')
for index, row in df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']


print(product_descriptions)


dfp = pd.read_excel(full_path, sheet_name = 'Prices')
for index, row in dfp.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']


print(product_prices)


# Q3
import datetime

def get_order_data(invoice_num, filename, directory):
    template = openpyxl.load_workbook(directory + filename, data_only=True)
    selected_sheet = template['OrderForm']

    order_dict = {}

    company = selected_sheet['C15'].value
    name = selected_sheet['G14'].value
    email = selected_sheet['C13'].value
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    address = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zipcode = str(selected_sheet['C19'])

    order_dict['InvoiceNum'] = int(invoice_num)
    order_dict['Date'] = date
    order_dict['CustomerCompany'] = company
    order_dict['CustomerAddress1'] = address
    order_dict['CustomerAddress2'] = f'{city}, {state} {zipcode}'
    order_dict['CustomerName'] = name

    total = 0
    row = 24
    item = 0 

    while True:
        unit_num = selected_sheet[f'B{row}'].value
        qty = selected_sheet[f'C{row}'].value
        if unit_num is None:
            break
        desc = str(product_descriptions[unit_num])
        price = product_prices[unit_num]

        amount = price * qty
        total = total + amount

        order_dict[f'Description{item}'] = desc
        order_dict[f'Price{item}'] = f'${price:,.2f}'
        order_dict[f'Quantity{item}'] = str(qty)
        order_dict[f'Amout{item}'] = f'$(amount:,.2f)'
    
        row = row + 1
        item = item + 1

    order_dict['Total'] = f'${total:,.2f}'
    return order_dict, name, email, date

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(f'{order_data, name, email, date}')


# Q4
def create_invoice(invoice_num, order_data, directory):
    create_pdf('MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)


create_invoice(8515, order_data, directory)


# Q5
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):
    gmail_user = 'r.annlyman@gmail.com'

    recipients = email


    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'


    body = f"""Dear {name},
   
    Attached is the invoice dated {date}.


    Thank you for your business.
   
    Python Tech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone (801) 555-2372
    Email: sales@pythontechco.com"""


    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)


    filename = f'Invoice {invoice_num}.pdf'
    path_to_file = os.path.join(directory, filename)


    path_to_file = directory+f'\\Invoice {invoice_num}.pdf'


    message = add_attachment(path_to_file, message)
       


    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)


# Send the email


    mail_server.send_message(message)
    mail_server.quit()


send_email(8515,'Thomas Holt','r.annlyman@gmail.com','11/13/2023', directory)


# Q6
import os
filenames = os.listdir(directory)
print(filenames)

number = 8515


for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(number, file, directory)
        create_invoice(number, order_data, directory)
        send_email(8515, name, 'r.annlyman@gmail.com', date, directory)
        number = number + 1
