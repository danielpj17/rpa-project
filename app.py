# Imports
import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows# Insert your python code here


# Question 1
while True:
    directory = (single_input("Please input the directory path where the Excel files are located:")).replace('"', '')
    directory = directory + "\\"
    filenames = os.listdir(directory)
    if "Order1.xlsx" in filenames and "Order2.xlsx" in filenames and "Order3.xlsx" in filenames and "Order4.xlsx" in filenames and "ProductData.xlsx" in filenames and "MasterInvoice.pdf" in filenames:
        break
    else:
        show_message('Error', 'One or more of the required files are missing. Please ensure that the required files are present in the specified directory')

# Question 2
product_descriptions = {}
product_prices = {}

df_desc = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
for index, row in df_desc.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df_price = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
for index, row in df_price.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

# Question 3
def get_order_data(invoice_num, filename, directory):
    order_data = {}

    df_orderform = pd.read_excel(directory + filename, sheet_name='OrderForm', dtype=str)
    df_orderform = df_orderform.rename(columns={'Unnamed: 1': 'Col1', 'Unnamed: 2': 'Col2'})
    df_orderform = df_orderform[['Col1', 'Col2']]

    number = 0
    productnum = 0
    dict = {}
    order_data = {}
    order_data1 = {}
    products = {}
    quantity = {}
    prod_desc = {}
    prod_price = {}
    amount = {}
    for t, value in df_orderform.iterrows():
        dict[number] = value['Col1']
        order_data1[value['Col1']] = value['Col2']
        number = number + 1
        if number > 22:
            products['Product' + str(productnum)] = value['Col1']
            quantity['Product' + str(productnum)] = value['Col2']
            productnum = productnum + 1

    for i, product in products.items():
        prod_desc[i] = product_descriptions[product]
        prod_price[i] = product_prices[product]
        amount[i] = float(quantity[i]) * float(prod_price[i])

    date = order_data1['Order Date'].split('-')
    year = date[0]
    month = date[1]
    day = date[2][:2]
    date = month + '/' + day + '/' + year

    city = order_data1['City']
    state = order_data1['State']
    zip_code = order_data1['Zipcode']
    address2 = city + ', ' + state + ' ' + zip_code
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerAddress1'] = order_data1['Address'].strip()
    order_data['CustomerName'] = order_data1['Name']
    order_data['CustomerCompany'] = order_data1['Company Name']
    order_data['CustomerAddress2'] = address2
    order_data['CustomerEmail'] = order_data1['Email']
    name = order_data['CustomerName']
    email = order_data['CustomerEmail']
    date = order_data['Date']

    total = 0
    for y, test in enumerate(quantity):
        order_data['Price' + str(y)] = f'${prod_price[test]:.2f}'
        order_data['Quantity' + str(y)] = quantity[test]
        order_data['Amount' + str(y)] = f'${amount[test]:.2f}'
        # order_data['Product' + str(y)] = products[test]
        order_data['Description' + str(y)] = prod_desc[test]
        total = total + amount[test]
    order_data['Total'] = f'${total:.2f}'

    
    return order_data, name, email, date


order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(order_data)
print(name)
print(email)
print(date)

# Question 4
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f"Invoice {invoice_num}.pdf", order_data)

create_invoice(8515, order_data, directory)

# Question 5

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):

    gmail_user = "danieljohnson6@gmail.com"

    # List of recipients

    recipients = [email]

    # Email subect

    subject = "Your order from PythonTech Co. - Invoice " + str(invoice_num)

    # Body of the email

    body = f"""Dear {name},

    Attached is your invoice date {date}.

    Thank you for your business.

    PythonTech Co.
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

    path_to_file1 = f'{directory}Invoice {invoice_num}.pdf'

    message = add_attachment(path_to_file1, message)
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


invoice_num = 8515
name = order_data['CustomerName']
email = order_data['CustomerEmail']
date = order_data['Date']
directory = directory

send_email(invoice_num, name, email, date, directory)

# Question 6
filenames = os.listdir(directory)
invoice_num = 8515
for filename in filenames:
    if "Order" in filename and "xlsx" in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1
