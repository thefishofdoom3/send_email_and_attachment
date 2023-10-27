import g_drive
from datetime import datetime, timedelta
import datetime
import imaplib
import logging
import os
import smtplib
import ssl
import sys
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template

import gspread
from google.oauth2.service_account import Credentials
from tabulate import tabulate

datestring = datetime.date.today()

cwd = os.getcwd()
# ----------------------- LOG ----------------------------
log_Format = logging.Formatter("%(levelname)s %(asctime)s - %(message)s")
sh = logging.StreamHandler(sys.stdout)
sh.setFormatter(log_Format)
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(sh)

# ----------------------- CLASS ----------------------------


class CompanyOrder:
    def __init__(self, invoiceNumber, invoiceDate, dueDate, payDate, overdueDays, totalAmount, penalty, paymentRequest, bccEmail, fileIDD):
        self.invoiceNumber = invoiceNumber
        self.invoiceDate = invoiceDate
        self.dueDate = dueDate
        self.payDate = payDate
        self.overdueDays = overdueDays
        self.totalAmount = totalAmount
        self.penalty = penalty
        self.paymentRequest = paymentRequest
        self.bccEmail = bccEmail
        self.fileIDD = fileIDD


class CompanyObject:
    def __init__(self, id):
        self.id = id
        self.brandEmails = []
        self.orders = []

    def addBrandEmail(self, brandEmail):
        if brandEmail == "":
            return

        splits = brandEmail.split(",")
        for split in splits:
            if split not in self.brandEmails:
                self.brandEmails.append(split.strip())

    def addOrder(self, invoiceNumber, invoiceDate, dueDate, payDate, overdueDays, totalAmount, penalty, paymentRequest, bccEmail, fileIDD):
        self.orders.append(CompanyOrder(invoiceNumber, invoiceDate, dueDate, payDate, overdueDays, totalAmount, penalty, paymentRequest, bccEmail, fileIDD))


class CompanyManager:
    def __init__(self):
        self.companyList = []

    def companyExists(self, id):
        exists = False
        for obj in self.companyList:
            if (obj.id == id):
                exists = True
        return exists

    def getCompanyFromList(self, id):
        for obj in self.companyList:
            if (obj.id == id):
                return obj


# ----------------------- UTILS ----------------------------
def stringToNumber(str):
    splits = str.split(",")
    floatRepresentation = "".join(splits)
    fsplits = floatRepresentation.split('.')
    return int(fsplits[0])


def create_and_send_email(company, mode):
    # Create table based on orders
    table = []
    paymentRequest = ''
    sum_totalAmount = 0
    sum_penalty = 0
    table.append(['Số hoá đơn', 'Ngày hoá đơn', 'Thời hạn thanh toán',
                 'Ngày thanh toán', 'Số ngày quá hạn', 'Lãi phạt/ ngày','Tổng tiền phí dịch vụ (đã bao gồm thuế GTGT)','Số tiền phạt thực tế','Đề nghị thanh toán'])
    for order in company.orders:
        sum_totalAmount += stringToNumber(order.totalAmount)
        sum_penalty += stringToNumber(order.penalty)
        table.append([order.invoiceNumber, order.invoiceDate, order.dueDate, order.payDate, order.overdueDays, '0.05%', order.totalAmount, order.penalty, order.paymentRequest])
        paymentRequest += order.paymentRequest+', '
    table.append(['','','', '','', 'Tổng cộng', "{:,.0f}".format(sum_totalAmount), "{:,.0f}".format(sum_penalty),''])
    paymentRequestFinal = paymentRequest[:len(paymentRequest)-2]
    sum_penalty1 = "{:,.0f}".format(sum_penalty)

    seven_days_later = datestring + timedelta(days=7)
    seven_days_later1 = seven_days_later.strftime('%d/%m/%Y')

    
    #define params to use in html below
    params = {
        'companyName':  company.id,
        'table':  tabulate(table, headers="firstrow", tablefmt='html'),
        'sum_penalty': sum_penalty1,
        'paymentRequestFinal': paymentRequestFinal,
        'thedate': seven_days_later1
    }

    # Format table with inline style
    to_replace = {
        ' style="text-align: right;"': '',
        '<table>': '<table style="border: 1px solid black; border-collapse: collapse; padding: 8px">',
        '<th>': '<th style="border: 1px solid black; border-collapse: collapse; font-weight: bold; padding: 8px">',
        '<td>': '<td style="border: 1px solid black; border-collapse: collapse; padding: 8px; text-align: right">',
    }

    for k, v in to_replace.items():
        params['table'] = params['table'].replace(k, v)

    index = params['table'].rfind('<tr>')

    params['table'] = params['table'][: index] + \
        '<tr style="font-weight: bold">' + \
        str(params['table'][index+len('<tr>'):])

    # Create html and attach table
    html = """\
        <html>
        <head>
            <style type="text/css">
            table {
                text-align: center;
                vertical-align: bottom;
            }

            table,
            th,
            td {
                border: 1px solid black;
                border-collapse: collapse;
            }

            th,
            td {
                padding: 8px;
            }
            
            th {
                font-weight: bold;
            }
            </style>
        </head>
        <body>
            <span>
            content
            </span>
        </body>
        </html>
    """

    t = Template(html).safe_substitute(params)

 
    # Set up the SMTP/IMAP server
    username = 'nguyenleminhkhoa2@gmail.com'
    password = 'app_password_here'

    
    if mode == "send":
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
    else:
        server = imaplib.IMAP4_SSL('imap.gmail.com')

    server.login(username, password)

    # Create email
    msg = MIMEMultipart()
    msg['From'] = username
    
    msg['To'] = ', '.join(company.brandEmails)
    msg['CC'] = order.bccEmail
    current_date = datetime.date.today()
    formatted_date = current_date.strftime('%d/%m/%Y')
    subject = "[SHOPEE] THÔNG BÁO PHẠT CHẬM THANH TOÁN NGÀY $date - $company"
    p = {
        'date': formatted_date,
        'company': company.id
    }
    msg['Subject'] = Template(subject).safe_substitute(p)
    msg.attach(MIMEText(t, 'html'))

        

# ----------------------- ATTACH PDF FILE ----------------------------
    # Search and get file_id
    finalFileName = order.fileIDD
    
    file_id = g_drive.get_file_id(finalFileName)

    # Download file to local
    filename = g_drive.download_file(file_id)

    # Attach file to email
    with open(f"{cwd}/output/{datestring}/{filename}", "rb") as attachment:
        # Add the attachment to the message
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    msg.attach(part)

    if mode == "send":
        server.send_message(msg)
    else:
        msg = str(msg).encode("utf-8")
        server.append('[Gmail]/Drafts', '', imaplib.Time2Internaldate(
            time.time()), msg)
      


# ----------------------- MAIN ----------------------------
def main():
    companyManager = CompanyManager()
# Set up Google Sheets API credentials
    credentials = Credentials.from_service_account_file('ps.json')
    scoped_credentials = credentials.with_scopes(
        ['https://www.googleapis.com/auth/spreadsheets'])

# Authorize and create the client
    client = gspread.authorize(scoped_credentials)

# Open the Google Sheet using the URL
    sheet = client.open_by_url(
        'sheet_url').worksheet('tab_name')

# Get the parameter values from specific cells
    data = sheet.get_all_values()

# Group orders and unique brand emails(recepients) based on unique company's name
    for i in range(len(data)):
        if i > 0:
            id = data[i][43]
            exists = companyManager.companyExists(id)
            if exists == False:
                company = CompanyObject(id)
                company.addBrandEmail(data[i][44])
                company.addOrder(data[i][47], data[i][48], data[i]
                                 [49], data[i][50], data[i][51], data[i][53], data[i][54], data[i][55], data[i][45], data[i][55])
                companyManager.companyList.append(company)
            else:
                company = companyManager.getCompanyFromList(id)
                company.addBrandEmail(data[i][44]) 
                company.addOrder(data[i][47], data[i][48], data[i]
                                 [49], data[i][50], data[i][51], data[i][53], data[i][54], data[i][55], data[i][45], data[i][55])
# For each company, perform create and send email to all brand emails belong to that company
    for company in companyManager.companyList:
        print(company.id)
        create_and_send_email(company, 'send')
     


main()
