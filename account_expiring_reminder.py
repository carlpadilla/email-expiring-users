#! python3
# account_expiring_reminder.py - Sends emails based on expiration date in spreadsheet.
import openpyxl
import smtplib
import sys
from datetime import timedelta, datetime

# 15 days until password expires
endDate = datetime.today() + timedelta(days=15)


# Open the spreadsheet and get the latest dues status.
# spreadsheet should have three columns name, email, experation date
wb = openpyxl.load_workbook('FILE_PATH_Here.xlsx')

sheet = wb.get_sheet_by_name('Sheet1')

lastCol = sheet.max_column

expiryMonth = sheet.cell(row=3, column=lastCol).value

# Check each member's expiry status.
expiringUsers = {}
for r in range(3, sheet.max_row + 1):
    expiry = sheet.cell(row=r, column=lastCol).value
    # print(expiry)

    if expiry < endDate:
        name = sheet.cell(row=r, column=1).value
        email = sheet.cell(row=r, column=2).value
        expiringUsers[name] = email
# uncomment to view users in expiringUsers Obj
# print(expiringUsers)


#  Log in to email account.
# Enter smtp address and port number
smtpObj = smtplib.SMTP('smtp.office365.com', 587)

# establishing a connection to the server
smtpObj.ehlo()

# enables TLS encryption for your connection on port 587
smtpObj.starttls()


#
smtpObj.login('example@domain.com', input('Enter email password: '))

''' sendmail() method requires three arguments. from address, recipient's, and email body.
The start of the email body string must begin with 'Subject: \n' for the
subject line of the email. The '\n' newline character separates the subject
line from the main body of the email.
'''


# TODO: Send out reminder emails.
for name, email in expiringUsers.items():
    message = f"""\
Subject: {name} Domain Account Expiring soon.

Dear {name},

message here
    """
    print(f'Sending email to {email}...')
    sendmailStatus = smtpObj.sendmail(
        'example@domain.com', email, message)
    if sendmailStatus != {}:
        print(f'There was a problem sending email to {name}: {email}')
        smtpObj.quit()
