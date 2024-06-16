# -*- coding: utf-8 -*-
"""
Created on Tue Dec 27 16:45:19 2022

@author: 2282174
"""
"""
Part 1:
Import necessary libraries to automate the process of downloading attachments from gmail
1. imaplib is used for accessing emails over imap protocol
2. email is used for managing email messages
3. os is used for manipulating the installed operating system as well as the PC file system.
Login to the #Username:ib96d0assignment@gmail.com #Password:Assignment2022 GMAIL ID
Make sure to keep 1 path for all operations - downloading attachments, saving merged file & sales report so that the function can run smoothly"""

import imaplib
import email
import os

# Connect to the Gmail IMAP server
imap_server = imaplib.IMAP4_SSL('imap.gmail.com', 993)

# Login to your account
imap_server.login ('ib96d0assignment@gmail.com', 'ltpvftoecqzmenhs') #we need to allow imap access on gmail and generate an app-sepecific password to use with our Python script - switch on the two factor authentication and then create the app password.
#Username:ib96d0assignment@gmail.com #Password:Assignment2022 (You can login to GMAIL - There is only one email in it with 3 attachments = txt, xlsx, and csv) #Given password is gmail generated app specific password.

# Select the INBOX folder
imap_server.select('INBOX')

# Search for all messages in GMAIL
status, data = imap_server.uid("SEARCH", None, "All")
inbox_item_list = data[0].split()

# Get the latest message
new = inbox_item_list[-1] #this will fetch the latest gmail (yesterday's GMAIL received from offices - make sure the file with attachments is the latest mail in your inbox)

result2, email_data = imap_server.uid('fetch', new, '(RFC822)') #fetch

raw_email = email_data[0][1].decode("utf-8") #decode
email_message = email.message_from_string(raw_email)

# # Iterate through all of the message's attachments
for part in email_message.walk():
    # Check if the part is an attachment
    if part.get_content_maintype() == 'multipart':
        continue
    if part.get('Content-Disposition') is None:
        continue

    # Get the file name and create a folder for it
    file_name = part.get_filename()
    if bool(file_name):
      filepath = os.path.join(file_name) #you have to select a particular path in laptop to download all three attachments from email
      if not os.path.isfile(filepath):
        fp = open(filepath, 'wb')
        fp.write(part.get_payload(decode=True))
        fp.close()

# # Close the connection to the server
imap_server.close()
imap_server.logout()

""" Part 2: to merge files from laptop path in which we saved those attachments earlier we will be using pandas
pandas is an open-source library that is made mainly for working with relational or labeled data both easily"""

import pandas as pd

# Read in the CSV file
file_csv = pd.read_csv('Office2.csv')

# Read in the XLSX file
file_xlsx = pd.read_excel('Office1.xlsx')

# Read in the TXT file
file_txt = pd.read_csv('Office4.txt', delimiter=", ", engine= "python")  # specify the delimiter as a tab to specify next value in txt file (will be used for any other file than excel and csv)

# Merge all three dataframes/datasets
file_merged = pd.concat([file_csv, file_xlsx, file_txt])

# Save the merged dataframe to a new CSV file
file_merged.to_csv('mergedfile.csv', index=False) #Naming Merged file 

""" 
Part 3: Given we have SKUs, Sales, Return, Loss, and Stock as our dataset labels, we will be analysing which SKU (out of all 3 offices SKUs) has the highest sales in which unit.
Office 1 SKUs are named as A1, A2, A3, .....................A11
Office 2 SKUs are named as B1, B2, B3, .....................B11
Office 4 SKUs are named as D1, D2, D3, .....................D11

we are analysing the sales using pie chart plotting

import matplotlib library to plot diagrams in python and 
csv to read the merged file which is in CSV format
"""
import matplotlib.pyplot as plt
import pandas as pd

# Read the CSV file into a dataframe (df)
df = pd.read_csv("mergedfile.csv")

# Extract the data and labels from the dataframe
data = df["Sales"]
labels = df["SKU"]

# Set the figure size
plt.figure(figsize=(30,39))

# Show percentage on Pie chart to analyze the sales percentage of SKUs out of total sales for all 3 offices.
plt.pie(data, autopct='%1.1f%%')

# Create the bar plot
plt.pie(data, labels=labels)

# Change the legend name (legend is the index shown on pie chart)
legend = plt.legend(title="SKU Name", fontsize=25, loc="upper left")
legend.get_title().set_fontsize(25)

# Add a title for the pie slices
plt.title("Sales")

# Create the pie chart
plt.show()

"""" 
Similarly we can plot the loss (on sold SKUs to analyze which SKU has the most chances of getting damaged)
"""

import matplotlib.pyplot as plt
import pandas as pd

# Read the CSV file into a dataframe (df)
df = pd.read_csv("mergedfile.csv")

# Extract the data and labels from the dataframe
data = df["Loss"]
labels = df["SKU"]

# Set the figure size
plt.figure(figsize=(30,39))

# Show percentage on Pie chart to analyze the Loss percentage of SKUs out of total Loss (on SKUs) for all 3 offices.
plt.pie(data, autopct='%1.1f%%')

# Create the bar plot
plt.pie(data, labels=labels)

# Change the legend name (legend is the index shown on pie chart)
legend = plt.legend(title="SKU Name", fontsize=25, loc="upper left")
legend.get_title().set_fontsize(25)

# Add a title for the pie slices
plt.title("Losses")

# Create the pie chart
plt.show()

"""analysis can be run for rest particulars - Return and Stock as well
Part 4: after plot analasis of sales and losses, we created a SalesReport MS Word Document, which will be sent to the sales manager using python
for that we will be using smtplib library to send email from gmail,
the base64 module is used to encode and decode data,
MIME (Multipurpose Internet Mail Extensions) is a standard way of describing a data type in the body of an HTTP message or email.
"""
import smtplib
import os
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Set up the SMTP server
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login('ib96d0assignment@gmail.com', 'ltpvftoecqzmenhs') #same email used earlier to download attachment from

# Create the email message
msg = MIMEMultipart()
msg['To'] = 'ib96d0assignment@gmail.com' #sending file to the same email address
msg['Subject'] = 'Summary Report' #Gmail Subject

# Add the message body
msg.attach(MIMEText('Hey, Please find the summary report in the attachment')) #Body Text

# Open the file in binary mode
with open('SalesReport.pdf', 'rb') as f: #report from computer path has been attached (make sure file exists in the path)
    # Add the file as an attachment to the email
    attachment = MIMEBase('application', 'octet-stream')
    attachment.set_payload(f.read())

# Encode the attachment in base64 format
encoders.encode_base64(attachment)

# Add header with the file name
attachment.add_header('Content-Disposition', f'attachment; filename={f.name}')
msg.attach(attachment)

# Send the email
server.sendmail('ib96d0assignment@gmail.com', 'ib96d0assignment@gmail.com', msg.as_string()) #sending email to self gmail address, however it can be sent to any gmail address. 

# Close the SMTP server
server.quit()

""" THANK YOU """ 
