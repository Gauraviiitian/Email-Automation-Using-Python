#!/usr/bin/env python
# coding: utf-8

# In[1]:


import smtplib
import email
import pandas as pd
import numpy as np
import getpass


# In[2]:


# Function to read the contacts from a given contact file and return a
# list of names and email addresses
def get_contacts(filename):
    contacts = pd.read_excel(filename)
    names = contacts['Name']
    emails = contacts['Email ID']
    return names, emails


# In[3]:


from string import Template

# read the message template which is stored in the supplied filename
def read_template(filename):
    with open(filename, 'r') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


# In[4]:


# set up the SMTP server
s = smtplib.SMTP(host='smtp.gmail.com', port=587)
s.starttls()

# input username
UserName = input("Enter Username: ")
# input password
Password = getpass.getpass("Enter Password: ")

s.login(UserName, Password)


# In[5]:


# get contact names and email ids
names, emails = get_contacts('contacts.xlsx') 
# message templater̥
message_template = read_template('template.txt')
# print(names[0], emails[0])


# In[6]:


# import necessary packages
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

subject = input("Enter the Subject: ")
# For each contact, send the email:
for name, email in zip(names, emails):
    msg = MIMEMultipart()       # create a message

    # add in the actual person name to the message template
    message = message_template.substitute(PERSON_NAME=name.title())

    # setup the parameters of the message
    msg['From']=UserName
    msg['To']=email
    msg['Subject']=subject

    # add in the message body
    msg.attach(MIMEText(message, 'plain'))

    # send the message via the server set up earlier.
    s.send_message(msg)
    
    del msg


# In[ ]:




