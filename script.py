"""
Script that reads contact list and message template
and send the mail to all recipients
"""

import sys
import os
# import the smtplib module. It should be included in Python by default
import smtplib

# import Template to replace strings in text body
from string import Template

# import necessary packages
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

MY_ADDRESS = "myemailaddress"
PASSWORD = "mypassword"

filename = "filename"
filepath = "filepath"


# Get contacts file name and message template file name
def get_filenames():
    contactfile = input("Contact file name: ")
    messagefile = input("Message template file name: ")
    return contactfile, messagefile


# Function to read the contacts from a given contact file
# @param filename - file that contains contact information
# @return - 2 lists containing names & email addresses
def get_contacts(filename):
    names = []
    emails = []
    try:
        # Open file in Read-only
        with open(filename, mode='r', encoding='utf-8') as contacts_file:
            for a_contact in contacts_file:
                names.append(a_contact.split()[0])
                emails.append(a_contact.split()[1])
    except IOError:
        print("An error occurred while trying to read the file.")

    return names, emails


# Function to read the message template file
# @param filename - file that contains message template
# @return - Template object
def read_template(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as template_file:
            template_file_content = template_file.read()
    except IOError:
        sys.exit("An error occurred while trying to read the file.")

    return Template(template_file_content)


try:
    # set up the SMTP server
    # s = smtplib.SMTP(host='smtp.gmail.com', port=587)
    s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
    s.starttls()    # Encrypt SMTP commnads that follow
except:
    sys.exit("An error occurred while trying to establish connection with server.")
try:
    s.login(MY_ADDRESS, PASSWORD)
except:
    sys.exit("An error occurred while attempting login. Please verify email address and password.")


def main():
    contactfile, messagefile = get_filenames()
    names, emails = get_contacts(contactfile)        # Read contacts from file
    message_template = read_template(messagefile)

    # For each contact, send the email:
    for name, email in zip(names, emails):
        msg = MIMEMultipart()       # create a message

        # add in the actual person name to the message template
        message = message_template.substitute(PERSON_NAME=name.title())

        # setup the parameters of the message
        msg['From']=MY_ADDRESS
        msg['To']=email
        # msg['Date']=formatdate(localtime=True)
        msg['Subject']="This is TEST"

        # add in the message body
        msg.attach(MIMEText(message, 'plain'))

        attachment = open((filepath + filename), "rb")

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-disposition', "attachment; filename= %s" % os.path.basename(filename))

        msg.attach(part)

        text = msg.as_string()

        # send the message via the server set up earlier.
        # s.send_message(msg)
        s.sendmail(MY_ADDRESS, email, text)
        del msg

    s.quit()


if __name__ == '__main__':
    main()
