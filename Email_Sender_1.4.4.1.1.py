#! python3

# This program sends out personalized, plain text emails from a csv file. 
# The csv has to have all fields filled out.
# The list has to have a header row with the following: First Name, Full Name, Username, Domain.
# Have to have python 3.7 and pandas installed.
# Make sure the current directory is correct.

# Wish List of features to add in the future
# Html email support
# Ease of use like pick and choose your settings and it'll automatically set things up

# Import everything we need

import os, pandas, smtplib, time

from email.message import EmailMessage
from email.headerregistry import Address
from random import randint

# Input relevant information

fromName = '' # Your full name 'First Last'
fromUsername = '' # The bit before the @ in your email address
fromDomain = '' # The bit after the @ in your email address
fromEmailAddress = '' # Your full email address
fromEmailPassword = '' # Your email password
listName = '' # The name of the csv file with .csv at the end

# Change to the right directory

os.chdir('') # Path to where the csv file is

# Open the CSV file

df = pandas.read_csv(listName) # Have to have the following headers: First Name, Full Name, Username, Domain. If there aren't headers, this won't work. 

# Create the loop

for index, row in df.iterrows():

	# Create the email
	toFirstName = row['First Name']
	toName = row['Full Name']
	toUsername = row['Username']
	toDomain = row['Domain']
	msg = EmailMessage()
	msg['Subject'] = '' # What do you want the email subject to say?
	msg['From'] = Address(fromName, fromUsername, fromDomain)
	#msg['Bcc'] = Address(fromName, fromUsername, fromDomain) # Some email providers don't include emails sent out through this program in the sent folder. Uncomment this if you want to make sure you recieve a copy of each email sent. 
	msg['To'] = Address(toName, toUsername, toDomain)
	msg.set_content("""\
Hi {0},

Here is my message.

Thanks,


Kyle
""".format(toFirstName))
	#print(msg) # Not really necessary unless you want to run some tests. Can be useful if you want to make sure the first name variable matches the email information. 

	# Send the email
	with smtplib.SMTP('smtp.office365.com', 587) as s: # Have to change this depending on email provider. https://automatetheboringstuff.com/chapter16/ for smtp info
		s.ehlo()
		s.starttls()
		s.login(fromEmailAddress, fromEmailPassword)
		print('Sending email to ' + (msg['To']) + '...')
		s.send_message(msg)
		print('Success! You sent an email to ' + (msg['To']) + '.')
		s.quit()
		time.sleep(randint(37,111))


