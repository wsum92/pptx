# -*- coding: utf-8 -*-
"""
Created on Sat May  4 15:12:02 2019

@author: willi
"""

def send_email(week_num, report_name):
    import email
    import email.mime.application
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.image import MIMEImage
    import smtplib
    import csv
    
    # Create a text/plain message
    group = []
    msg = email.mime.multipart.MIMEMultipart()
    msg['Subject'] = 'Weekly Report'
    msg['From'] = 'trendspptx@gmail.com'
    
    answer = input('Use emails from file? (y/n): ')
    if answer.lower() == 'y':
        answer = input('Use default list? (y/n): ')
        if answer.lower() == 'y':
            with open('email_list.csv', 'r') as f:
                    for line in csv.reader(f):
                        group.append(line[0])
        else:
            answer = input("Enter file name '.csv',\ne.g. email_list.csv: ")
            with open(answer, 'r') as f:
                    for line in csv.reader(f):
                        group.append(line[0])            
    else:
        to = input('Enter email to send to,\ne.g. trendspptx@gmail.com: ')
        print("When finished type 'done': ")
        
        while to.lower() != 'done':
            group.append(to)
            to = input('Please enter next email: ')

    print('\nSending emails to:\n', group)
    
    msg['To'] = ','.join(group)
    
    # The main body is just another attachment
    body = email.mime.text.MIMEText("Report for Week " + str(week_num))
    msg.attach(body)
    
    # Input the file location, including the file name and type
    fp = open(report_name,'rb')
    
    # edit subtype to replicate the document type
    att = email.mime.application.MIMEApplication(fp.read(),_subtype="pptx")
    fp.close()
    att.add_header('Content-Disposition','attachment',filename=report_name)
    msg.attach(att)
    
    s = smtplib.SMTP("smtp.gmail.com", 587, timeout=120)
    s.starttls()
    
    # Your login information
    s.login('trendspptx','M******5')
    
    # Email information (from address, [recipient email addresses])
    s.sendmail('trendspptx@gmail.com',group, msg.as_string())
    s.quit()