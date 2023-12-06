import pandas as pd
import smtplib

df = pd.read_excel('Trail.xlsx')

for index, row in df.iterrows():
    email = row['Email']
    gender = row['P']
    subject = 'Request for Internship Opportunity'
    body1 = '''
Dear Sir,

I hope this email reaches you well. I am sending this email to ask about an internship opportunity for this summer. As a student of IIT KHARAGPUR, I would like to gain practical experience in the Data Analytics field. I think this is the perfect platform to learn and grow. I am particularly interested in Data Analytics and am impressed by the company's commitment to the usage of Analytics and Tech Consolidation.

I am excited about the possibility of contributing to your team while also developing my skills and knowledge. 

Here is the link for my CV

Thank you for your time and consideration. I look forward to hearing from you.
Sincerely, 

Neil Songa
'''
    body2 = '''
Dear Madam,

I hope this email reaches you well. I am sending this email to ask about an internship opportunity for this summer. As a student of IIT KHARAGPUR, I would like to gain practical experience in the Data Analytics field. I think this is the perfect platform to learn and grow. I am particularly interested in Data Analytics and am impressed by the company's commitment to the usage of Analytics and Tech Consolidation.

I am excited about the possibility of contributing to your team while also developing my skills and knowledge. 

Here is the link for my CV

Thank you for your time and consideration. I look forward to hearing from you.
Sincerely, 

Neil Songa
'''
    if gender == 'Mr.':
      body=body1
    else:
      body=body2
    # Enter your email and password below
    # You will get the sender password in your account security>>app passwords
    #if you are unable to find app passwords the complete the 2 step verification in security
    sender_email = 'neilpratap2016@gmail.com'
    sender_password = 'cipl isxr yrug pxvk'
    
    smtp_server = 'smtp.gmail.com'
    port = 587
    server = smtplib.SMTP(smtp_server, port)
    server.starttls()
    server.login(sender_email, sender_password)
    
    message = f'Subject: {subject}\n\n{body}'
    server.sendmail(sender_email, email, message)
    
    server.quit()
