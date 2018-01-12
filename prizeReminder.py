#mass emailer for the hackathon
#Patrick Riley, inspired by http://yuqli.com/?p=653
#12/1/17
#Python 3.6.1 (v3.6.1:69c0db5, Mar 21 2017, 18:41:36)


import openpyxl, pprint, smtplib, base64
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.message import EmailMessage
from email.utils import make_msgid


# Read tables into a Python object

print ('Opening workbook...')
wb = openpyxl.load_workbook('HackBI Registration.xlsx')
# Get all sheet names
wb.get_sheet_names()

#gets the specific sheet
response = wb.get_sheet_by_name('responses') # Note here name is case sensitive
print(response)

#used to save the corresponding names and emails
names = []
emails = []
names.append("Patrick Riley")
emails.append("rileyp@bishopireton.org")
print ('Reading rows...')
#reads in all the rows of names and emails and then adds them to their respective lists
for row in range(2, 97): # for loop is open in the end
    grade = str(response['E'+str(row)].value).strip()
    if  grade == "9.0" or grade == "10.0" or grade == "11.0" or grade == "12.0" or grade == "None":
        names.append(str(response['B'+str(row)].value).strip() + " " + str(response['C'+str(row)].value).strip())
        emails.append(str(response['D'+str(row)].value).strip())

#names.append("Jack Quimby")
#emails.append("Quimbyj@gonzaga.org")
"""
names.append("Peter Murphy")
emails.append("murphyp1@bishopireton.org")
names.append("Emily Roddy")
emails.append("roddye@bishopireton.org")
names.append("Terri Kelly ")
emails.append("kellyt@bishopireton.org")
names.append("Alexander Winstanley")
emails.append("winstanleya@bishopireton.org")
names.append("Sean Gibbons")
emails.append("gibbonss@bishopireton.org")
names.append("Braden Hoagland")
emails.append("hoaglandb1@bishopireton.org")
names.append("Jared")
emails.append("ponmakhaj@bishopireton.org")
"""

#email server stuff, login and connect
smtpObj = smtplib.SMTP('smtp.gmail.com:587')
smtpObj.starttls() # Upgrade the connection to a secure one using TLS (587)
# If using SSL encryption (465), can skip this step
smtpObj.ehlo() # To start the connection
smtpObj.login('hackbi@bishopireton.org', 'Windows12')


#this is where the email actually is created and sent
#creates an email and adds some basic attributes
for i in range(0, len(emails)):
    msg = EmailMessage()
    msg['From'] = 'HackBI <hackbi@bishopireton.org>' # Note the format
    msg['Subject'] = 'Bishop Ireton High School Hackathon!'
    flierCid = make_msgid()
    msg.set_content("this is a test")
    msg.add_alternative("""\
    <html>
    <head>
      <title>HackBI PRIZES!!!</title>
    </head>

    <body style='background: linear-gradient(rgb(252, 77, 69),rgb(252, 77, 69),rgb(255, 179, 71), rgb(255, 179, 71));'>
      <div id="wrapper" style="text-align:center;">
        <img src="http://www.hackbi.org/img/banner.jpg" alt="HackBI Logo" width="80%vw" style='text-align:center;font-family:"Arial Black",sans-serif; font-size: 24px;z-index: 1;' />
      </div>

      <p style='font-family:"Arial Black",sans-serif;font-size: 18pt;text-align:left;color:rgb(232, 235, 239);-webkit-text-stroke: 1px red;'>Hey there Coder!<br><br>
        You are receiving this message because you are registered for Hack BI.<br><br>
        There are still a few slots open, but registration closes on 1/13.  In case you have some friends who are on the fence, we thought we would let you know about some of the prizes - to encourage them to join you!<br><br><br>
      </p>
    
      <dl style='font-family:"Arial Black",sans-serif;font-size: 18pt;text-align:left;color:rgb(232, 235, 239);-webkit-text-stroke: 1px red;'>
        <dt>These prizes apply ONLY to students in high school:<br><br>
          <dd><u>Grand Prize</u> - <span style='font-size:22pt;color:red;-webkit-text-stroke: 1px rgb(232, 235, 239) '><b>Paid summer internship</b> with a Software Development Company</span> in Alexandria, VA (one for each member of the team)<br><br>
            <dd><u>Best Game</u>- <span style='font-size:22pt;color:red;-webkit-text-stroke: 1px rgb(232, 235, 239) '><b>Oculus Rifts</b></span> (all members of the team get one)<br><br>
            </dl>

            <p style='font-family:"Arial Black",sans-serif;font-size: 18pt;text-align:left;color:rgb(232, 235, 239);-webkit-text-stroke: 1px red;'>
              and there are 4 other categories!!<br><br>
              One in four students will walk away with a prize. Granted, not all prizes are at the category of a paid internship or an Oculus, but everyone who attends has a really great chance of winning one!<br><br>
              One of the categories is exclusively for girls' teams - so, gals - get those friends who are just not sure to commit and join you!! <br><br>
              Please share the information with your teachers and friends and make sure they register by January 13th! <br><br>
              Can't wait to see you at Bishop Ireton!<br>
              <blockquote style='font-family:"Arial Black",sans-serif;font-size: 18pt;text-align:left;color:rgb(232, 235, 239);-webkit-text-stroke: 1px red;'>-Hack BI Team</blockquote>
            </p>

      </body>
      </html>
      """.format(flierCid=flierCid[1:-1]), subtype='html')

        #print("sending email(s)")
        #sends the email to that email being iterated through

    msg['To'] = '%s <' % names[i] + emails[i] + '>'
    smtpObj.sendmail('hackbi@bishopireton.org',emails[i],msg.as_string())
    print("Email sent to %s " % str(emails[i]))

#closes everything up

smtpObj.quit()

