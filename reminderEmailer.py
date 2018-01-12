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
wb = openpyxl.load_workbook('Registration Manipulated (For E-Mails).xlsx')
# Get all sheet names
wb.get_sheet_names()

#gets the specific sheet
response = wb.get_sheet_by_name('responses') # Note here name is case sensitive
#response = wb.get_sheet_by_name('Patrick')
print(response)

#used to save the corresponding names and emails
names = []
emails = []
print ('Reading rows...')

#reads in all the rows of names and emails and then adds them to their respective lists
for row in range(2, 68): # for loop is open in the end
    names.append(str(response['C'+str(row)].value).strip()) # Note here how to manipulate strings in Python
    emails.append(str(response['D'+str(row)].value).strip())

names.append("Jack Quimby")
emails.append("Quimbyj@gonzaga.org")
names.append("Patrick Riley")
emails.append("rileyp@bishopireton.org")
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
	 <title>HackBI Email Flyer</title>
         </head>
	 <!-- Actual Body tag stuff: style='background: linear-gradient(rgb(252, 77, 69),rgb(252, 77, 69),rgb(255, 179, 71), rgb(255, 179, 71));'-->


         <body style='background: linear-gradient(rgb(252, 77, 69),rgb(252, 77, 69),rgb(255, 179, 71), rgb(255, 179, 71));'>
	 <div id="wrapper" style="text-align:center;position:relative;z-index: 1;margin-bottom:-15%">
		 <img src="http://www.hackbi.org/img/banner.jpg" alt="HackBI Logo" width="80%vw" style='text-align:center;font-family:"Arial Black",sans-serif; font-size: 24px;z-index: 1;' />
	 </div>

	 <div id="Reminders" style='text-align:center;line-height:150%;font-family:"Arial Black",sans-serif;position:relative;z-index: 1;margin-top:6%'>
		 <h1 align=center style='line-height:150%;font-size:40.0pt;color:white;-webkit-text-stroke: 2px rgb(232, 235, 239);'><strong>HACKBI IS 13 DAYS AWAY!</strong></h1>
		 <h1 align=center style='line-height:150%;font-size:18.0pt;color:rgb(232, 235, 239);'><u>Remember to get those permission forms in! You cannot compete without them!!!</u></h1>
	 </div>



	 <table style='margin: 0 auto 0 auto;'>
		 <tr style='text-align:left;font-family:"Arial Black",sans-serif; font-size: 30px; color:red;-webkit-text-stroke: 1px rgb(232, 235, 239);'>
			 <!--<th style='text-align:center;'>Prizes!!!</th>
			 <th>Important Links</th>-->
		 </tr>
		 <tr>
			 <td rowspan='2' style='border:5px solid rgb(232, 235, 239); padding:15px;'>
				 <h2 style='margin-top: -17%;font-family:"Arial Black",sans-serif;font-size: 30pt;text-align:center;color:red;-webkit-text-stroke: 1px rgb(232, 235, 239);'>PRIZES!!</h2>
				 <ul style='font-family:"Arial Black",sans-serif; color:red;font-size: 24px;'>
					 <li style='font-size: 32px;'><u>MYSTERY GRAND PRIZE</u></li>
					 <li>Amazon Echo Dots</li>
					 <li>Elegoo Smart Robot Car Kit</li>
					 <li>Bluetooth Speakers</li>
					 <li>Mechanical Keyboards</li>
					 <li>//CODE Robot Repair</li>
				 </ul>
			 </td>

			 <td style='text-align:left;font-family:"Arial Black",sans-serif; font-size: 24px;border-top:5px solid rgb(232, 235, 239);border-right:5px solid rgb(232, 235, 239);padding:20px;'>
				 <br><a href="http://www.hackbi.org" target="_blank">HackBI Website</a> (Invite your friends!)<br>
				 <br>Check out this <a href="https://goo.gl/forms/ICVhbjq0tppoZ5kj2" target="_blank">Worshop Interest Form!</a>
			 </td>
		 </tr>

		 <tr>
			 <td style='text-align:top;font-family:"Arial Black",sans-serif; font-size: 24px;border-right:5px solid rgb(232, 235, 239);border-bottom:5px solid rgb(232, 235, 239);padding:10px;'>
			 <h3 style='text-align:left;font-size: 32px;color:red;-webkit-text-stroke: 1px rgb(232, 235, 239);margin-bottom:-.5%'>Important Forms</h3>
			 <a href="http://hackbi.org/files/parentPermission.pdf">Parent Permission Form</a><br>
			 <a href="http://hackbi.org/files/emergencyCarePermission.doc">Emergency Care Form</a><br>
			 <a href="http://hackbi.org/files/disciplinary.docx">Disciplinary Information</a>
			 <br><br>
			 </td>
		 </tr>

	 </table>
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

