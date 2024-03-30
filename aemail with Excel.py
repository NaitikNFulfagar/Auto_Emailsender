import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

email_list = pd.read_excel('C:\\Users\\email auto\\email.xlsx')
names = email_list['Name']
emails = email_list['E-mail ID']
mobno = email_list['Contact no. (WhatsApp No.)']




for i in range(len(names)):
  mail=smtplib.SMTP('smtp.gmail.com', 587)
  mail.ehlo()
  mail.starttls()
  sender='Abc Company'
  mail.login('abc@gmail.com','****password****')

  message = MIMEMultipart("alternative")
  message["Subject"] = "Congratulations!!🥁🎊"
  message["From"] = sender
  message["To"] = emails[i]

  html = """\
  <html>
    <body>
  <p><strong>E-VERSE 23 Registrations Successful&nbsp;</strong></p> 
  <p>&nbsp;</p> 
  <p>Dear <strong>"""+names[i]+"""</strong>,</p> 
  <p>&nbsp;</p> 
  <p>Congratulations!!🥁🎊 You have successfully registered for E-VERSE23💫!!! Organized by E-CELL INOVCREATA GESCOE. We are overjoyed to have you participate in this thrilling event.🤩</p> 
  <p>&nbsp;</p> 
  <p><strong>🪄Event Details🪄</strong>:</p> 
  <p><strong>🗓Dates &nbsp;:&nbsp;</strong>26th, 27th &amp; 28th October 2023</p> 
  <p><strong>⏰️Time &nbsp; :</strong> 10:00 a.m. to 5:00 p.m.</p> 
  <p><strong>🎯Venue :</strong> M.S. Gosavi Hall, Gokhale Education Society&rsquo;s, R.H.Sapat College of Engineering Management Studies of research NASHIK-422005.</p> 
  <p>&nbsp;</p> 
  <p>👀➡️We look forward to providing you with a dynamic platform where you can explore the latest trends in innovation, entrepreneurship, and technology. Get ready to engage with industry experts, participate in insightful workshops, and network with like-minded individuals. Additionally, we have some interesting plans for you all✨️.</p> 
  <p>&nbsp;</p> 
  <p><strong>🎓Your Unique QR Code</strong>:</p> 
  <p>Your unique QR code will serve as your pass to enter the event hall. Please ensure to have the QR code accessible on your device or as a printed copy for smooth entry.</p> 
  <p>&nbsp;</p> 
  <img src="https://api.qrserver.com/v1/create-qr-code/?size=250x250&data="""+str(mobno[i])+""" \" />
  <p>❓️If you have any questions or require any additional information, please don&apos;t hesitate to reach out to us.</p> 
  <p>Contact here :</p> 
  <p>Aayush Kamble&nbsp;&nbsp; : 9175123913</p>
  <p>Prathamesh Patil : 8600406179</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p> 
  <p>&nbsp;</p> 
  <p><strong>🌐Follow our social media handles to not miss any updates🌐 :</strong></p> 
  <p>&nbsp;</p> 
  <p><strong>WhatsApp group link</strong> : https://chat.whatsapp.com/E8ANaS5kJFhEOcx228Whla</p> 
  <p>&nbsp;</p> 
  <p><strong>Instagram :&nbsp;</strong>https://instagram.com/ecell.gescoe?igshid=MzRlODBiNWFlZA==</p> 
  <p><strong>Linkedin :</strong></p> 
  <p>https://www.linkedin.com/company/e-cell-inovcreata/</p> 
  <p>&nbsp;</p> 
  <p><strong>Facebook :</strong></p> 
  <p>https://www.facebook.com/ecell.gescoe/</p>
  <p><strong>&nbsp;</strong></p> 
  <p>&nbsp;</p> 
  <p>&nbsp;</p> 
  <p><em>Get ready for an unforgettable experience at E-VERSE23. We can&apos;t wait to see you there!🌟🔥🚀</em></p> 
  <p>&nbsp;</p> 
  <strong><p>Best regards,</p> 
  <p>Ms.Sakshi Sangale</p> 
  <p>CA E-CELL INOVCREATA GESCOE</p></strong> 
  <p>&nbsp;</p>
    </body>
  </html>
  """

  #part1 = MIMEText(text1, "plain")
  part2 = MIMEText(html, "html")
  #part3 = MIMEText(text2, "plain")

  #message.attach(part1)
  message.attach(part2)
  #message.attach(part3)

  mail.sendmail(sender, emails[i], message.as_string())
  mail.close()




