import datetime
import email
import smtplib
import ssl
import string
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from random import sample

import openpyxl
import xlsxwriter

x = datetime.datetime.now()


shoulderEx = []
backEx = []
legEx = []
coreEx = []
movement = []
fullBody = []
i = 0
n = 0
weeks = []
exerciseWeek = []
exersises = []
noWeek = {}
wb = openpyxl.load_workbook('workouts.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')
main_exercises = [sheet['A1'].value, sheet['B1'].value,sheet['C1'].value,sheet['D1'].value]
bi_exercises = [sheet['E1'].value, sheet['F1'].value]

for Col in sheet['A2':'A10']:
    for objects in Col:
        shoulderEx.append(objects.value)

for Col in sheet['B2':'B10']:
    for objects in Col:
        backEx.append(objects.value)

for Col in sheet['C2':'C8']:
    for objects in Col:
        legEx.append(objects.value)

for Col in sheet['D2':'D13']:
    for objects in Col:
        coreEx.append(objects.value)

for Col in sheet['E2':'E8']:
    for objects in Col:
        movement.append(objects.value)

for Col in sheet['F2':'F8']:
    for objects in Col:
        fullBody.append(objects.value)

for i in range(12):
    exerciseWeek = []
    weeks = []
    if i % 2 == 0:
        exerciseWeek = sample(main_exercises, int(len(main_exercises)))
        exerciseWeek.append(bi_exercises[0])
    else:
        exerciseWeek = sample(main_exercises, int(len(main_exercises)))
        exerciseWeek.append(bi_exercises[1])
    for n in range(5):
        if exerciseWeek[n] == main_exercises[0]:
            exersises = (sample(shoulderEx,7))
            exerciseWeek[n] = [exerciseWeek[n]]
            exerciseWeek[n].append(exersises)
        elif exerciseWeek[n] == main_exercises[1]:
            exersises = (sample(backEx,7))
            exerciseWeek[n] = [exerciseWeek[n]]
            exerciseWeek[n].append(exersises)
        elif exerciseWeek[n] == main_exercises[2]:
            exersises = (sample(legEx,7))
            exerciseWeek[n] = [exerciseWeek[n]]
            exerciseWeek[n].append(exersises)
        elif exerciseWeek[n] == main_exercises[3]:
            exersises = (sample(coreEx,7))
            exerciseWeek[n] = [exerciseWeek[n]]
            exerciseWeek[n].append(exersises)
        elif exerciseWeek[n] == bi_exercises[0]:
            exersises = (sample(movement,7))
            exerciseWeek[n] = [exerciseWeek[n]]
            exerciseWeek[n].append(exersises)
        elif exerciseWeek[n] == bi_exercises[1]:
            exersises = (sample(fullBody,7))
            exerciseWeek[n] = [exerciseWeek[n]]
            exerciseWeek[n].append(exersises)
    weeks.append(exerciseWeek)
    noWeek["Week " + str(i+1)] = weeks

workbook = xlsxwriter.Workbook('Surf-Workout-12-Weeks-'+ x.strftime("%d-%b-%Y") + '.xlsx') 
bold = workbook.add_format({'bold': True})
cell_format = workbook.add_format()
cell_format.set_align('vcenter')
cell_format.set_align('center')
cell_format.set_text_wrap()
bold.set_align('vcenter')
bold.set_align('center')
bold.set_text_wrap()
bold.set_font_size(16)
for week in range(len(noWeek)):
    worksheet = workbook.add_worksheet("Week " + str(week + 1))
    worksheet.set_column(0, 4, 30)
    for day in range(len(noWeek["Week " + str(week+1)][0])):
        worksheet.write(string.ascii_uppercase[day] + '1', str(noWeek["Week " + str(week+1)][0][day][0]), bold)
        l = 2
        for work in range(len(noWeek["Week " + str(week+1)][0][0][1])):
            worksheet.write(string.ascii_uppercase[day] + str(l), str(noWeek["Week " + str(week+1)][0][day][1][work]), cell_format)
            l += 1


#print(noWeek["Week 2"][0][0][0])
#print(len(noWeek["Week 2"][0]))

#print(noWeek["Week " + str(week+1)][0][0][1][0]) This shows each individual exercise

workbook.close()

subject = "Workout Routine for 12 Weeks"
body = "Your fucking eamil dipshit for working out"
sender_email = "pythonjenkins@gmail.com"
receiver_email = "jamesljeffrey1995@gmail.com"
password = $EMAIL_PASS

# Create a multipart message and set headers
message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receiver_email
message["Subject"] = subject
message["Bcc"] = receiver_email  # Recommended for mass emails

# Add body to email
message.attach(MIMEText(body, "plain"))

filename = "Surf-Workout-12-Weeks-" + x.strftime("%d-%b-%Y") + ".xlsx" # In same directory as script

# Open PDF file in binary mode
with open(filename, "rb") as attachment:
    # Add file as application/octet-stream
    # Email client can usually download this automatically as attachment
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())

# Encode file in ASCII characters to send by email    
encoders.encode_base64(part)

# Add header as key/value pair to attachment part
part.add_header(
    "Content-Disposition",
    f"attachment; filename= {filename}",
)

# Add attachment to message and convert message to string
message.attach(part)
text = message.as_string()

# Log in to server using secure context and send email
context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, text)
