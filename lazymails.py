''' THIS IS MY FIRST PYTHON GUI PROJECT. IF YOU WANT TO ADD SOME FEATURES AND CONTRIBUTE TO IT, FEEL FREE TO DO SO!
BASICALLY WHAT THIS PROGRAM DOES - IT SEND THE CUSTOMIZED 'NAME' MASS EMAILS USING AN EXCEL FILE WITH JUST ONE CLICK MAKING A LIL BIT OF OUR LIFE EASIER.'''

#IMPORTING MODULES

from tkinter import *
from tkinter import filedialog
import pandas as p
import smtplib , ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#END OF IMPORTING MODULES

#MAIN SCREEN INIT
master       = Tk()
master.title('LazyMails V1.0')
master.geometry("650x690")

#END OF MAIN SCREEN INIT

#open dialog to select file to send email
def browsefunc():
    global contacts
    filename = filedialog.askopenfilename()
    Label(master, text=filename, font=('Calibri', 11)).grid(row=4,sticky=W, padx=80,pady=10)
    try:
        data = p.read_excel(filename)
        contacts = p.DataFrame(data)
    except Exception as e:
        print(e)
        notif.config(text="Error! Please Select A Excel File.", fg="red")
        print("Error")

#END OF OPEN DIALOG

#instructions grid menu
def instructions_show():
    alabel.grid(row=9, sticky=W,padx=285)
    blabel.grid(row=10, sticky=W,padx=50)
    clabel.grid(row=11, sticky=W,padx=50)
    dlabel.grid(row=12, sticky=W,padx=50)

def instructions_hide():
    alabel.grid_forget()
    blabel.grid_forget()
    clabel.grid_forget()
    dlabel.grid_forget()

#END OF INSTRUCTIONS GRID

#sender function
def send():
    try:
        username = temp_username.get()
        password = temp_password.get()
        body = txt.get("1.0",END)
        if username=="" or password=="" or body=="":
            notif.config(text="All fields required", fg="red")
            return
        else:
            for i in range(len(contacts)):
                Name , Email = contacts.iloc[i] #unpacking the data
                port = 465
                smtp_server = "smtp.gmail.com"
                message = MIMEMultipart("alternative")
                message['To'] = Email
                message['Subject'] = temp_subject.get()
                finalmessage = "Hello {},\n{}".format(Name.split()[0],body) #you can change anything instead of string saying "Hello"
                message.attach(MIMEText(finalmessage,'html')) #it is a html type so if you want to add colors in your mail use html color tag
                text = message.as_string()
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(smtp_server,port,context=context) as server:
                    server.login(username,password)
                    server.sendmail(username,message["To"],text)
                    notif.config(text="Email has been sent successfully", fg="green")
                    for i in range(len(contacts)):
                        cur_label='label' + str(i)
                        cur_label = Label(master,text="Sent To: {}\n".format(contacts))
                        cur_label.grid(row=1,sticky=W,column=1)
                print("Sent To",Name)
            print("All Email Sent")
    except Exception as e:
        print(e)
        notif.config(text="Error Sending Email.",fg="red")

#END OF SENDER FUCNCTION

#RESET FUNCTION
def reset():
  usernameEntry.delete(0,'end')
  passwordEntry.delete(0,'end')
  subjectEntry.delete(0,'end')
  txt.delete("1.0",'end')
  notif.config(text=" CLEAN ",fg="blue")

#END OF RESET FUNCTION

#Labels

Label(master, text="LazyMails", font=('Calibri',15)).grid(row=0, sticky=N)
Label(master, text="Please use the form below to send an email", font=('Calibri',11)).grid(row=1, sticky=N, padx=5 ,pady=10)
Label(master, text="Email", font=('Calibri', 11)).grid(row=2,sticky=W, padx=10)
Label(master, text="Password", font=('Calibri', 11)).grid(row=3,sticky=W, padx=5)
Label(master, text="To", font=('Calibri', 11)).grid(row=4,sticky=W, padx=5)
Label(master, text="Subject", font=('Calibri', 11)).grid(row=5,sticky=W, padx=5)
Label(master, text="Body", font=('Calibri', 11)).grid(row=6,sticky=W, padx=5)
notif = Label(master, text="", font=('Calibri', 11),fg="red")
notif.grid(row=8,sticky=S)

#END OF LABELS

#TempStorage
temp_username = StringVar()
temp_password = StringVar()
temp_subject  = StringVar()
#temp_body     = StringVar()

#END OF TEMPSTORAGE

#Entries
usernameEntry = Entry(master, textvariable = temp_username,width=40,font=('Calibri',16))
usernameEntry.grid(row=2,column=0, padx=130,pady=10,sticky=N)
passwordEntry = Entry(master, show="*", textvariable = temp_password,width=40,font=('Calibri',16)) #if you are working alone you can remove 'show=*' to see your password
passwordEntry.grid(row=3,column=0, padx=80,pady=10,sticky=N)
subjectEntry  = Entry(master, textvariable = temp_subject,width=40,font=('Calibri',16))
subjectEntry.grid(row=5,column=0,padx=80,pady=10,sticky=N)
txt=Text(master,width=40,height=8,font=('Calibiri'))
txt.grid(row=6,column=0,padx=80,sticky=N)

#END OF ENTRIES

#Buttons
Button(master, text = "Send", command = send).grid(row=8,   sticky=W,  pady=15, padx=5)
Button(master, text = "Reset", command = reset).grid(row=8,  sticky=W,  padx=45, pady=40)
Button(master, text = "Open Excel File", command = browsefunc).grid(row=8,  sticky=W,  padx=85, pady=40)
Button(master, text = "Show Instructions", command = instructions_show).grid(row=8,  sticky=W,  padx=178, pady=40)
Button(master, text = "Hide Instructions", command = instructions_hide).grid(row=8,  sticky=W,  padx=285, pady=40)

#END OF BUTTONS

#instructions
alabel=Label(master, text="INSTRUCTIONS!", font=('Calibri',14))
blabel=Label(master, text="1.For To(Sender Address) Field Just Select Your Excel File & It Should Do Your Work.", font=('Calibri',12))
clabel=Label(master, text="2.You Can Send A Plain Email By Entering Your Data In Body Field.", font=('Calibri',12))
dlabel=Label(master, text="3.You Can Send A HTML(Designed) Email By Using Tags Given In ReadMe File.", font=('Calibri',12))

#END OF INSTRUCTIONS

#copyright
Label(master, text="Git - Lazyprogrammerrr", font=('Calibri',8)).grid(row=14, sticky=N) #MYGIT

mainloop()
