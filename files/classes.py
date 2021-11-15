from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from os import path, getcwd
import pandas as pd
import re
from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template
import requests
from files.functions import process


class Title:
    def __init__(self, frame):
        self.title = Label(frame, text="Placement Software",
                           width=20, font=("bold", 20))
        self.title.place(x=90, y=25)


class Fields:
    def __init__(self, frame, text, locx, locy):
        self.label = Label(frame, text=text, width=20, font=("bold", 10))
        self.label.place(x=locx, y=locy)
        self.var = StringVar(value="")
        self.entry = Entry(frame, textvariable=self.var)
        self.entry.place(x=locx + 160, y=locy)


class BranchCheckBox:
    def __init__(self, frame, text, locx, locy, command):
        self.var = IntVar()
        self.check = Checkbutton(
            frame, text=text, variable=self.var, command=command)
        self.check.place(x=locx, y=locy)


class BranchCheckBoxGroup:
    def __init__(self, frame, locx, locy):
        self.label = Label(frame, text="Branches")
        self.label.place(x=locx, y=locy)
        self.all = BranchCheckBox(
            frame, "ALL", locx-160, locy+20, self.changetoall)
        self.all.var.set(1)
        self.cse = BranchCheckBox(
            frame, "CSE", locx-80, locy+20, self.changenotall)
        self.ece = BranchCheckBox(
            frame, "ECE", locx, locy+20, self.changenotall)
        self.it = BranchCheckBox(frame, "IT", locx+80,
                                 locy+20, self.changenotall)
        self.eee = BranchCheckBox(
            frame, "EEE", locx+160, locy+20, self.changenotall)
        '''Made by kuldip and Brayan'''

    def changetoall(self):
        if self.all.var.get() == 1:
            self.cse.var.set(0)
            self.ece.var.set(0)
            self.it.var.set(0)
            self.eee.var.set(0)

    def changenotall(self):
        if self.all.var.get() == 1:
            self.all.var.set(0)


class CGPAPercCB:
    def __init__(self, frame, text, locx, locy, command):
        self.var = IntVar()
        self.check = Checkbutton(
            frame, text=text, variable=self.var, command=command)
        self.check.place(x=locx, y=locy)


class CgpaPerc:
    def __init__(self, frame, locx, locy):
        self.cgpa = CGPAPercCB(frame, "Btech CGPA",
                               locx-50, locy, self.changetocgpa)
        self.perc = CGPAPercCB(frame, "Btech %", locx +
                               50, locy, self.changetoperc)
        self.perc.var.set(1)

    def changetocgpa(self):
        if self.cgpa.var.get() == 1:
            self.perc.var.set(0)

    def changetoperc(self):
        if self.perc.var.get() == 1:
            self.cgpa.var.set(0)


class OnlyField:
    def __init__(self, frame, locx, locy):
        self.var = StringVar(value="")
        self.entry = Entry(frame, textvariable=self.var)
        self.entry.place(x=locx + 160, y=locy)


class Importer:

    def __init__(self, frame, text, command, locx, locy):
        self.button = Button(frame, text=text, width=15,
                             command=command, bg='brown', fg='white')
        self.button.place(x=locx, y=locy)
        self.label = Label(frame, text="", width=17,
                           bg='white', borderwidth=1, relief="sunken")
        self.label.place(x=locx + 135, y=locy+5)


class ERROR:
    def __init__(self, text):
        messagebox.showerror("ERROR", text)


class Software:

    location = ''
    email_loc = ''

    def __init__(self, frame):
        try:
            self.title = Title(frame)
            self.tenth = Fields(frame, "10th %", 80, 75)
            self.twelfth = Fields(frame, "12th %", 80, 112)
            self.diploma = Fields(frame, "Diploma %", 80, 150)
            self.btech = CgpaPerc(frame, 80, 187)
            self.btechMarks = OnlyField(frame, 80, 187)
            self.maxbacklog = Fields(frame, "Backlog Accepted", 80, 225)
            self.companyname = Fields(frame, "Company Name", 80, 262)
            self.branches = BranchCheckBoxGroup(frame, 215, 300)
            self.database = Importer(
                frame, 'Import DataBase', self.importFile, 105, 375)
            Button(frame, text='Submit', width=20, command=self.getvals,
                   bg='brown', fg='white').place(x=180, y=410)
            self.emaillist = Importer(
                frame, "Import Email List", self.importemaillist, 105, 465)
            Button(frame, text='Send Email', width=20, command=self.sendmaillist,
                   bg='brown', fg='white').place(x=180, y=512)
            Button(frame, text='Send SMS', width=20, command=self.sendsms,
                   bg='brown', fg='white').place(x=180, y=542)
            Button(frame, text='Developed By', width=20, command=self.dev,
                   bg='black', fg='white').place(x=180, y=575)
        except:
            ERROR("SOME ERROR OCCURED. PLEASE CONTACT DEVELOPMENT TEAM")

    def dev(self):
        messagebox.showinfo(
            "Team", "Placement Software has been developed by:\nIEEE MSIT")

    def importFile(self):
        self.location = filedialog.askopenfilename()
        self.database.label.config(text="..."+self.location[-16:])

    def getvals(self):
        m10 = self.tenth.entry.get()
        m12 = self.twelfth.entry.get()
        bt = self.btechMarks.entry.get()
        dip = self.diploma.entry.get()
        backs = self.maxbacklog.entry.get()
        company = self.companyname.entry.get()
        stream = ''
        if(self.branches.all.var.get() == 1):
            stream = 'CSE|IT|ECE|EEE'
        if(self.branches.cse.var.get() == 1):
            stream += 'CSE|'
        if(self.branches.it.var.get() == 1):
            stream += 'IT|'
        if(self.branches.ece.var.get() == 1):
            stream += 'ECE|'
        if(self.branches.eee.var.get() == 1):
            stream += 'EEE|'
        if(stream[-1] == '|'):
            stream = stream.rstrip('|')

        btech = ""
        if(self.btech.cgpa.var.get() == 1):
            btech = "cgpa"
        if(self.btech.perc.var.get() == 1):
            btech = "perc"

        if company == '':
            ERROR("Enter Company Name")
            return None

        if self.location == "":
            ERROR("Select a File")
            return None

        result = process(m10, m12, bt, btech, dip, backs,
                         company, self.location, stream)
        if result == "FILE ERROR":
            ERROR("Select a valid Excel or csv file.")
        elif result == "FILE OPEN":
            ERROR("Close all excel and csv files")
        else:
            messagebox.showinfo("Success", "File saved at "+result)

    def importemaillist(self):
        self.email_loc = filedialog.askopenfilename()
        self.emaillist.label.config(text="..."+self.email_loc[-16:])

    def sendmaillist(self):
        if self.email_loc == "":
            ERROR("Select Email List")
            return None
        else:
            filetype = self.email_loc.split(".")[-1]
            if filetype.upper() != 'CSV' and filetype.upper() != 'XLS' and filetype.upper() != 'XLSX' and filetype.upper() != 'XLSM':
                ERROR("Select a valid Excel or csv file.")
                return None

        data_e = pd.read_excel(self.email_loc)
        data_e_col = data_e.columns.values
        data_e_col = list(data_e_col)

        regex = re.compile(r'(e|E)(mail|Mail|MAIL)')
        email = list(filter(regex.match, data_e_col))
        email = "".join(email)
        regex = re.compile(r'(s|S)(tudent|TUDENT|)\s(N|n)(AME|ame)')
        sname = list(filter(regex.match, data_e_col))
        sname = "".join(sname)

        f = open(path.join(getcwd(), 'Templates', 'emailTemplate.txt'), 'r')
        t = f.read()
        t = Template(t)
        fromaddr = ""

        try:
            i = 0
            for person in data_e[sname]:
                message = t.substitute(PERSON_NAME=person)
                toaddr = data_e.at[i, email]
                msg = MIMEMultipart()
                msg['From'] = fromaddr
                msg['To'] = toaddr
                msg['Subject'] = "Placement Software"
                msg.attach(MIMEText(message, 'plain'))
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.ehlo()
                server.starttls()
                server.ehlo()

                server.login("", "")

                text = msg.as_string()
                server.sendmail(fromaddr, toaddr, text)
                i = i+1
            messagebox.showinfo("Success", "Mails Sent")
        except:
            ERROR("Check Internet Connection")
            return None

    def sendsms(self):
        if self.email_loc == "":
            ERROR("Select Email List")
            return None
        else:
            filetype = self.email_loc.split(".")[-1]
            if filetype.upper() != 'CSV' and filetype.upper() != 'XLS' and filetype.upper() != 'XLSX' and filetype.upper() != 'XLSM':
                ERROR("Select a valid Excel or csv file.")
                return None
        data_e = pd.read_excel(self.email_loc)
        data_e_col = data_e.columns.values
        data_e_col = list(data_e_col)
        regex = re.compile(
            r'(((M|m)(obile)\s(Number|number|No.|No|no))|((P|p)(hone)|((P|p)(h)(\.?))\s(Number|number|No.|No|no)))')
        mobile = list(filter(regex.match, data_e_col))
        mobile = "".join(mobile)
        numbers = []
        for i, r in data_e.iterrows():
            numbers.append(str(r[mobile]))

        numbers = ','.join(numbers)
        url = "https://www.fast2sms.com/dev/bulk"

        '''payloadformat = "sender_id=FSTSMS&message=test&language=english&route=p&numbers=9999999999,888888888"'''
        f = open(path.join(getcwd(), "Templates", "smsTemplate.txt"), 'r')
        message = f.read().rstrip("\n")
        payload = "sender_id=FSTSMS&message="+message + \
            "&language=english&route=p&numbers="+numbers
        headers = {
            'authorization': "",
            'Content-Type': "application/x-www-form-urlencoded",
            'Cache-Control': "no-cache",
        }
        try:
            response = requests.request(
                "POST", url, data=payload, headers=headers)
            messagebox.showinfo("Success", "SMS Sent")
            print(response.text)
        except:
            ERROR("Check Internet Connection")
            return None
