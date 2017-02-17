# -*- coding: utf-8 -*-
"""
Created on Wed Feb 15 19:33:02 2017

@author: Kumar.Singh
"""

from win32com.client import Dispatch
import pandas as pd
import schedule
import os
import win32com.client as win32
from datetime import datetime
import time
import re

os.chdir(r'C:\Users\kumar.singh\Desktop\sharepoint')

#spLink = r'https://musigma-my.sharepoint.com/personal/anantdeep_parihar_mu-sigma_com/Documents/mu.xlsx?web=1 '
spLink = r"https://musigma.sharepoint.com/sites/DU5–Horizontal%20Initiatives/Shared%20Documents/Quality%20Initiatives/muQ.xlsx?web=1 "
saveTo = r'C:\Users\kumar.singh\Desktop\sharepoint\SP.xlsx'



def spfetcher(spLink, saveTo):
    xl = Dispatch("Excel.Application")
    wb = xl.Workbooks.Open(spLink)
    wb.SaveAs(saveTo)
    wb.Close()
    xl.Quit()
    df = pd.read_excel(saveTo)
    os.remove(saveTo)
    return(df)

def mailer(body, to ):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
#   mail.Subject = '[Reminder]Fill MuQ on Sharepoint immediately'
    mail.Subject = 'Update muQ on sharepoint!'
    mail.HTMLBody = body# this field is optional #add sharepoint link
    mail.Send()

#df    = spfetcher(spLink, saveTo)             #fetching sharepoint table
email = pd.read_excel('emails_muq.xlsx')    #fetching table with email ids
email['All'] = email['AL'] + ';' +email['Team members']


all_inbox = 0
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder("6")
all_inbox = inbox.Items

all_inbox.Sort("ReceivedTime", True)

c = 0
for i in all_inbox:
    c+=1
b = 0
for i in range(1,c):
    rec_time = all_inbox[i].ReceivedTime
    if (datetime(rec_time.year, rec_time.month, rec_time.day) == datetime.now().date() ):
        b +=1
    else:
        break


def mail():
    all_inbox = 0
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder("6")
    all_inbox = inbox.Items

    all_inbox.Sort("ReceivedTime", True)

    c = 0
    for i in all_inbox:
        c+=1
    b = 0
    for i in range(1,c):
        try:
            rec_time = all_inbox[i].ReceivedTime
        except:
            pass
        if (datetime(rec_time.year, rec_time.month, rec_time.day).date() == datetime.now().date() ):
            b +=1
#            print(all_inbox[i].Subject)
        else:
            break
    print('mail working')
    df    = spfetcher(spLink, saveTo)
    unsent = df.ix[pd.isnull(df.iloc[:,1:]).sum(axis  = 1)==13, 0]
    mail_reply = list()
    for i in range(0,b):
#        print(bool(re.search("unable to fill muq",all_inbox[i].Subject,re.I)))
        if bool(re.search("unable to fill muq",all_inbox[i].Subject,re.I) ):
#            print(i)
#            print(re.search(r'(.*) Unable to fill muQ',all_inbox[i].Subject,re.I ).group(1))
            try:
                mail_reply.append(re.search(r'[\'\"]?(.*) Unable to fill muQ',all_inbox[i].Subject,re.I ).group(1))
            except:
                pass
    for i in range(0,len(email.ix[:,0].isin(unsent))):
        if (df.ix[i,0] in mail_reply):
            df.loc[i,'can_reply'] = False
        else:
            df.loc[i,'can_reply'] = True
    emailsTo = email.ix[email.ix[:,0].isin(unsent) & df['can_reply'],]
    
    print(datetime.now() , emailsTo)
    if datetime.now().minute > -1:        
        for i in range(0,len(emailsTo)):
            print('emails sent to  AL', emailsTo.iloc[i,1], 'from' ,emailsTo.iloc[i,0])
            body = """Hi, your team %s has missed the muQ deadline. Please update the scorecard on the following link: <p>%s </p>
                        Note: If you're unable to update the scorecard due to some reason, then reply to this mail with the subject
                        '%s Unable to fill muQ' and specify the reason in the mail body.""" %(emailsTo.iloc[i,0],spLink,emailsTo.iloc[i,0])
            mailer(body , emailsTo.iloc[i,1])
##    else:
##        for i in range(0,len(emailsTo)):
##            print('emails sent to  team and AL', emailsTo.iloc[i,1], 'from' ,emailsTo.iloc[i,0])
##            body = """Hi, your team %s has missed the muQ deadline. Please update the scorecard on the following link: <p>%s </p>
##                        Note: If you're unable to update the scorecard due to some reason, then reply to this mail with the subject
##                        '%s Unable to fill muQ' and specify the reason in the mail body.""" %(emailsTo.iloc[i,0],spLink,emailsTo.iloc[i,0])
##            mailer(body , emailsTo.iloc[i,3])

def starts():
    print('starts working at ' , datetime.now())
    schedule.every(15).minutes.do(mail)

#mail()

def firstMail():
    print('first mail sending')
    body = """Hello All,

            <p>Another Friday and we are trying to go automated! </p>
            <p>Please ensure you have the Quality Hour at 12 and update your scorecards at the following location:

            <p>Note: Please fill only <b> numbers</b> in columns starting with the # symbol. Please refrain from typing characters in these columns. Also edit the excel only in browser, not in Excel Application. <p>%s</p>

            Thanks""" %(spLink)
    to   = "; ".join(list(email.ix[:,1]))
#    to = 'kumar.singh@mu-sigma.com'
    mailer(body, to)


schedule.every().friday.at("13:33").do(starts)
print(datetime.now())
##schedule.every().friday.at("12:45").do(firstMail)
##
##schedule.every().friday.at("13:00").do(starts)

while True:
    schedule.run_pending()
    time.sleep(1)








