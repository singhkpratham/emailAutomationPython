from win32com.client import Dispatch
import pandas as pd
import schedule
import os
import win32com.client as win32
from datetime import datetime
import time

spLink = r'https://musigma-my.sharepoint.com/personal/anantdeep_parihar_mu-sigma_com/Documents/mu.xlsx?web=1 '
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
    mail.Subject = '[Reminder]Fill MuQ on Sharepoint immediately pythonw'
    mail.HTMLBody = body# this field is optional #add sharepoint link
    mail.send

#df    = spfetcher(spLink, saveTo)             #fetching sharepoint table
email = pd.read_excel('emails_muq.xlsx')    #fetching table with email ids
email['All'] = email['AL'] + ';' +email['Team members']

##def mailToAl():
##    for i in range(0,len(emailsTo)):
##        body = "<p>Hi, Your team, %s, has missed the muQ deadline</p>.<p> Please fil it ASAP. </p>" %(emailsTo.iloc[i,0])
##        mailer(body , emailsTo.iloc[i,1])
##
##
##def mailToAll():
##    for i in range(0,len(emailsTo)):
##        body = "<p>Hi, Your team, %s, has missed the muQ deadline</p>.<p> Please fil it ASAP. </p>" %(emailsTo.iloc[i,0])
##        mailer(body , emailsTo.iloc[i,3])

def starts():
#    print('starts working')
    schedule.every(1).minutes.do(mail)
    

def mail():
    df    = spfetcher(spLink, saveTo)
    unsent = df.ix[pd.isnull(df.iloc[:,1:]).sum(axis  = 1)==13, 0]
    emailsTo = email.ix[email.ix[:,0].isin(unsent),]
#    print('mail working')
#    print(datetime.now())
    if datetime.now().minute != 0:
        for i in range(0,len(emailsTo)):
            body = "<p>Hi, Your team, %s, has missed the muQ deadline</p>.<p> Please fil it ASAP. </p>" %(emailsTo.iloc[i,0])
            mailer(body , emailsTo.iloc[i,1])
    else:
        for i in range(0,len(emailsTo)):
            body = "<p>Hi, Your team, %s, has missed the muQ deadline</p>.<p> Please fil it ASAP. </p>" %(emailsTo.iloc[i,0])
            mailer(body , emailsTo.iloc[i,3])
        
schedule.every().tuesday.at("19:21").do(starts)

while True:
    schedule.run_pending()
    time.sleep(1)




















    

