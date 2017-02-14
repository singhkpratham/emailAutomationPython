from win32com.client import Dispatch
import pandas as pd
import schedule
import os
import win32com.client as win32
from datetime import datetime
import time
#spLink = r'https://musigma-my.sharepoint.com/personal/kumar_singh_mu-sigma_com/Documents/debt.xlsx?web=1'
#spLink = r'https://musigma-my.sharepoint.com/personal/anantdeep_parihar_mu-sigma_com/Documents/muQ.xlsx?web=1'
spLink = r'https://musigma-my.sharepoint.com/personal/anantdeep_parihar_mu-sigma_com/Documents/mu.xlsx?web=1 '
saveTo = r'C:\Users\kumar.singh\Desktop\sharepoint\SP.xlsx'

def spfetcher(spLink, saveTo):
    xl = Dispatch("Excel.Application")
    #xl.Visible = True 

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
    mail.Subject = '[Reminder]Fill MuQ on Sharepoint immediately'
    #mail.body = 'WORK'
    mail.HTMLBody = body# this field is optional #add sharepoint link
    mail.send

#mailer(to = "kumar.singh@mu-sigma.com",body = "visit the link and fill it up %s" %(spLink))

df    = spfetcher(spLink, saveTo)             #fetching sharepoint table
email = pd.read_excel('emails_muq.xlsx')    #fetching table with email ids 
unsent = df.ix[pd.isnull(df.iloc[:,1:]).sum(axis  = 1)==13, 0] #which subgroups didnt update

emailsTo = email.ix[email.ix[:,0].isin(unsent),]   #df of email ids to send emails
emailsTo['All'] = emailsTo['AL'] + ';' +emailsTo['Team members']
#AlemailsToString = ";".join(list(emailsTo.ix[:,1]))
 
def mailToAl():
    for i in range(0,len(emailsTo)):
        body = "<p>Hi, Your team, %s, has missed the muQ deadline</p>.<p> Please fil it ASAP. </p>" %(emailsTo.iloc[i,0])
        mailer(body , emailsTo.iloc[i,1])


def mailToAll():
    for i in range(0,len(emailsTo)):
        body = "<p>Hi, Your team, %s, has missed the muQ deadline</p>.<p> Please fil it ASAP. </p>" %(emailsTo.iloc[i,0])
        mailer(body , emailsTo.iloc[i,3])





#datetime.now().time() > datetime.strptime("14:35" ,"%H:%M").time()



def starts():
    print('starts working')
    schedule.every(15).minutes.do(mail)
    

def mail():
    print('mail working')
    print(datetime.now().minute)
    if datetime.now().minute != 0:
        mailToAll()
    else:
        mailToAll()
        
schedule.every().tuesday.at("16:57").do(starts)

while True:
    schedule.run_pending()
    time.sleep(1)




















    

