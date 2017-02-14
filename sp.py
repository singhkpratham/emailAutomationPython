from win32com.client import Dispatch
import pandas as pd

import os
import win32com.client as win32
#spLink = r'https://musigma-my.sharepoint.com/personal/kumar_singh_mu-sigma_com/Documents/debt.xlsx?web=1'
#spLink = r'https://musigma-my.sharepoint.com/personal/anantdeep_parihar_mu-sigma_com/Documents/muQ.xlsx?web=1'
spLink = r'https://musigma-my.sharepoint.com/personal/anantdeep_parihar_mu-sigma_com/Documents/mu.xlsx?web=1 '
saveTo = r'C:\Users\kumar.singh\Desktop\sharepoint\SP.xlsx'

def fetcher(spLink, saveTo):
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

#mailer(to = "anantdeep.parihar@mu-sigma.com",body = "this works")

df    = fetcher(spLink, saveTo)             #fetching sharepoint table
email = pd.read_excel('emails_muq.xlsx')    #fetching table with email ids 
unsent = df.ix[pd.isnull(df.iloc[:,1:]).sum(axis  = 1)>0, 0] #which subgroups didnt update

emailsTo = email.ix[email.ix[:,0].isin(unsent),]   #df of email ids to send emails

emailsToString = ";".join(list(emailsTo.ix[:,1])     #string of Al's id who will be sent emails
    
                  



