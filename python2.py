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
import re

os.chdir(r'C:\Users\kumar.singh\Desktop\sharepoint')

#spLink = r'https://musigma-my.sharepoint.com/personal/anantdeep_parihar_mu-sigma_com/Documents/mu.xlsx?web=1 '
#spLink = r"https://musigma.sharepoint.com/sites/DU5–Horizontal%20Initiatives/Shared%20Documents/Quality%20Initiatives/muQ.xlsx?web=1 "
spLink = r'https://musigma.sharepoint.com/sites/DU5–Horizontal%20Initiatives/Shared%20Documents/Quality%20Initiatives/muQ%20status_02242017.xlsx?web=1'
saveTo = r'C:\Users\kumar.singh\Desktop\sharepoint\SP.xlsx'
firstMailBody = """Hello All,
                    <p>After last Friday’s successful trial, we are trying to fully automate this process.
                    Please ensure you have the Quality Hour at 12 and update your scorecards at the following location:</p>
                    <p>%s</p>
                    Please fill only <strong>numbers</strong> in columns starting with the # symbol.
                    Avoid typing characters in these columns.
                    Also edit the excel only in <strong>browser</strong>, not in Excel Application.
                    <p>Thanks.</p>""" %(spLink)
reminderMailBody = """Hi, your team %s has missed the muQ deadline.Please update the scorecard
                         on the following link: <p>%s<p>If you're unable to update the
                                scorecard due to some reason, then reply to this mail with the subject '%s Unable to fill muQ'
                                and specify the reason in the mail body. Please copy the subject as it is.<p>Note: This is an
                                automatically generated mail that gets triggered every 15 minutes. To stop these mails please
                                either fill your scorecard or reply to this mail with the mail subject as specified above."""
FULemailid = "c.satish@mu-sigma.com"

email = pd.read_excel('emails_muq.xlsx')    #fetching table with email ids
email.ix[email['Team members'].isnull(),'Team members'] = " "
email['All'] = email['AL'] + ';' +email['Team members']

all_inbox = 0
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder("6")
all_inbox = inbox.Items

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
    mail.Subject = 'Update muQ on sharepoint!'
    mail.HTMLBody = body       # this field is optional 
    mail.Send()

def defaulters():
    print("inside defaulters")
    df    = spfetcher(spLink, saveTo)   #fetch the excel sheet
    df= df.ix[df['Team'].isnull() == False,:]
    unsent = df.ix[pd.isnull(df.iloc[:,3:]).sum(axis  = 1)==15, "Team"]
    return(df,unsent)

def keywordReplied():
    print("checking keywords")
    all_inbox.Sort("ReceivedTime", True)
    b = 0
    for i in range(0,len(all_inbox)):
        try:
            rec_time = all_inbox[i].ReceivedTime
        except:
            pass
        if (datetime(rec_time.year, rec_time.month, rec_time.day).date() == datetime.now().date() ):
            b +=1
        else:
            break
    mail_reply = list()
    for i in range(0,b):
        if bool(re.search("unable to fill muq",all_inbox[i].Subject,re.I) ):
            try:
                mail_reply.append(re.search(r'[\'\"]?(.*) Unable to fill muQ',all_inbox[i].Subject,re.I ).group(1))
            except:
                pass
    return(mail_reply)


def keywordAndUnsent():

    print('mail working')

    df , unsent = defaulters()
    mail_reply = keywordReplied()
    
    for i in range(0,len(df)):
        if (df.ix[i,'Team'] in mail_reply):
            df.loc[i,'can_reply'] = False
        else:
            df.loc[i,'can_reply'] = True
    emailsReqd = email.ix[email.ix[:,0].isin(unsent),]
    emailsReqd = pd.merge(emailsReqd, df[['Team','can_reply']],
                          left_on ="Subgroup name", right_on = "Team")
    emailsTo = emailsReqd.ix[emailsReqd.ix[:,0].isin(unsent) & emailsReqd['can_reply'],]
    return emailsTo

def firstMail():
    print('first mail sending')
    body = firstMailBody
#    to   = "; ".join(list(email.ix[:,1]))
#    to = to + "; Abhinav.Dasgupta@mu-sigma.com; Abhishek.Chopra@mu-sigma.com"
    to = 'kumar.singh@mu-sigma.com'
    mailer(body, to)

def mailToFUL(teamNameSeries):
    body = "Following teams haven't filled muQ yet:<p> %s" %("<p>".join(list(teamNameSeries)))
    mailer(body, FULemailid )

def reminderSender():
    emailsTo = keywordAndUnsent()
    for i in range(0,len(emailsTo)):
        if datetime.now().minute > 6:
            print('emails sent to  AL', emailsTo.iloc[i,1], 'from' ,emailsTo.iloc[i,0])
##            mailer(body , emailsTo.iloc[i,1])
            mailer(reminderMailBody %(emailsTo.iloc[i,0],spLink,emailsTo.iloc[i,0]) , 'kumar.singh@mu-sigma.com')
        else:        
            print('emails sent to  team and AL', emailsTo.iloc[i,1], 'from' ,emailsTo.iloc[i,0])
            mailer(reminderMailBody, 'anantdeep.parihar@mu-sigma.com')
  #          mailer(reminderMailBody , emailsTo.iloc[i,3])
  #          mailToFUL(emailsTo['Team'])
  

def starts():
    print('starts working at ' , datetime.now())
    schedule.every(15).minutes.do(reminderSender)
#schedule.every().friday.at("13:33").do(starts)
print(datetime.now())
#schedule.every().friday.at("12:45").do(firstMail)
##
##schedule.every().friday.at("13:00").do(starts)

##while True:
##    schedule.run_pending()
##    time.sleep(1)









