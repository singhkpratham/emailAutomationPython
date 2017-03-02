# -*- coding: utf-8 -*-

"""
Created on Wed Feb 15 19:33:02 2017

@author: Kumar.Singh
"""

import pandas as pd
import schedule
import os
import win32com.client as win32
from datetime import datetime
import re
import time

os.chdir(r'C:\Users\kumar.singh\Desktop\sharepoint')

#spLink = r'https://musigma-my.sharepoint.com/personal/anantdeep_parihar_mu-sigma_com/Documents/mu.xlsx?web=1 '
#spLink = r"https://musigma.sharepoint.com/sites/DU5–Horizontal%20Initiatives/Shared%20Documents/Quality%20Initiatives/muQ.xlsx?web=1 "
spLink = r'https://musigma.sharepoint.com/sites/DU5–Horizontal%20Initiatives/Shared%20Documents/Quality%20Initiatives/muQ%20status_02242017.xlsx?web=1'
saveTo = r'C:\Users\kumar.singh\Desktop\sharepoint\SP.xlsx'
firstMailBody = """<font face="Calibri" >Hello All,
                    <p>After last Friday's successful trial, we are trying to fully automate this process.
                    Please ensure you have the Quality Hour at 12 and update your scorecards at the following location:</p>
                    <p>%s</p>
                    Please fill only <strong>numbers</strong> in columns starting with the # symbol.
                    Avoid typing characters in these columns.
                    Also edit the excel only in <strong>browser</strong>, not in Excel Application.
                    <p>Please note that there are extra columns to fill this time around (Q-hour summary shared, # utilities & # flows etc.)</p>
                    <p>Thanks.</p></font>""" %(spLink)
                    
reminderMailBody = """<font face="Calibri" >Hi, your team %s has missed the muQ deadline.Please update the scorecard
                         on the following link: <p>%s<p>If you're unable to update the
                                scorecard due to some reason, then reply to this mail with the subject '%s Unable to fill muQ'
                                and specify the reason in the mail body. Please copy the subject as it is.<p>Note: This is an
                                automatically generated mail that gets triggered every 15 minutes. To stop these mails please
                                either fill your scorecard or reply to this mail with the mail subject as specified above.</font>"""
FULemailid = "kumar.singh@mu-sigma.com"

email = pd.read_excel('emails_muq.xlsx')    #fetching table with email ids
email.ix[email['Team members'].isnull(),'Team members'] = " "
email['All'] = email['AL'] + '; ' +email['Team members']


def spfetcher(spLink, saveTo):
    print('fetching data from SP')
    xl = win32.Dispatch("Excel.Application")
    wb = xl.Workbooks.Open(spLink)
    wb.SaveAs(saveTo)
    wb.Close()
    xl.Quit()
    df = pd.read_excel(saveTo)
    os.remove(saveTo)
    return(df)

df    = spfetcher(spLink, saveTo)
email = email.ix[email.ix[:,0].isin(df['Team']),]

def mailer(body, to ):
    print('sending mail to:' , to)
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
    print('checking which team/s has replied with keyword')
    all_inbox = 0
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder("6")
    all_inbox = inbox.Items
    print("checking keywords")
    all_inbox.Sort("ReceivedTime", True)
    b = 0
    for i in range(0,len(all_inbox)): #finding number of emails received today
        try:
            rec_time = all_inbox[i].ReceivedTime
        except:
            pass
        if (datetime(rec_time.year, rec_time.month, rec_time.day).date() == datetime.now().date() ):
            b +=1
        else:
            break
    mail_reply = list()
    for i in range(0,b):              #finding names of Teams who have replied with keyword specified, appending in mail_reply
        if bool(re.search("unable to fill muq",all_inbox[i].Subject,re.I) ):
            try:
                mail_reply.append(re.search(r'[\'\"]?(.*) Unable to fill muQ',all_inbox[i].Subject,re.I ).group(1))
            except:
                pass
    print('teams replied with keyword are:' , mail_reply)
    return(mail_reply)


def keywordAndUnsent():

    print('keyword and unsent working')

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
    print('keywordAndUnsent exiting succesfully')
    return (emailsTo)

def firstMail():
    print('first mail sending at:', datetime.now())
    body = firstMailBody
    to   = "; ".join(list(email.ix[:,'All']))
    to = to + "; Abhinav.Dasgupta@mu-sigma.com; Abhishek.Chopra@mu-sigma.com"
#    to = 'kumar.singh@mu-sigma.com'
    mailer(body, to)

def mailToFUL(teamNameSeries):    
    print('sending mail to FUL')
    if len(teamNameSeries) != 0:
        body = "Following teams haven't filled muQ yet:<p> %s" %("<p>".join(list(teamNameSeries)))
    else :
        body = "All the teams have filled muQ"
    mailer(body, FULemailid )

def reminderSender():
    print('sending reminder started at ' , datetime.now())
    emailsTo = keywordAndUnsent()
    if len(emailsTo) == 0:
        mailToFUL(emailsTo['Team'])
        raise SystemExit()
    
    if datetime.now().minute > 4:
        for i in range(0,len(emailsTo)):
            print('emails sent to  AL', emailsTo.loc[i,'AL'], 'from' ,emailsTo.loc[i,'Subgroup name'])
            mailer(reminderMailBody %(emailsTo.iloc[i,0],spLink,emailsTo.iloc[i,0]) , #emailsTo.loc[i,'AL'])
                'kumar.singh@mu-sigma.com')
                
    else:
        mailToFUL(emailsTo['Team'])
        for i in range(0,len(emailsTo)):        
            print('emails sent to  team and AL', emailsTo.loc[i,"AL"], 'from' ,emailsTo.loc[i,'Subgroup name'])
            mailer(reminderMailBody %(emailsTo.iloc[i,0],spLink,emailsTo.iloc[i,0]), 'anantdeep.parihar@mu-sigma.com')
 #           mailer(reminderMailBody %(emailsTo.iloc[i,0],spLink,emailsTo.iloc[i,0]) , emailsTo.loc[i,'All'])
            
    print('sending reminder finished at ' , datetime.now())
  

def starts():
    print('starts function working at ' , datetime.now())
    schedule.every(15).minutes.do(reminderSender)


schedule.every().friday.at("14:00").do(starts)

print('script started at: ',datetime.now())

#schedule.every().friday.at("14:45").do(reminderSender )

#schedule.every().friday.at("11:30").do(firstMail)

#while True:
#    schedule.run_pending()
#    time.sleep(1)




a = """Hello All,
<font face="Comic sans MS" size="5">Verdana<p>After last Friday’s successful trial, we are trying to fully automate this process.
Please ensure you have the Quality Hour at 12 and update your scorecards at the following location:</p>
<p>%s</p></font>
Please fill only <strong>numbers</strong> in columns starting with the <strong>#</strong> symbol. Please refrain from typing characters in these columns. Also edit the excel only in <strong>browser</strong>, not in Excel Application.
<p>Please note that there are extra columns to fill this time around (Q-hour summary shared, # utilities & # flows etc.)</p>

<p>Thanks.</p>"""




