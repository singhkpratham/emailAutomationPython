import os
import pandas as pd
os.chdir('C:\\Users\\kumar.singh\\Desktop\\sharepoint')
from win32com.client import Dispatch
import re

#spLink = r"https://musigma.sharepoint.com/sites/DU5–Horizontal%20Initiatives/Shared%20Documents/Quality%20Initiatives/muQ.xlsx?web=1 "
spLink = r'https://musigma.sharepoint.com/sites/DU5–Horizontal%20Initiatives/Shared%20Documents/Quality%20Initiatives/muQ%20status_03032017.xlsx?web=1'
#spLink = r'https://musigma-my.sharepoint.com/personal/anantdeep_parihar_mu-sigma_com/Documents/mu.xlsx?web=1 '
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
df = spfetcher(spLink, saveTo)


for i in range(0,df.shape[1]):    
    for j in range(0,df.shape[0]):        
        try:            
            df.iloc[j,i] =  int(re.sub("[^\w]","",df.ix[j,i]))
            print(j , i)
        except:
            try:
                if bool(re.search("NA|NO" , df.ix[j,i], re.I)):
                    df.ix[j,i] = 0
            except:
                pass

noRed   =          df.ix[:,'# Red'].sum()
noGreen =          df.ix[:,'# Green'].sum()
noTotalRedGreen =  noRed + noGreen
perRed  =          round(noRed/noTotalRedGreen*100)
perGreen =         round(noGreen/noTotalRedGreen*100)
noTotalDel =       df.ix[:,'# Total Deliverables'].sum()
noGroups =         df.ix[:,0].count()
noErrors =         df.ix[:,'# Errors'].sum()
noRCA    =         df.ix[:,'# RCA shared with Leadership'].sum()
noScoreb =         df.ix[df['Is the scoreboard updated? (Yes/No)'].str.contains('yes|yo|haan' ,
                            case =False) == True ,'Is the scoreboard updated? (Yes/No)'].count()  #change
noLessThan2Weeks = df.ix[df.ix[:,'# Weeks without error'] < 2,'# Weeks without error'].count()
noMoreThan6Weeks = df.ix[df.ix[:,'# Weeks without error'] > 6,'# Weeks without error'].count()

collate = pd.DataFrame({
    'FU' : [5],
    'Response Time':[''],
    '# of scoreboards | # Subgroups':['%s|%s'%(noScoreb,noGroups)],
    '# of errors this week|# of RCAs received by FUL':['%s|%s'%(noErrors,noRCA)],
    '# of groups < 2 weeks since last error (%)' : ['%s (%s%%)' %(noLessThan2Weeks,
                                            round(noLessThan2Weeks / df.ix[:,2].count()*100))],
    '# of groups > 6 weeks since last error (%)' : ['%s (%s%%)' %(noMoreThan6Weeks,
                                            round(noMoreThan6Weeks / df.ix[:,2].count()*100))],
    '# red (%)+ | # green (%)+ | # Tracked (%)++ | Total # ' : ['%s(%s%%)|%s(%s%%)|%s(%s%%)|%s' %( noRed,
           perRed ,noGreen  ,perGreen , noTotalRedGreen   ,round(noTotalRedGreen/noTotalDel*100)    ,noTotalDel ) ]                                                       
})

collate = collate[['FU','Response Time' ,'# of scoreboards | # Subgroups' ,
                   '# of errors this week|# of RCAs received by FUL',
                   '# of groups < 2 weeks since last error (%)',
                   '# of groups > 6 weeks since last error (%)' ,
                   '# red (%)+ | # green (%)+ | # Tracked (%)++ | Total # ' ]]

collate.to_csv("collated.csv", index = False)

grp = df.groupby('Account').sum()
 

a = df.ix[df['Is the scoreboard updated? (Yes/No)'].str.contains('yes|yo|haan' ,
                            case =False) == True ,0:3].groupby('Account').count()
a['Account'] = a.index

c = df.groupby('Account').sum().ix[:,['# Errors', '# RCA shared with Leadership','# Red',
                                      '# Green','# Total Deliverables',
                                      '# Deliverable tracked for Red vs Green' ]]
c['Account'] = c.index

e = df.ix[df['# Weeks without error'] < 2 , :].groupby("Account").count().loc[:,['# Total Deliverables']]
e['Account'] = e.index
f = df.ix[df['# Weeks without error'] > 6 , :].groupby("Account").count().ix[:,['# Weeks without error']]
f['Account'] = f.index

e = noWeeks = pd.merge(e , f ,    on = 'Account' , how = 'outer')

collated = pd.merge(a , pd.merge(c,e,on='Account',how='outer' ), on = 'Account',how='outer')
collated.fillna(0, inplace= True)
collated['# of scoreboards | # Subgroups'] = collated['Scoreboard'].map(str) + " | " + collated['Scoreboard'].map(str)
collated['# of errors this week|# of RCAs received by FUL'] = collated['# Errors'].map(str) + " | " + collated['# RCA shared with Leadership'].map(str)
collated['# of groups < 2 weeks since last error (%)'] = collated['# Total Deliverables_y'].map(str) + '(' + (collated['# Total Deliverables_y']/collated['Team']*100).map(str)   + ')'
collated['# of groups < 2 weeks since last error (%)'] = collated['# Weeks without error'].map(str) + '(' + (collated['# Weeks without error']/collated['Team']*100).map(str)   + ')'
collated['# red (%)+ | # green (%)+ | # Tracked (%)++ | Total # '] = (
        collated['# Red'].map(str) + '(' ')'
        )
























