import os
import pandas as pd
os.chdir('C:\\Users\\kumar.singh\\Desktop\\sharepoint')
from win32com.client import Dispatch
import re

spLink = r"https://musigma.sharepoint.com/sites/DU5–Horizontal%20Initiatives/Shared%20Documents/Quality%20Initiatives/muQ.xlsx?web=1 "

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

noRed   =          df.ix[:,10].sum()
noGreen =          df.ix[:,11].sum()
noTotalRedGreen =  noRed + noGreen
perRed  =          round(noRed/noTotalRedGreen*100)
perGreen =         round(noGreen/noTotalRedGreen*100)
noTotalDel =       df.ix[:,8].sum()
noGroups =         df.ix[:,0].count()
noErrors =         df.ix[:,3].sum()
noRCA    =         df.ix[:,5].sum()
noScoreb =         df.ix[:,1].count()
noLessThan2Weeks = df.ix[df.ix[:,2] < 2,2].count()
noMoreThan6Weeks = df.ix[df.ix[:,2] > 6,2].count()

collate = pd.DataFrame({
    'FU' : [5],
    'Response Time':[''],
    '# of scoreboards | # Subgroups':['%s|%s'%(noScoreb,noGroups)],
    '# of errors this week|# of RCAs reviewed by FUL':['%s|%s'%(noErrors,noRCA)],
    '# of groups < 2 weeks since last error (%)' : ['%s (%s%%)' %(noLessThan2Weeks, round(noLessThan2Weeks / df.ix[:,2].count()*100))],
    '# of groups > 6 weeks since last error (%)' : ['%s (%s%%)' %(noMoreThan6Weeks, round(noMoreThan6Weeks / df.ix[:,2].count()*100))],
    '# red (%)+ | # green (%)+ | # Tracked (%)++ | Total # ' : ['%s(%s%%)|%s(%s%%)|%s(%s%%)|%s' %( noRed,perRed ,noGreen  ,perGreen ,
                                                                                             noTotalRedGreen   ,round(noTotalRedGreen/noTotalDel*100)    ,noTotalDel ) ]                                                       
})

collate = collate[['FU','Response Time' ,'# of scoreboards | # Subgroups' , '# of errors this week|# of RCAs reviewed by FUL',
                   '# of groups < 2 weeks since last error (%)','# of groups > 6 weeks since last error (%)' ,'# red (%)+ | # green (%)+ | # Tracked (%)++ | Total # ' ]]

collate.to_csv("collated.csv", index = False)


