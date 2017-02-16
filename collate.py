import os
import pandas as pd
os.chdir('C:\\Users\\kumar.singh\\Desktop\\sharepoint')
from win32com.client import Dispatch
import re

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
df = spfetcher(spLink, saveTo)

a = df.count()

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


##collate = pd.DataFrame({
##    'FU' : [5],
##    'Response Time':[''],
##    '# of scoreboards | # Subgroups':['%s|%s'%(a[1],a[0])],
##    '# of errors this week|# of RCAs received by FUL':['%s|%s'%(df.ix[:,3].sum(),df.ix[:,5].sum()],
##})
##

