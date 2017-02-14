import pandas as pd
import re
import numpy as np
from win32com.client import Dispatch

all_inbox = 0
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder("6")
all_inbox = inbox.Items

a = open('C:\\Users\\kumar.singh\\Desktop\\NTD1.html', "w")
b = all_inbox[2740].htmlbody
a.write(b)
a.close()

c = b.split("\r\n")[len(b.rsplit("\r\n"))-1] #to get just last line
d = re.findall("<table.*?<\\/table>",c)      #getting only tables out
e = ""
for i in d:                                  #joining all tables in 1 string
    e+=i
f = re.sub("<[sp].*?>|<\\/[sp](pan)?>","",e) #removing span and p  tags
g = re.sub("<o:.*?o:p>","", f)               #removing <o:p> tags
h = re.sub( "background:(.*?);.+?>(.+?)<\\/td>" ,
            ">\\1 \\2</td>" , g)             #putting background color in cell
i = re.sub(  "<([^/]\\w*)\\s.*?>",
             "<\\1>"  , h)                   #cleaning the html attributes
a = open("C:/Users/kumar.singh/Desktop/New folder/Git/practiceR/2.html", 'w')
a.write(i)
a.close
tb = pd.read_html("C:/Users/kumar.singh/Desktop/New folder/Git/practiceR/2.html"
                  , header = 0)

