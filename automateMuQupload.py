# -*- coding: utf-8 -*-
"""
Created on Fri Mar  3 15:44:45 2017

@author: kumar.singh
"""
'''
https://musigma-my.sharepoint.com/personal/kumar_singh_mu-sigma_com/Documents/
'''
import win32com.client as win32
from datetime import datetime
import pandas as pd
import os

spLink = r"C:\Users\kumar.singh\Desktop\sharepoint\muQ_FU5_Template.xlsx"

#path = 'https://musigma-my.sharepoint.com/personal/kumar_singh_mu-sigma_com/Documents/'
path = "https://musigma.sharepoint.com/sites/DU5–Horizontal%20Initiatives/Shared%20Documents/Quality%20Initiatives/"
fileName = 'muQ_status' + str(datetime.now().date())  + '.xlsx'
saveTo = path + fileName
xl = win32.Dispatch("Excel.Application")
wb = xl.Workbooks.Open(spLink)#  IgnoreReadOnlyRecommended = True , Editable = False)
wb.SaveAs(saveTo)
wb.Close()
xl.Quit()
#df = pd.read_excel(saveTo)
#os.remove(saveTo)

#
#import subprocess
#command = "start chrome /new-tab " + spLink
#subprocess.Popen(command,shell=True)
#
#import webbrowser
#chrome_path = r'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
#b = webbrowser.get(chrome_path)#.open(url)
#b.open(spLink)