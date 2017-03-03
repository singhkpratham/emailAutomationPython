# -*- coding: utf-8 -*-
"""
Created on Fri Mar  3 15:44:45 2017

@author: kumar.singh
"""
path = "https://musigma.sharepoint.com/sites/DU5–Horizontal%20Initiatives/Shared%20Documents/Quality%20Initiatives/"


'''
https://musigma-my.sharepoint.com/personal/kumar_singh_mu-sigma_com/Documents/
'''
import win32com.client as win32
from datetime import datetime


spLink = r"https://musigma-my.sharepoint.com/personal/kumar_singh_mu-sigma_com/Documents/emails_muq.xlsx"

spLink = r'https://musigma.sharepoint.com/sites/DU5–Horizontal%20Initiatives/Shared%20Documents/Quality%20Initiatives/muQ%20status_03032017.xlsx?web=0'


path = 'https://musigma-my.sharepoint.com/personal/kumar_singh_mu-sigma_com/Documents/'
fileName = 'muQ' + str(datetime.now().date())  + '.xlsx'
saveTo = path + fileName
xl = win32.Dispatch("Excel.Application")
wb = xl.Workbooks.Open(spLink,  IgnoreReadOnlyRecommended = True , Editable = False)
wb.SaveAs(saveTo)
wb.Close()
xl.Quit()
df = pd.read_excel(saveTo)
os.remove(saveTo)


import urllib,cookielib



url = a

# MacOS
chrome_path = r'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'

# Windows
# chrome_path = 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe %s'

# Linux
# chrome_path = '/usr/bin/google-chrome %s'

b = webbrowser.get(chrome_path)#.open(url)
b.open(spLink)


from selenium import webdriver

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--no-startup-window")
driver = webdriver.Chrome(r"C:\Users\kumar.singh\Desktop\chromedriver.exe", chrome_options=chrome_options)
driver.get('https://www.google.com')

import subprocess
command = "start chrome /new-tab " + spLink
subprocess.Popen(command,shell=True)
