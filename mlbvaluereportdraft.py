#!/usr/bin/env python3

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import win32com.client as win32
from bs4 import BeautifulSoup
import re

TARGET = 'http://www.rotowire.com'
TARGET2 = 'http://www.rotowire.com/daily/mlb/value-report.htm'
USERNAME = 'rdwelty'
PASSWORD = 'rdwroto1'
VALUE_TEMPLATE = r"C:\Users\RYAN\Anaconda\envs\py34\Baseball\ValueTemplate2.xlsx"
#STAT_TEMPLATE = r"C:\Users\Ryan\Anaconda\envs\py34\Baseball\DayStatTemplate.xlsx"
FIRSTROW = 3
LASTCOLUMN = 18

htmlid_regex = re.compile(r'\d+')

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True


wb = excel.Workbooks.Open(VALUE_TEMPLATE)       #|
ws = wb.Worksheets("Sheet1")                    #|
wsp = wb.Worksheets("P")                        #|
wsc = wb.Worksheets("C")                        #|
ws1b = wb.Worksheets("1B")                      #|
ws2b = wb.Worksheets("2B")                      #|-  Holder variables for Values workbook and worksheets
wsss = wb.Worksheets("SS")                      #|
ws3b = wb.Worksheets("3B")                      #|
wsof = wb.Worksheets("OF")                      #|


br = webdriver.Firefox()
br.get(TARGET)
#time.sleep(5)
username = br.find_element_by_name('username')
username.send_keys(USERNAME)

#short delay
#time.sleep(5)
password = br.find_element_by_name('p1')
password.send_keys(PASSWORD)

#short delay
#time.sleep(5)
br.find_element_by_name('Submit').click()

#delay here
#time.sleep(5)
#br.find_element_by_link_text('MLB').click()
#br.find_element_by_link_text('All MLB Pages').click()
#br.find_element_by_link_text('Try Them Out').click()
#br.find_element_by_link_text('MLB Value Report').click()

br.get(TARGET2)

html = br.page_source
soup = BeautifulSoup(html)

table = soup.find("table")
rows = table.findAll('tr')
del rows[:2]


for index1, item1 in enumerate(rows, start=3):
    
    #idraw = item1.find()
    #ws.Cells(index1, 1).Value = item1.
    cols = item1.find_all('td')
    for index2, item2 in enumerate(cols, start=1):
        if index2 == 1:
            ws.Cells(index1, index2).Value = item2.getText()
        if index2 == 2:
            player_id = htmlid_regex.search(repr(str(item2)))
            if player_id:
                ws.Cells(index1, index2).Value = player_id.group()
            ws.Cells(index1, (index2+1)).Value = item2.getText()
        else:
            ws.Cells(index1, (index2+1)).Value = item2.getText()
        #ws.Cells(index1, index2).Value = item2[0].getText()
        #ws.Cells(index1, index2).Value = item2[1].find('a').getText()
        #ws.Cells(index1, index2).Value = item2[2].getText()
        #ws.Cells(index1, index2).Value = item2[3].getText()
        #ws.Cells(index1, index2).Value = item2[4].getText()
        #ws.Cells(index1, index2).Value = item2[5].getText()
        #ws.Cells(index1, index2).Value = item2[6].find('a').getText()
        #ws.Cells(index1, index2).Value = item2[7].getText()
        #ws.Cells(index1, index2).Value = item2[8].getText()
        #ws.Cells(index1, index2).Value = item2[9].getText()
        #ws.Cells(index1, index2).Value = item2[10].getText()
        #ws.Cells(index1, index2).Value = item2[11].getText()
        #ws.Cells(index1, index2).Value = item2[12].getText()
        #ws.Cells(index1, index2).Value = item2[13].getText()
        #ws.Cells(index1, index2).Value = item2[14].getText()
    
    
LastRow = FIRSTROW
while (str(ws.Cells(LastRow, 1).Value)) != "None":
    LastRow +=1

####  CUSTOM CALCULATIONS   ####    
for i in range(FIRSTROW, LastRow, 1):
    ws.Cells(i, 18).Value = float(ws.Cells(i, 11).Value) / float(ws.Cells(i, 10).Value)
    

    
####  POSITION VALUE SORT    ####

for j in range(FIRSTROW, LastRow, 1):
    row = FIRSTROW
    ws.Range(ws.Cells(j, 1), ws.Cells(j, LASTCOLUMN)).Copy()
    if (str(ws.Cells(j, 1))) == "P":
        while (str(wsp.Cells(row, 1))) != "None":
            row +=1
        #wsp.Cells(1,1).Value = "1"
        wsp.Select()
        #wsp.Range(wsp.Cells(row, 1)).Select()
        wsp.Cells(row, 1).Select()
        wsp.Paste()
    elif (str(ws.Cells(j, 1))) == "C":
        while (str(wsc.Cells(row, 1))) != "None":
            row +=1
        wsc.Select()
        wsc.Cells(row, 1).Select()
        wsc.Paste()
    elif (str(ws.Cells(j, 1))) == "1B":
        while (str(ws1b.Cells(row, 1))) != "None":
            row +=1
        ws1b.Select()
        ws1b.Cells(row, 1).Select()
        ws1b.Paste()
    elif (str(ws.Cells(j, 1))) == "2B":
        while (str(ws2b.Cells(row, 1))) != "None":
            row +=1
        ws2b.Select()
        ws2b.Cells(row, 1).Select()
        ws2b.Paste()
    elif (str(ws.Cells(j, 1))) == "SS":
        while (str(wsss.Cells(row, 1))) != "None":
            row +=1
        wsss.Select()
        wsss.Cells(row, 1).Select()
        wsss.Paste()
    elif (str(ws.Cells(j, 1))) == "3B":
        while (str(ws3b.Cells(row, 1))) != "None":
            row +=1
        ws3b.Select()
        ws3b.Cells(row, 1).Select()
        ws3b.Paste()
    elif (str(ws.Cells(j, 1))) == "OF":
        while (str(wsof.Cells(row, 1))) != "None":
            row +=1
        wsof.Select()
        wsof.Cells(row, 1).Select()
        wsof.Paste()
        

#def DailyPlayerStat(Player):
#    br.find_element_by_link_text(Player).click()
#    html = br.page_source
#    soup = BeautifulSoup(html)
#    gamelog = soup.find()
#    br.back()
#delay here
#time.sleep(4)
