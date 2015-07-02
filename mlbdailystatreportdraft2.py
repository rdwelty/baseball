#!/usr/bin/env python3

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import win32com.client as win32
from bs4 import BeautifulSoup
import re

TARGET = 'http://www.rotowire.com'
TARGET2 = 'http://www.rotowire.com/baseball/player_stats.htm'
TARGET3 = 'http://www.rotowire.com/baseball/player_stats.htm?pos=P'
USERNAME = 'rdwelty'
PASSWORD = 'rdwroto1'
VALUE_SHEET = r"C:\Users\RYAN\Anaconda\envs\py34\Baseball\07012015_Value.xlsx"
STAT_TEMPLATE = r"C:\Users\Ryan\Anaconda\envs\py34\Baseball\DayStatTemplate2.xlsx"

FIRSTROW = 3
LASTCOLUMN = 18

error_log = []
Player = ""

class logtableEmptyError(Exception):
    #error_log.append("logtable Empty: " + Player)
    pass
    


excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True

wb = excel.Workbooks.Open(VALUE_SHEET)          #|
ws = wb.Worksheets("Sheet1")                    #|
wsp = wb.Worksheets("P")                        #|
wsc = wb.Worksheets("C")                        #|
ws1b = wb.Worksheets("1B")                      #|
ws2b = wb.Worksheets("2B")                      #|-  Holder variables for Values workbook and worksheets
wsss = wb.Worksheets("SS")                      #|
ws3b = wb.Worksheets("3B")                      #|
wsof = wb.Worksheets("OF")                      #|

statwb = excel.Workbooks.Open(STAT_TEMPLATE)    #|
statwsp = statwb.Worksheets("P")                #|
statwsc = statwb.Worksheets("C")                #|
statws1b = statwb.Worksheets("1B")              #|
statws2b = statwb.Worksheets("2B")              #|-  Holder variables for Stats workbook and worksheets
statwsss = statwb.Worksheets("SS")              #|
statws3b = statwb.Worksheets("3B")              #|
statwsof = statwb.Worksheets("OF")              #|


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

####  FIND PLAYER STATS    ####



####    PITCHERS    ####

#br.find_element_by_link_text('Pitchers').click()
br.get(TARGET3)

wb.Activate()
wsp.Select()

LastRow = FIRSTROW
while (str(wsp.Cells(LastRow, 1).Value)) != "None":
    LastRow +=1
    
for i in range(FIRSTROW, LastRow, 1):
    Pos = wsp.Cells(i, 1).Value
    PlayerID = str(wsp.Cells(i, 2).Value)
    PlayerID = PlayerID.split('.')[0]
    Player = wsp.Cells(i, 3).Value
    Player = Player.rstrip()
    Team = wsp.Cells(i, 4).Value
    Bats = wsp.Cells(i, 5).Value
    Throws = wsp.Cells(i, 6).Value
    
    player_string = str('/baseball/player.htm?id=%s' % PlayerID)
    print(player_string)
    
    
    try:
        br.find_element_by_css_selector("a[href$='%s']" % PlayerID).click()
    except NoSuchElementException:
        print("NoSuchElementException")
        error_log.append("NoSuchElementException: " + Player)
        continue
        
    player_html = br.page_source
    player_soup = BeautifulSoup(player_html)
    #logtable = player_soup.find(id="gamelog")
    #if logtable == "None":
    #    logtable = player_soup.find(id="gamelog")
    #if logtable =="None":
    #    error_log.append("Logtable Empty: " + Player)
    #    continue
   
    try:
        logtable = player_soup.find(id="gamelog")
        if logtable == []:
            raise logtableEmptyError
            #break
    except logtableEmptyError:
        print("logtableEmptyError: " + Player)
        error_log.append("logtableEmptyError: " + Player)
        continue
   
    #row1 = logtable.find('tbody').find('tr')
    try:
        tbody = logtable.find('tbody')
    except AttributeError:
        print("AttributeError - logtable empty: ", Player)
        continue
    
    trows = tbody.find_all('tr')
    
    statwb.Activate()
    statwsp.Select()
    print("Active Workbook = ",excel.ActiveWorkbook.Name)
    for item in trows:
        if "didnotpitch" in str(item): continue
        else:
            trow1 = item
            break
            
    tdlist = trow1.find_all('td')    
    row = FIRSTROW
    while (str(statwsp.Cells(row, 1))) != "None":
        row +=1
    statwsp.Cells(row, 1).Value = Pos
    statwsp.Cells(row, 2).Value = PlayerID
    statwsp.Cells(row, 3).Value = Player
    statwsp.Cells(row, 4).Value = Team
    statwsp.Cells(row, 5).Value = Bats
    statwsp.Cells(row, 6).Value = Throws
    for index3, item3 in enumerate(tdlist, start=7):
        statwsp.Cells(row, index3).Value = item3.text
        
    del player_html
    del player_soup
    del logtable
    br.back()
    wb.Activate()
    
####    C    ######

br.get(TARGET2)

wb.Activate()
wsc.Select()

LastRow = FIRSTROW
while (str(wsc.Cells(LastRow, 1).Value)) != "None":
    LastRow +=1
    
for i in range(FIRSTROW, LastRow, 1):
    Pos = wsc.Cells(i, 1).Value
    PlayerID = str(wsc.Cells(i, 2).Value)
    PlayerID = PlayerID.split('.')[0]
    Player = wsc.Cells(i, 3).Value
    Player = Player.rstrip()
    Team = wsc.Cells(i, 4).Value
    Bats = wsc.Cells(i, 5).Value
    Throws = wsc.Cells(i, 6).Value
    
    player_string = repr("/baseball/player.htm?id=" + str(PlayerID))
    
    try:
        br.find(attrs={'href':player_string}).click()
    except NoSuchElementException:
        error_log.append("NoSuchElementException: " + Player)
        continue
    
    player_html = br.page_source
    player_soup = BeautifulSoup(player_html)
    
    try:
        logtable = player_soup.find(id="gamelog")
        if logtable == []:
            raise logtableEmptyError
            #break
    except logtableEmptyError:
        print("logtableEmptyError: " + Player)
        error_log.append("logtableEmptyError: " + Player)
        continue
    
    #row1 = logtable.find('tbody').find('tr')
    try:
        tbody = logtable.find('tbody')
    except AttributeError:
        print("AttributeError - logtable empty: ", Player)
        continue
    
    trow1 = tbody.find('tr')
    
    statwb.Activate()
    statwsc.Select()
    print("Active Workbook = ", excel.ActiveWorkbook.Name)
            
    tdlist = trow1.find_all('td')    
    row = FIRSTROW
    while (str(statwsc.Cells(row, 1))) != "None":
        row +=1
    statwsc.Cells(row, 1).Value = Pos
    statwsc.Cells(row, 2).Value = PlayerID
    statwsc.Cells(row, 3).Value = Player
    statwsc.Cells(row, 4).Value = Team
    statwsc.Cells(row, 5).Value = Bats
    statwsc.Cells(row, 6).Value = Throws
    for index3, item3 in enumerate(tdlist, start=7):
        statwsc.Cells(row, index3).Value = item3.text
        
    del player_html
    del player_soup
    del logtable
    br.back()
    wb.Activate()
    
####    1B    ######

br.get(TARGET2)

wb.Activate()
ws1b.Select()

LastRow = FIRSTROW
while (str(ws1b.Cells(LastRow, 1).Value)) != "None":
    LastRow +=1
    
for i in range(FIRSTROW, LastRow, 1):
    Pos = ws1b.Cells(i, 1).Value
    PlayerID = ws1b.Cells(i, 2).Value
    Player = ws1b.Cells(i, 3).Value
    Player = Player.rstrip()
    Team = ws1b.Cells(i, 4).Value
    Bats = ws1b.Cells(i, 5).Value
    Throws = ws1b.Cells(i, 6).Value
    
    player_string = repr("/baseball/player.htm?id=" + str(PlayerID))
    
    try:
        br.find(attrs={'href':player_string}).click()
    except NoSuchElementException:
        error_log.append("NoSuchElementException: " + Player)
        continue
    
    player_html = br.page_source
    player_soup = BeautifulSoup(player_html)
    
    try:
        logtable = player_soup.find(id="gamelog")
        if logtable == []:
            raise logtableEmptyError
            #break
    except logtableEmptyError:
        print("logtableEmptyError: " + Player)
        error_log.append("logtableEmptyError: " + Player)
        continue
    
    #row1 = logtable.find('tbody').find('tr')
    try:
        tbody = logtable.find('tbody')
    except AttributeError:
        print("AttributeError - logtable empty: ", Player)
        continue
    
    trow1 = tbody.find('tr')
    
    statwb.Activate()
    statws1b.Select()
    print("Active Workbook = ", excel.ActiveWorkbook.Name)
            
    tdlist = trow1.find_all('td')    
    row = FIRSTROW
    while (str(statws1b.Cells(row, 1))) != "None":
        row +=1
    statws1b.Cells(row, 1).Value = Pos
    statws1b.Cells(row, 2).Value = PlayerID
    statws1b.Cells(row, 3).Value = Player
    statws1b.Cells(row, 4).Value = Team
    statws1b.Cells(row, 5).Value = Bats
    statws1b.Cells(row, 6).Value = Throws
    for index3, item3 in enumerate(tdlist, start=7):
        statws1b.Cells(row, index3).Value = item3.text
        
    del player_html
    del player_soup
    del logtable
    br.back()
    wb.Activate()
    
    ####    2B    ######

br.get(TARGET2)

wb.Activate()
ws2b.Select()

LastRow = FIRSTROW
while (str(ws2b.Cells(LastRow, 1).Value)) != "None":
    LastRow +=1
    
for i in range(FIRSTROW, LastRow, 1):
    Pos = ws2b.Cells(i, 1).Value
    PlayerID = ws2b.Cells(i, 2).Value
    Player = ws2b.Cells(i, 3).Value
    Player = Player.rstrip()
    Team = ws2b.Cells(i, 4).Value
    Bats = ws2b.Cells(i, 5).Value
    Throws = ws2b.Cells(i, 6).Value
    
    player_string = repr("/baseball/player.htm?id=" + str(PlayerID))
    
    try:
        br.find(attrs={'href':player_string}).click()
    except NoSuchElementException:
        error_log.append("NoSuchElementException: " + Player)
        continue
    
    player_html = br.page_source
    player_soup = BeautifulSoup(player_html)
    
    try:
        logtable = player_soup.find(id="gamelog")
        if logtable == []:
            raise logtableEmptyError
            #break
    except logtableEmptyError:
        print("logtableEmptyError: " + Player)
        error_log.append("logtableEmptyError: " + Player)
        continue
    
    #row1 = logtable.find('tbody').find('tr')
    try:
        tbody = logtable.find('tbody')
    except AttributeError:
        print("AttributeError - logtable empty: ", Player)
        continue
    
    trow1 = tbody.find('tr')
    
    statwb.Activate()
    statws2b.Select()
    print("Active Workbook = ", excel.ActiveWorkbook.Name)
            
    tdlist = trow1.find_all('td')    
    row = FIRSTROW
    while (str(statws2b.Cells(row, 1))) != "None":
        row +=1
    statws2b.Cells(row, 1).Value = Pos
    statws2b.Cells(row, 2).Value = PlayerID
    statws2b.Cells(row, 3).Value = Player
    statws2b.Cells(row, 4).Value = Team
    statws2b.Cells(row, 5).Value = Bats
    statws2b.Cells(row, 6).Value = Throws
    for index3, item3 in enumerate(tdlist, start=7):
        statws2b.Cells(row, index3).Value = item3.text
        
    del player_html
    del player_soup
    del logtable
    br.back()
    wb.Activate()
    
####    SS    ######

br.get(TARGET2)

wb.Activate()
wsss.Select()

LastRow = FIRSTROW
while (str(wsss.Cells(LastRow, 1).Value)) != "None":
    LastRow +=1
    
for i in range(FIRSTROW, LastRow, 1):
    Pos = wsss.Cells(i, 1).Value
    PlayerID = wsss.Cells(i, 2).Value
    Player = wsss.Cells(i, 3).Value
    Player = Player.rstrip()
    Team = wsss.Cells(i, 4).Value
    Bats = wsss.Cells(i, 5).Value
    Throws = wsss.Cells(i, 6).Value
    
    player_string = repr("/baseball/player.htm?id=" + str(PlayerID))
    
    try:
        br.find(attrs={'href':player_string}).click()
    except NoSuchElementException:
        error_log.append("NoSuchElementException: " + Player)
        continue
    
    
    player_html = br.page_source
    player_soup = BeautifulSoup(player_html)
    
    try:
        logtable = player_soup.find(id="gamelog")
        if logtable == []:
            raise logtableEmptyError
            #break
    except logtableEmptyError:
        print("logtableEmptyError: " + Player)
        error_log.append("logtableEmptyError: " + Player)
        continue
    
    #row1 = logtable.find('tbody').find('tr')
    try:
        tbody = logtable.find('tbody')
    except AttributeError:
        print("AttributeError - logtable empty: ", Player)
        continue
    
    trow1 = tbody.find('tr')
    
    statwb.Activate()
    statwsss.Select()
    print("Active Workbook = ", excel.ActiveWorkbook.Name)
            
    tdlist = trow1.find_all('td')    
    row = FIRSTROW
    while (str(statwsss.Cells(row, 1))) != "None":
        row +=1
    statwsss.Cells(row, 1).Value = Pos
    statwsss.Cells(row, 2).Value = PlayerID
    statwsss.Cells(row, 3).Value = Player
    statwsss.Cells(row, 4).Value = Team
    statwsss.Cells(row, 5).Value = Bats
    statwsss.Cells(row, 6).Value = Throws
    for index3, item3 in enumerate(tdlist, start=7):
        statwsss.Cells(row, index3).Value = item3.text
        
    del player_html
    del player_soup
    del logtable
    br.back()
    wb.Activate()
    
    ####    3B    ######

br.get(TARGET2)

wb.Activate()
ws3b.Select()

LastRow = FIRSTROW
while (str(ws3b.Cells(LastRow, 1).Value)) != "None":
    LastRow +=1
    
for i in range(FIRSTROW, LastRow, 1):
    Pos = ws3b.Cells(i, 1).Value
    PlayerID = ws3b.Cells(i, 2).Value
    Player = ws3b.Cells(i, 3).Value
    Player = Player.rstrip()
    Team = ws3b.Cells(i, 4).Value
    Bats = ws3b.Cells(i, 5).Value
    Throws = ws3b.Cells(i, 6).Value
    
    player_string = repr("/baseball/player.htm?id=" + str(PlayerID))
    
    try:
        br.find(attrs={'href':player_string}).click()
    except NoSuchElementException:
        error_log.append("NoSuchElementException: " + Player)
        continue
    
    player_html = br.page_source
    player_soup = BeautifulSoup(player_html)
    
    try:
        logtable = player_soup.find(id="gamelog")
        if logtable == []:
            raise logtableEmptyError
            #break
    except logtableEmptyError:
        print("logtableEmptyError: " + Player)
        error_log.append("logtableEmptyError: " + Player)
        continue
    
    #row1 = logtable.find('tbody').find('tr')
    try:
        tbody = logtable.find('tbody')
    except AttributeError:
        print("AttributeError - logtable empty: ", Player)
        continue
    
    trow1 = tbody.find('tr')
    
    statwb.Activate()
    statws3b.Select()
    print("Active Workbook = ", excel.ActiveWorkbook.Name)
            
    tdlist = trow1.find_all('td')    
    row = FIRSTROW
    while (str(statws3b.Cells(row, 1))) != "None":
        row +=1
    statws3b.Cells(row, 1).Value = Pos
    statws3b.Cells(row, 2).Value = PlayerID
    statws3b.Cells(row, 3).Value = Player
    statws3b.Cells(row, 4).Value = Team
    statws3b.Cells(row, 5).Value = Bats
    statws3b.Cells(row, 6).Value = Throws
    for index3, item3 in enumerate(tdlist, start=7):
        statws3b.Cells(row, index3).Value = item3.text
        
    del player_html
    del player_soup
    del logtable
    br.back()
    wb.Activate()
    
    ####    OF    ######

br.get(TARGET2)

wb.Activate()
ws1b.Select()

LastRow = FIRSTROW
while (str(wsof.Cells(LastRow, 1).Value)) != "None":
    LastRow +=1
    
for i in range(FIRSTROW, LastRow, 1):
    Pos = wsof.Cells(i, 1).Value
    PlayerID = wsof.Cells(i, 2).Value
    Player = wsof.Cells(i, 3).Value
    Player = Player.rstrip()
    Team = wsof.Cells(i, 4).Value
    Bats = wsof.Cells(i, 5).Value
    Throws = wsof.Cells(i, 6).Value
    
    player_string = repr("/baseball/player.htm?id=" + str(PlayerID))
    
    try:
        br.find(attrs={'href':player_string}).click()
    except NoSuchElementException:
        error_log.append("NoSuchElementException: " + Player)
        continue
    
    player_html = br.page_source
    player_soup = BeautifulSoup(player_html)
    
    try:
        logtable = player_soup.find(id="gamelog")
        if logtable == []:
            raise logtableEmptyError
            #break
    except logtableEmptyError:
        print("logtableEmptyError: " + Player)
        error_log.append("logtableEmptyError: " + Player)
        continue
    
    #row1 = logtable.find('tbody').find('tr')
    try:
        tbody = logtable.find('tbody')
    except AttributeError:
        print("AttributeError - logtable empty: ", Player)
        continue
    
    trow1 = tbody.find('tr')
    
    statwb.Activate()
    statwsof.Select()
    print("Active Workbook = ", excel.ActiveWorkbook.Name)
            
    tdlist = trow1.find_all('td')    
    row = FIRSTROW
    while (str(statwsof.Cells(row, 1))) != "None":
        row +=1
    statwsof.Cells(row, 1).Value = Pos
    statwsof.Cells(row, 2).Value = PlayerID
    statwsof.Cells(row, 2).Value = Player
    statwsof.Cells(row, 3).Value = Team
    statwsof.Cells(row, 4).Value = Bats
    statwsof.Cells(row, 5).Value = Throws
    for index3, item3 in enumerate(tdlist, start=7):
        statwsof.Cells(row, index3).Value = item3.text
        
    del player_html
    del player_soup
    del logtable
    br.back()
    wb.Activate()