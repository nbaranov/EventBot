﻿#! /bin/python3
# -*- coding: utf-8 -*-

import datetime
import time
import os
import sys

from bs4 import BeautifulSoup as bs
from getpass import getpass
from random import randint
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options


######################################

login = "nikita.baranov" #input('Enter your login from Nokia portal: ')
password = "8OSmerkaNoKi1"  #getpass('Enter your password from Nokia portal (password will be hide): ')

executor_list = {
    u'nikita.baranov@nokia.com' : u'Баранов Никита Викторович',
    u'grigory.vlasov@nokia.com' : u'Власов Грирорий Юрьевич',
    u'ivan.shentcev@nokia.com' : u'Шенцев Иван Сергеевич',
    u'galina.kusurova@nokia.com' : u'Кусурова Галина Николаевна',
    u'roman.yuyukin@nokia.com' : u'Ююкин Роман Петрович',
    u'alexey.kolyadin@nokia.com' : u'Колядин Алексей Александрович',
    u'stanislav.baranov@nokia.com' : u'Баранов Станислав Евгеньевич',
    u'nikita.lychagin@nokia.com' : u'Лычагин Никита Вячеславович',
    u'dmitry.zaytcev@nokia.com' : u'Зайцев Дмитрий Андреевич',
    u'aleksandr.matveev@nokia.com' : u'Матвеев Александр Юрьевич',
    u'alexey.tolstoy@nokia.com' : u'Толстой Алексей Николаевич',
    u'mikhail.turzenok@nokia.com' : u'Турзенок Михаил Алексеевич',
    u'andrey.kuzminykh@nokia.com' : u'Кузминых Андрей Вячеславович',
#    u'maxim.rozenshtein@nokia.com' : u'Розенштейн Максим Яковлевич'
}

good_status = [
    'Назначен группе',
    'Назначен сотруднику',
    'В работе'
]

in_work = {}
unknowt_executor = []
######################################

    ### functions ###

# get all information of event
def getTable():
    table = driver.find_element_by_xpath('//*[@id="eventDetailsFormId:j_idt711"]/tbody')
    code = table.get_attribute("outerHTML")
    soup = bs(code, "lxml")
    table = soup.find_all('tr')
    return table    

# get event status
def getStatus(table):
    status = table[0].find_all('td')
    status = status[1].find('span').text 
    return status 

# get requestor email
def getEmail(table):
    email = table[3].find_all('td')
    email = email[5].text 
    return email 

# get time start work on event
def getDate(table):
    date = table[5].find_all('td')
    date = date[3].text
    year = int(date[6:10])
    month = int(date[3:5])
    day = int(date[0:2])
    hour = int(date[11:13].strip(":"))
    minutes = int(date[13:16].strip(":"))
    date = datetime.datetime(year, month, day, hour, minutes, 0, 0)
    return date

# check how time event in work
def checkTime(start, finish, delay_min):
    delta = finish - start
    minutes = int(delta.seconds / 60)
    if minutes > delay_min:
        return True

# time if format like in nokia portal
def timenow():
    return datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")

# write to log
def toLog(do, name, id):
    tolog = f'{timenow()} ивент {id} {do} {name}. https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/works/workOne.jsf?id={id}\n'
    with open("logs.log", "a", encoding="utf-8") as log:
        log.write(str(tolog))
    print(tolog)

# To appoint event to executor
def toExecutor(status, email, event):
    if status == good_status[0]:
        try:
            driver.find_element_by_id("eventDetailsFormId:repe:2:btn").click()
            time.sleep(3)
            executor_form = driver.find_element_by_id('changeEventStatusDlgForm:executorOneMenuId_label').click()
            time.sleep(1)
            executor_form = driver.find_elements_by_xpath('//*[@id="changeEventStatusDlgForm:executorOneMenuId_filter"]')
            executor_form[0].send_keys(executor_list[email])
            time.sleep(1)
            executor_form[0].send_keys(Keys.RETURN)
            #driver.find_element_by_id('changeEventStatusDlgForm:changeEventDialogActivityTypesId_label').click()
            #time.sleep(0.6)
            #activity_form = driver.find_element_by_xpath('/html/body/div[35]/div[1]/input')
            #time.sleep(0.6)
            #activity_form.send_keys(u"Creation/deletion of VPLS service")
            #time.sleep(0.6)
            #activity_form.send_keys(Keys.RETURN)
            time.sleep(0.6)
            driver.find_element_by_id('changeEventStatusDlgForm:changeEventStatusBtn').click()
            time.sleep(2)
            toLog("назначен на ", executor_list[email], event["id"])
        except:
            toLog("Произошла ошибка при назначении", "", event["id"])

# to take event in work
def toWork(email, event):
        try:
            driver.find_element_by_id('eventDetailsFormId:repe:3:btn').click()
            time.sleep(2)
            button = driver.find_elements_by_id('changeEventStatusDlgForm:startTimeCalendarId_input')
            #button[0].click()
            #time.sleep(0.3)
            button[0].send_keys(Keys.TAB)
            time.sleep(0.3)
            driver.find_element_by_id('changeEventStatusDlgForm:changeEventStatusBtn').click()
            time.sleep(2)
            in_work.update({email : [event["id"], getDate(getTable())]})
            toLog("взят в работу ", executor_list[email], event["id"])
        except:
            toLog("Произошла ошибка при взятии в работу", "", event["id"])

# to close event 
def toClose(email, event_id):
    table = getTable()

    status = getStatus(table)
    if status == "В работе":
        try:
            driver.find_element_by_id('eventDetailsFormId:repe:5:btn').click()
            time.sleep(3)
            button = driver.find_elements_by_id('changeEventStatusDlgForm:finishTimeCalendarId_input')
            #button[0].click()
            #time.sleep(0.3)
            button[0].send_keys(Keys.TAB)
            time.sleep(0.3)
            driver.find_element_by_id('changeEventStatusDlgForm:changeEventStatusBtn').click()
            time.sleep(2)
            toLog("закрыт на", executor_list[email], event_id)
        except:
            toLog("Произошла ошибка при закрытии", "", event_id)


# start browser
def startBrowser():
    if os.name == "nt":
        hide = webdriver.ChromeOptions()
        hide.headless = True
        driver = webdriver.Chrome(executable_path="Driver\windows.exe", options=hide)
        #driver = webdriver.Chrome(executable_path="Driver\windows.exe")
    else:
        hide = webdriver.FirefoxOptions()
        hide.headless = True
        #driver = webdriver.Firefox(executable_path='Driver/geckodriver', options=hide)
        driver = webdriver.Firefox(executable_path='Driver/geckodriver') #wisible browser for test        
    return driver

# autorization on nokia portal
def autorization(driver):
    driver.get('https://portal.voronezh.gdc.nokia.com/nsn-portal/index.jsf')
    time.sleep(4)
    login_form = driver.find_element_by_id('frmLogin:username')
    login_form.send_keys(login)
    time.sleep(0.2)
    pass_form = driver.find_element_by_id('frmLogin:password')
    pass_form.send_keys(password)
    driver.find_element_by_id('frmLogin:btnLogin').click()
    time.sleep(1)
    print("Выполнена авторизация")

def workWithEventCicle(a):
    global countFor
    global unknowt_executor
    try:
        for i in range(a, len(events_list)):
            countFor = i
            if events_list[i] not in unknowt_executor:
                workWithEvent(events_list[i])
    except:
        print(f"{timenow()} ошибка в цикле for продолжить с элемента {countFor}")
        print(sys.exc_info()[1])
        driver.refresh()
        workWithEventCicle(countFor)


def workWithEvent(event):
    driver.get(f'https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/works/workOne.jsf?id={event["id"]}')
    print(f'Открыт event {event["id"]}')
    time.sleep(1)

    table = getTable()
    status = getStatus(table)
    email = getEmail(table)

    if email not in executor_list.keys():
        unknowt_executor.append(event)
        return
    # if its our event to appoint to executor
    if status == good_status[0] and email in executor_list.keys():
        toExecutor(status, email, event)
        time.sleep(2)
        status = getStatus(getTable())
    # if event already appoint check event in work and if event in work more 5 minutes close is and to take new
    if status == good_status[1] and email in executor_list.keys():
        if email in in_work:
            if checkTime(in_work[email][1], datetime.datetime.now(), 5):
                driver.get(f'https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/works/workOne.jsf?id={in_work[email][0]}')
                time.sleep(1)
                toClose(email, in_work[email][0])
                in_work.pop(email)
                driver.get(f'https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/works/workOne.jsf?id={event["id"]}')
                time.sleep(1)
                toWork(email, event)
        else:
            toWork(email, event)
    #if event already in work, add its to "in_work" list and check time in work
    elif status == good_status[2] and email in executor_list.keys():
        if email not in in_work:
            in_work.update({email : [event["id"], getDate(getTable())]})
        elif email in in_work and in_work[email][0] != event["id"]:
            driver.get(f'https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/works/workOne.jsf?id={in_work[email][0]}')
            time.sleep(1)
            toClose(email, in_work[email][0])
            in_work.pop(email)
            driver.get(f'https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/works/workOne.jsf?id={event["id"]}')
            time.sleep(1)
            in_work.update({email : [event["id"], getDate(getTable())]})
        elif checkTime(getDate(table), datetime.datetime.now(), 120):
            toClose(email, event["id"])
            in_work.pop(email) 

    # get list of events
def getEventList(driver):
    driver.get('https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/works/works.jsf')
    time.sleep(2)
    driver.find_elements_by_xpath('//*[@id="eventsFormId:eventsTableId:j_id29"]/option[3]')[0].click()
    time.sleep(3)
    events_table = driver.find_element_by_id('eventsFormId:eventsTableId_data')
    events_table = events_table.get_attribute("outerHTML")

    soup = bs(events_table, "lxml")
    events_table = soup.find_all('tr')
    events_list = []

    for line in events_table:
        tags = line.find_all('td')
        if tags[2].text != "Ожидание":
            link = tags[3].find('a')
            span = tags[17].find('span')
            
            events_list.append({
                'status' : tags[2].text,
                'id' : link.get('title'),
                'executor' : span.text
            })
    return events_list    

def clearUnknowtExecutorList(unknowt_executor, events_list):
    for event in unknowt_executor:
        if event not in events_list:
            unknowt_executor.remove(event)    

driver = startBrowser()
autorization(driver)

while True:
    satrtTime = datetime.datetime.now()
    countFor = 0
    try:
        # check run browser
        try:
            driver.get_window_position(windowHandle='current')
        except:
            print(sys.exc_info()[1])
            driver = startBrowser()
            autorization(driver)

        try:
            events_list = getEventList(driver)
        except:
            autorization(driver)
            events_list = getEventList(driver)

        # Go events in revers list 
        events_list.reverse()
        clearUnknowtExecutorList(unknowt_executor, events_list)
        workWithEventCicle(countFor)
            
    except:
        print(f"{timenow()} ошибка в цикле while")
        print(sys.exc_info()[1])
        time.sleep(2) #пауза чтобы закрыть скрипт
        continue

    print(f"{timenow()} в работе:")
    for line in in_work:
        print(line, in_work[line])
    #print(f"{timenow()} Пропускаем:")
    #for line in unknowt_executor:
    #    print(line)
    #repeat after 5 minutes
    if not checkTime(satrtTime, datetime.datetime.now(), 5):
        time.sleep(300)
