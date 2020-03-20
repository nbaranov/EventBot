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


######################################
login = input('Enter your login from Nokia portal: ')
password = getpass('Enter your password from Nokia portal (password will be hide): ')

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
    u'aleksandr.matveev@nokia.com' : u'Матвеев Александр Юрьевич'
}

good_status = [
    'Назначен группе',
    'Назначен сотруднику',
    'В работе'
]

in_work = {}
######################################

    ### functions ###

# get all information of event
def getTable():
    table = driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/form/div[2]/div/table/tbody")
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
    tolog = f'{timenow()} ивент {id} {do} {name}. https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/events/eventOne.jsf?id={id}\n'
    with open("logs.txt", "a", encoding="utf-8") as log:
        log.write(str(tolog))

# To appoint event to executor
def toExecutor(status, email):
    if status == good_status[0]:
        try:
            driver.find_element_by_xpath('/html/body/div[4]/div[3]/div/form/div[1]/div/button[3]/span').click()
            time.sleep(3)
            executor_form = driver.find_element_by_id('changeEventStatusDlgForm:executorOneMenuId_label').click()
            time.sleep(1)
            executor_form = driver.find_elements_by_xpath('/html/body/div[35]/div[1]/input')
            executor_form[0].send_keys(executor_list[email])
            time.sleep(0.3)
            executor_form[0].send_keys(Keys.RETURN)
            driver.find_element_by_id('changeEventStatusDlgForm:changeEventDialogActivityTypesId_label').click()
            time.sleep(0.3)
            activity_form = driver.find_element_by_xpath('/html/body/div[36]/div[1]/input')
            time.sleep(0.3)
            activity_form.send_keys(u"Creation/deletion of VPLS service")
            time.sleep(0.3)
            activity_form.send_keys(Keys.RETURN)
            time.sleep(0.3)
            driver.find_element_by_id('changeEventStatusDlgForm:changeEventStatusBtn').click()
            time.sleep(2)
            toLog("назначен на ", executor_list[email], event["id"])
        except:
            toLog("Произошла ошибка при назначении", "", event["id"])
            toLog(sys.exc_info()[1],"","")

# to take event in work
def toWork(email):
        try:
            driver.find_element_by_id('eventDetailsFormId:repe:3:btn').click()
            time.sleep(2)
            button = driver.find_elements_by_id('changeEventStatusDlgForm:startTimeCalendarId_input')
            button[0].click()
            time.sleep(0.3)
            button[0].send_keys(Keys.TAB)
            time.sleep(0.3)
            driver.find_element_by_id('changeEventStatusDlgForm:changeEventStatusBtn').click()
            time.sleep(2)
            in_work.update({email : [event["id"], getDate(getTable())]})
            toLog("взят в работу ", executor_list[email], event["id"])
        except:
            toLog("Произошла ошибка при взятии в работу", "", event["id"])
            toLog(sys.exc_info()[1],"","")

# to close event 
def toClose(email, event_id):
    table = getTable()

    status = getStatus(table)
    if status == "В работе":
        try:
            driver.find_element_by_id('eventDetailsFormId:repe:5:btn').click()
            time.sleep(3)
            button = driver.find_elements_by_id('changeEventStatusDlgForm:finishTimeCalendarId_input')
            button[0].click()
            time.sleep(0.3)
            button[0].send_keys(Keys.TAB)
            time.sleep(0.3)
            driver.find_element_by_id('changeEventStatusDlgForm:changeEventStatusBtn').click()
            time.sleep(2)
            toLog("закрыт на", executor_list[email], event_id)
        except:
            toLog("Произошла ошибка при закрытии", "", event_id)
            toLog(sys.exc_info()[1],"","")


# start browser and autorization on nokia portal
def startBrowser():
    if os.name == "nt":
        driver = webdriver.Chrome("Driver\\windows.exe")
    else:
        driver = webdriver.Chrome('Driver/linux')
    driver.maximize_window()
    driver.get('https://portal.voronezh.gdc.nokia.com/nsn-portal/index.jsf')
    time.sleep(1)
    login_form = driver.find_element_by_id('frmLogin:username')
    login_form.send_keys(login)
    time.sleep(0.2)
    pass_form = driver.find_element_by_id('frmLogin:password')
    pass_form.send_keys(password)
    driver.find_element_by_id('frmLogin:btnLogin').click()
    time.sleep(1)
    return driver


driver = startBrowser()

while True:
    try:
        # check run browser
        try:
            driver.get_window_position(windowHandle='current')
        except:
            print(sys.exc_info()[1])
            driver = startBrowser()

        # get list of events
        driver.get('https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/events/events.jsf')
        time.sleep(2)
        lis = driver.find_elements_by_xpath("/html/body/div[4]/div[2]/div/form/div[2]/div[4]/select/option[3]")[0].click()
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

        # Go events in revers list 
        events_list.reverse()
        for event in events_list:
            driver.get(f'https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/events/eventOne.jsf?id={event["id"]}')
            time.sleep(1)

            table = getTable()
            status = getStatus(table)
            email = getEmail(table)

            # if its our event to appoint to executor
            if status == good_status[0] and email in executor_list.keys():
                toExecutor(status, email)
                time.sleep(2)
                status = getStatus(getTable())
            # if event already appoint check event in work and if event in work more 5 minutes close is and to take new
            if status == good_status[1] and email in executor_list.keys():
                if email in in_work:
                    if checkTime(in_work[email][1], datetime.datetime.now(), 5):
                        driver.get(f'https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/events/eventOne.jsf?id={in_work[email][0]}')
                        time.sleep(1)
                        toClose(email, in_work[email][0])
                        in_work.pop(email)
                        driver.get(f'https://portal.voronezh.gdc.nokia.com/nsn-portal/ims/events/eventOne.jsf?id={event["id"]}')
                        time.sleep(1)
                        toWork(email)
                else:
                    toWork(email)
            #if event already in work, add its to "in_work" list and check time in work
            elif status == good_status[2] and email in executor_list.keys():
                if email not in in_work:
                    in_work.update({email : [event["id"], getDate(getTable())]})
                elif checkTime(getDate(table), datetime.datetime.now(), 120):
                    toClose(email, event["id"])
                    in_work.pop(email)
    except:
        print(sys.exc_info()[1])
        print(f"{timenow()} ошибка в цикле")

    #repeat after 5 minutes
    print(f"{timenow()} в работе:")
    for line in in_work:
        print(line, in_work[line])
    time.sleep(300)
    
      