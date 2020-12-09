import pandas as pd
import datetime as dt
import win32com.client
#from selenium import webdriver
#from selenium.webdriver.common.by import By
#from selenium.webdriver.support.ui import WebDriverWait
#from selenium.webdriver.support import expected_conditions as EC
#import time
import os


begin_custom = dt.datetime(int(input("Beginning year:")),int(input("Beginning month: ")),int(input("Beginning day: ")))
end_custom = dt.datetime(int(input("Ending year:")),int(input("Ending month: ")),int(input("Ending day: ")))

#Get calendar object of all events
def get_calendar(begin,end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    return calendar

cal = get_calendar(begin_custom,end_custom)

#Create list of meetings
MeetingSubject = [meeting.subject for meeting in cal if meeting.ResponseStatus == 3 or meeting.ResponseStatus == 1]
MeetingStart = [meeting.start for meeting in cal if meeting.ResponseStatus == 3 or meeting.ResponseStatus == 1]
MeetingEnd = [meeting.end for meeting in cal if meeting.ResponseStatus == 3 or meeting.ResponseStatus == 1]
MeetingLocation = [meeting.location for meeting in cal if meeting.ResponseStatus == 3 or meeting.ResponseStatus == 1]

#Create dataframe of meeting lists
try:
    df = pd.DataFrame({'Subject': MeetingSubject,
                   'Start': MeetingStart,
                   'End': MeetingEnd,
                   'Location': MeetingLocation})
except:
    pass
#Split day and time
df['Start Date'] = pd.to_datetime(df['Start']).dt.date
df['Start Time'] = pd.to_datetime(df['Start']).dt.time
df['End Date'] = pd.to_datetime(df['End']).dt.date
df['End Time'] = pd.to_datetime(df['End']).dt.time

#export to csv
df.to_csv('MeetingLists.csv')

#In-progress: automatically upload to Google
#email = input("what is your email: ")
#password = input("what is your password: ")
#driver = webdriver.Chrome()
#driver.get("https://calendar.google.com/calendar/u/0/r/settings/export?tab=mc")

#driver.find_element_by_xpath('//*[@id="identifierId"]').send_keys(email)
#driver.find_element_by_xpath('//*[@id="identifierNext"]').click()
#time.sleep(0.5)
#driver.find_element_by_xpath('//*[@id="password"]/div[1]/div/div[1]/input').send_keys(password)
#driver.find_element_by_xpath('//*[@id="passwordNext"]').click()
#fileElem = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH , '//form[@jsname="GBqgNb"]//input')))
#fileElem.send_keys("MeetingLists.csv")



