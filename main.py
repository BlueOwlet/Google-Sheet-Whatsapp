from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = [

        'https://www.googleapis.com/auth/spreadsheets.readonly',
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/drive.file',
        'https://www.googleapis.com/auth/drive'

        ]

def login():
    # creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json',SCOPES)
    # client = gspread.authorize(creds)
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    # global creds
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    global service
    service = build('sheets', 'v4', credentials=creds)

class Group:
    def __init__(self,teacher,group,level,schedule):
        self.teacher=teacher
        self.group=group
        self.level=level
        self.schedule=schedule
        self.students=[]
    def addStudent(self,student):
        self.students.append(student)

import pyperclip
from selenium.webdriver.common.keys import Keys
class Message:
    def __init__(self, recipient,msg):
        # self.recipient=recipient
        self.recipient='Jessica IA New'
        self.msg=msg
    def send(self):
        print('Sending Message')
        searchBox=browser.find_element_by_xpath("//div[@class='_3qx7_']")
        searchBox.click()
        time.sleep(1)
        searchBox.send_keys(self.recipient)
        searchText=browser.find_element_by_xpath("//div[@class='_3FRCZ copyable-text selectable-text']")
        searchText.send_keys(self.recipient)
        recipient=browser.find_element_by_xpath("//span[@title='{}']".format(self.recipient))
        recipient.click()
        messageBox=browser.find_element_by_xpath("//div[@class='_3uMse']")
        pyperclip.copy(self.msg)
        messageBox.send_keys(Keys.CONTROL+'v')
        time.sleep(1)
        # sendButton=browser.find_element_by_xpath("//button[@class='_1U1xa']")
        # sendButton.click()


def GetData():
    # Call the Sheets API
    spreadsheet_id = '1gEsoxidCDUByobm0RU7nThlY_-EODVOTc2A6vT-8Sz8'
    sheet = service.spreadsheets()
    global data
    range='A:E'
    result = sheet.values().get(spreadsheetId=spreadsheet_id,range=range).execute()
    data = result.get('values', [])
    i=0
    students=[]
    global teachers
    teachers=[]
    try:
        for row in data:
            # print('Checking if row exists')
            if row:
                # print('Row exists')
                if row[0] is not '':
                    teachers.append(row[0])
    except Exception as e:
        print('Data Collection Done')
    return teachers


def GetUniqueTeacher(teachers):
    allTeachers = teachers
    # print(set(allTeachers))
    global uniqueTeachers
    uniqueTeachers = list(set(allTeachers))
    uniqueTeachers.sort()
    if 'Teacher' in uniqueTeachers:
        uniqueTeachers.remove('Teacher')
    print('Unique Teachers Filtered')
    return uniqueTeachers


def CreateClasses():
    global groups
    groups=[]
    try:
        for row in data:
            # print('Checking if row exists')
            if row:
                if 'IAG' in row[1] and row[0] in uniqueTeachers:
                    print('-----------------Initiating new Group-----------------')

                    teacher = row[0]
                    group=row[1]
                    level=row[2]
                    student=row[3]
                    schedule=row[4]
                    groups.append(Group(teacher,group,level,schedule))

                    groups[-1].addStudent(row[3])

                    print('{} {} {} {} {}'.format(groups[-1].group,groups[-1].teacher,groups[-1].level,groups[-1].students,groups[-1].schedule))
                elif row[0]=='' and row[3] != '':
                    groups[-1].addStudent(row[3])
    except Exception as e:
        print(e)
        print('Class Creation Complete')

def CreateMessages():

    for teacher in uniqueTeachers:
        print(teacher)
        msg='''Good morning {},
This is the automated summary of your classes in order to follow up new material for your groups or any situation.

SUMMARY:

'''.format(teacher)

        teacherGroups = [group for group in groups if group.teacher==teacher]
        for group in teacherGroups:
            msg+='''

{}
Students: {}
Level: {}
Schedule: {}

'''.format(group.group,group.students,group.level,group.schedule)

        msg+='''Please let me know if you need new material or if any of the schedules are outdated,

Cheers'''

        print(msg)
        message=Message(group.teacher,msg)
        time.sleep(1)
        try:
            message.send()
        except Exception as e:
            print(e)
            print('Failed with teacher: '.format(group.teacher))




from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time


def StartBrowser():
    global options
    global browser
    options = Options()
    browser = webdriver.Chrome('/usr/bin/chromedriver',options=options)

def WhatsappLogin():
    browser.get('https://web.whatsapp.com/')
    time.sleep(10)




def main():
    print('Dowloading teacher list')
    print('Logging in and Getting data from Google')
    login()
    print('Login Successful')
    teachers=GetData()
    uniqueTeachers=GetUniqueTeacher(teachers)
    print('Creating Teacher Classes')
    print('Classes Created Successfully')
    print('Creating Messages')
    CreateClasses()
    StartBrowser()
    WhatsappLogin()
    CreateMessages()





main()
