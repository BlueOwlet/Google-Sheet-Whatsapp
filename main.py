from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import json

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
    def __init__(self,teacher,group,level,schedule,feedback):
        self.teacher=teacher
        self.group=group
        self.level=level
        self.schedule=schedule
        self.feedback=feedback
        self.students=[]
    def addStudent(self,student):
        self.students.append(student)

import pyperclip
from selenium.webdriver.common.keys import Keys
class Message:
    def __init__(self, recipient,msg):
        # self.recipient=recipient
        self.recipient=recipient
        self.msg=msg
    def send(self):
        print('Sending Message')
        try:
            searchBox=browser.find_element_by_xpath("//div[@class='_3qx7_']")
        except Exception as e:
            resetSearchBoxButton=browser.find_element_by_xpath("//button[@class='MfAhJ']")
            resetSearchBoxButton.click()
            searchBox=browser.find_element_by_xpath("//div[@class='_3qx7_ _3fVV4']")

        searchBox.click()
        searchText=browser.find_element_by_xpath("//div[@class='_3FRCZ copyable-text selectable-text']")
        # searchText.click()
        searchText.send_keys(self.recipient)
        time.sleep(2)
        recipient=browser.find_element_by_xpath("//span[@title='{}']".format(self.recipient))
        recipient.click()
        time.sleep(2)
        messageBox=browser.find_element_by_xpath("//div[@class='_3uMse']")
        pyperclip.copy(self.msg)
        messageBox.send_keys(Keys.CONTROL+'v')
        time.sleep(1)
        sendButton=browser.find_element_by_xpath("//button[@class='_1U1xa']")
        sendButton.click()


def GetData(spreadsheet_id):
    # Call the Sheets API
    sheet = service.spreadsheets()
    global data
    range='A:Z'
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
                if 'IAG' in row[1] and row[0] != '':

                    teachers.append(row[0])
    except Exception as e:
        # print(teachers)
        print('Data in Array Storing Done')
    return teachers



def GetUniqueTeacher(teachers):
    allTeachers = teachers
    # input(set(allTeachers))
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
                    feedback=row[5]
                    groups.append(Group(teacher,group,level,schedule,feedback))

                    groups[-1].addStudent(row[3])

                    print('{} {} {} {} {}'.format(groups[-1].group,groups[-1].teacher,groups[-1].level,groups[-1].students,groups[-1].schedule))
                elif row[0]=='' and row[3] != '':
                    groups[-1].addStudent(row[3])
    except Exception as e:
        print(e)
        print('Class Creation Complete')

def CreateMessages():
    global whitelistTeachers
    for teacher in uniqueTeachers:
        print(teacher)
        msg='''Good morning {teacher}!
This is the automated weekly summary of your classes in order to follow up with new material for your groups please. Please let me know if you need the next lesson plan for any of the groups below.

SUMMARY:
'''.format(teacher=teacher)

        teacherGroups = [group for group in groups if group.teacher==teacher]
        # input(teacherGroups)
        for group in teacherGroups:
            msg+='''{}
Students: {}
Level: {}
Schedule: {}
Last Feedback: {}

'''.format(group.group,group.students,group.level,group.schedule,group.feedback)

        msg+="Please let me know if any of the schedules are outdated,\n All the best,\nInterAct"
        print(msg)
        # input(group)
        # print(group)
        message=Message(group.teacher,msg)
        try:
            if message.recipient not in whitelistTeachers:
                print('{} - Blacklisted'.format(message.recipient))
                # input('Blacklisted')
            else:
                # input('Whitelisted')
                print('Sending Message')
                print('{} - Whitelisted'.format(message.recipient))
                message.send()
                print('Message Sent!')
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
    # browser = webdriver.Chrome('/usr/bin/chromedriver',options=options)
    browser = webdriver.Chrome('chromedriver.exe',options=options)

def WhatsappLogin():
    browser.get('https://web.whatsapp.com/')
    time.sleep(10)

def CheckOS():
    import platform
    osName=platform.system()
    print('OS detected: {}'.format(osName))
    return osName

def Whitelist():
    global whitelistTeachers
    whitelist=True
    if whitelist == True:
        with open('whitelistTeachers.json') as jsonFile:
            whitelistTeachers = json.load(jsonFile)
            whitelistTeachers = list(whitelistTeachers["whitelist"])
            print('Whitelist: {}'.format(whitelistTeachers))


def GetApproval(spreadsheet_id):
    # Call the Sheets API
    global whitelistTeachers
    whitelistTeachers=[]
    sheet = service.spreadsheets()
    global dataHR
    range='A:Z'
    result = sheet.values().get(spreadsheetId=spreadsheet_id,range=range).execute()
    dataHR = result.get('values', [])
    firstAccess=True
    for row in dataHR:
        if row:
            if firstAccess==True:
                approvalIndex=row.index('Weekly Whatsapp')
                teacherIndex=row.index('Teacher')
                firstAccess=False
            try:
                if row[approvalIndex] in ['Yes','yes','y','Communicate','Send','Go']:
                    whitelistTeachers.append(row[0])
            except Exception as e:
                continue
    print('Whitelist: {}'.format(whitelistTeachers))



def main():
    mode='online'#local or online
    login()
    print('Login Successful')
    osName=CheckOS()
    if osName == 'Windows':
        driver='chromedriver.exe'
    else:
        driver='/usr/bin/chromedriver'
        input(driver)
    print('OS check: ',osName)

    spreadsheet_id='1gEsoxidCDUByobm0RU7nThlY_-EODVOTc2A6vT-8Sz8'
    teachers=GetData(spreadsheet_id)
    uniqueTeachers=GetUniqueTeacher(teachers)

    if mode=='online':
        spreadsheet_id='1ZHi2juyZwzZVd32jUyhe3W0KiVLM7oZEv9rd7v1Q9Wc'
        GetApproval(spreadsheet_id)
    else:
        Whitelist()

    print('Creating Messages')
    CreateClasses()
    StartBrowser()
    print('Browser Started')
    WhatsappLogin()
    CreateMessages()
    # browser.quit()




main()
