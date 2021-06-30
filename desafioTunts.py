from __future__ import print_function
import os.path
import math
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials


from google.oauth2 import service_account


# Preparing to make an authorized API call
SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes = SCOPES)


# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1lE-2Q5oOCFZWlYRG6am-TkuUzAnT10FFZ4H68FAwb6U'
SPREADSHEET_RANGE_NAME = 'engenharia_de_software!A1:H27'


# Returns the total number of classes that year
def getNumClass(sheet):
    totalClassSheet = sheet.values().get(spreadsheetId = SPREADSHEET_ID, range='engenharia_de_software!A2:H2').execute()
    values = totalClassSheet.get('values', []) #Returns the string containing the information needed
    totalClassS = values[0][0].split(": ")
    numClass = int(totalClassS[1])
    return numClass

# Verifies if the student passed or not
def treatStudents(numStudents, numClass, sheet):
    for i in range(4, numStudents + 4, 1):
        position = str(i)
        studentRange = 'engenharia_de_software!C' +position + ':F' +position
        studentInfoSheet = sheet.values().get(spreadsheetId = SPREADSHEET_ID, range = studentRange ).execute()
        studentInfo = studentInfoSheet.get('values', [])

        #Checking if the student missed out on more than 25% of classes
        if ( int(studentInfo[0][0]) > math.ceil(numClass/4) ):
            situation = [['Reprovado por Falta', 0]]
            sheet.values().update(spreadsheetId = SPREADSHEET_ID, range = 'engenharia_de_software!G' +position,
                                    valueInputOption = 'USER_ENTERED', body = {'values': situation}).execute()
            continue

        checkStudentGrade(studentInfo, numClass, position, sheet)
        
    return

def checkStudentGrade(studentInfo, numClass, position, sheet):
    avaregeG = ( int(studentInfo[0][1]) + int(studentInfo[0][2]) + int(studentInfo[0][3]) ) / 3
    if (avaregeG < 50):
        situation = [['Reprovado por Nota', 0]]
        sheet.values().update(spreadsheetId = SPREADSHEET_ID, range = 'engenharia_de_software!G' +position,
                                valueInputOption = 'USER_ENTERED', body = {'values': situation}).execute()
        return
    elif (avaregeG >= 70):
        situation = [['Aprovado', 0]]
        sheet.values().update(spreadsheetId = SPREADSHEET_ID, range = 'engenharia_de_software!G' +position,
                                valueInputOption = 'USER_ENTERED', body = {'values': situation}).execute()
        return
    #In case of a Final Exam
    situation = [['Exame Final', math.ceil(100 - avaregeG)]]
    sheet.values().update(spreadsheetId = SPREADSHEET_ID, range = 'engenharia_de_software!G' +position,
                            valueInputOption = 'USER_ENTERED', body = {'values': situation}).execute()
    return

def main():

    service = build('sheets', 'v4', credentials = creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    numStudents = 24
    numClass = getNumClass(sheet)

    treatStudents(numStudents, numClass, sheet)

if __name__ == '__main__':
    main()