import smartsheet
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os
import datetime
import string
class GoogleCalendarService:
    def __init__(self):
        # If modifying these scopes, delete the file token.json.
        SCOPES = ['https://www.googleapis.com/auth/calendar']

        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        self.service = build('calendar', 'v3', credentials=creds)
def generate_id(s):
    s = [letter for letter in s]
    possibleVals = list(string.ascii_uppercase + string.ascii_lowercase + string.digits + string.punctuation + " ")
    idList = []
    for letter in s:
        try:
            idList.append(str(possibleVals.index(letter)))
        except:
            idList.append("00000000")
    return ''.join(idList)
def grab_event_changes(newEventList, oldEventList):
    oldEventDict = {x['summary']: x for x in oldEventList}
    newEventDict = {x['summary']: x for x in newEventList}
    del oldEventList
    del newEventList
    newEvents = []
    deletedEvents = []
    updatedEvents = []

    # Finding new events
    for event in newEventDict.keys():
        try:
            oldEventDict[event]
        except:
            newEvents.append(newEventDict[event])

    # Finding deleted events
    for event in oldEventDict.keys():
        try:
            newEventDict[event]
        except:
            deletedEvents.append(oldEventDict[event])

    # Finding updated events
    for event in newEventDict.keys():
        try:
            if newEventDict[event] != oldEventDict[event]:
                updatedEvents.append(newEventDict[event])
        except:
            pass

    return {"created": newEvents,
            "deleted": deletedEvents,
            "updated": updatedEvents}
def handle_response(request_id, response, exception):
    if exception is not None:
        print(f"Error{request_id}{response}{exception}")
    else:
        print(f"Success\n{response}")
def generate_events(sheet):
    gevents = []
    emailMap = {
        "Cristiane Barion": "cbarion@beloteca.com",
        "Dorla Mirejovsky": "dmirejovsky@beloteca.com",
        "Fred Defesche": "fdefesche@beloteca.com",
        "Jenny Tran": "jtran@beloteca.com",
        "Katie Fisher": "kfisher@beloteca.com",
        "Lauren Ford": "lford@beloteca.com",
        "Ray Jenkins": "rjenkins@beloteca.com",
        "Ryan Luzum": "rluzum@beloteca.com"
    }
    for line in sheet:
        line = [cell.replace("\"", "") for cell in line.split("\",\"")]
        title = line[0] + " > " + line[1]
        id = ""
        startDate = datetime.datetime.strptime(line[2], "%x").strftime("%Y-%m-%d")
        endDate = datetime.datetime.strptime(line[3], "%x").strftime("%Y-%m-%d")
        summary = f"% Complete: {line[4]}\nTarget Start: {line[5]}\nTarget Finish: {line[6]}\nComments: {line[9]}"
        assigned = {"displayName": line[8],
                    "email": "belotecainc@gmail.com"}
        if line[8] in emailMap.keys():
            assigned['email'] = emailMap[line[8]]
        gevents.append({"summary": title,
                        "id": generate_id(title),
                        "description": summary,
                        "start": {
                            "date": startDate,
                            "timeZone": "America/Los_Angeles"
                        },
                        "end": {
                            "date": endDate,
                            "timeZone": "America/Los_Angeles"
                        },
                        "attendees": [
                            assigned
                        ]})
    return gevents

#initializing the smartsheet session
sm = smartsheet.Smartsheet(access_token="")
sm.Reports.get_report_as_csv(5463103753217924, "downloads/", "sheet.csv")

#modifying the smartsheet to be used
sheet = [line.replace("\n","") for line in open("sheet.csv", "r").readlines()][1:]
modList = [lineNo for lineNo in range(len(sheet)) if sheet[lineNo][0] != "\""][::-1]
for idx in modList: sheet[idx-1] += " "+sheet.pop(idx)

sheet1 = [line.replace("\n","") for line in open("sheet1.csv", "r").readlines()][1:]
modList1 = [lineNo for lineNo in range(len(sheet1)) if sheet1[lineNo][0] != "\""][::-1]
for idx in modList1: sheet1[idx-1] += " "+sheet1.pop(idx)

# initializing th egoogle calendar service to implement events and such
gc = GoogleCalendarService().service
batchUpdate = gc.new_batch_http_request(callback=handle_response)
CALENDAR_ID = "uofmsrvs68p8donubqkgdm0ht4@group.calendar.google.com"


# generating the events for the sheets
newEvents = generate_events(sheet)
oldEvents = generate_events(sheet1)

# Grabbing the differences and making them
eventChanges = grab_event_changes(newEvents, oldEvents)
for event in eventChanges['created']:
    batchUpdate.add(gc.events().insert(calendarId=CALENDAR_ID, body=event))
batchUpdate.execute()
batchUpdate = gc.new_batch_http_request(callback=handle_response)

for event in eventChanges['deleted']:
    batchUpdate.add(gc.events().delete(calendarId=CALENDAR_ID, body=event))
batchUpdate.execute()
batchUpdate = gc.new_batch_http_request(callback=handle_response)

for event in eventChanges['updated']:
    batchUpdate.add(gc.events().update(calendarId=CALENDAR_ID, eventId=event['id'], body=event))
batchUpdate.execute()
batchUpdate = gc.new_batch_http_request(callback=handle_response)

file1 = open("downloads/sheet.csv", "rb")
file2 = open("downloads/sheet1.csv", "wb")
file2.write(file1.read())