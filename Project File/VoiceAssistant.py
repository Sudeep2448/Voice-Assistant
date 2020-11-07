import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import speech_recognition as sr
import pyttsx3
import pytz
import subprocess

engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty('rate', 185)

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november',
          'december']
DAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
DAYS_EXTENSIONS = ['rd', 'st', 'nd', 'th']


def speak(text):
    engine.say(text)
    engine.runAndWait()

def instructions():
    print("-------Please read the instructions to use this program----------\n")
    speak("wait till start command pop up in terminal")
    print("To execute any command ,first wake up the assistant by spelling wake up")
    print("     and processed your request , this instructions implies for every command you speak")
    print("To view your events that are registered on your google account , you can use any phrase")
    print("To take down the short note and save it to your project file ,"
          " spell #make a note# ")
    print("     And wait for further instruction\n")


def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.energy_threshold = 1000
        audio = r.listen(source)
        said = ''
    try:
        said = r.recognize_google(audio_data=audio)
        print(said)
    except:
        pass

    return said.lower()


def check_credentials():
    creds = None
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

    service = build('calendar', 'v3', credentials=creds)
    return service


def get_google_events(day, service):
    # Call the Calendar API
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)
    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), timeMax=end_date.isoformat(),
                                          singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No upcoming events found.')
    else:
        speak(f"you have {len(events)} events on this day")
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])
            start_time = str(start.split('T')[1].split(':')[0])
            if int(start_time.split(':')[0]) < 12:
                start_time = str(start.split('T')[1].split(":")[0:2]) + "AM"
            else:
                a = int(start_time)
                a = a - 12
                start_time = str(a) + str(start.split('T')[1].split(":")[1]) + "PM"
            speak(event['summary'] + " at " + start_time)


def get_date(text):
    text = text
    today = datetime.date.today()

    if text.count('today') > 0:
        return today

    day = -1
    month = -1
    year = today.year
    day_of_week = -1
    listk = text.split()
    for word in listk:
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        elif word in DAYS:
            day_of_week = DAYS.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for ext in DAYS_EXTENSIONS:
                found = word.find(ext)
                if found > 0:
                    try:
                        day = int(word[:found])
                    except:
                        pass

    if month < today.month and month == -1:
        year = year + 1

    if day < today.day and month == -1 and day != -1:
        month = month + 1

    if day_of_week != -1 and month == -1 and day == -1:
        current_day_of_week = today.weekday()  # returns whole number
        dif = day_of_week - current_day_of_week  # returns integer

        if dif < 0:
            dif += 7
            if text.count("next") >= 1:
                dif += 7

        return today + datetime.timedelta(dif)  # returns in the format of date object

    if day != -1:
        try :
            return datetime.date(day=day, month=month, year=year)
        except:
            speak("I think you mispelled something , try again")

def make_a_note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(':', '-') + "-note.txt"
    with open(file_name, 'w') as f:
        f.write(text)

dict_path ={'chrome' : "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" ,
            'edge'   : "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe" ,
            'camera' : "subprocess.Popen('start microsoft.windows.camera:', shell=True)",
            'calculator': "C:\\Windows\\System32\\calc.exe",
            'microsoft store' : "'start ms-windows-store:'",
            'word': "'start winword'",
            "system configuration utility" : "'start msconfig'",
            'notepad' : "'start notepad'",
            'excel' : "'start excel'",
            'outlook' : "'start outlook'",
            'powerpoint' : "'start powerpnt'"}

instructions()
open_application = "open application"
colon='\\'
wake = "wake up"
note_making_str = ['make a note']
speak("Wait till google credentials get checked!!")
service = check_credentials()
# get_google_events(get_date("9th november"),service)

while True:
    print("start")
    print("listening")
    text = get_audio()

    if text.count(wake) > 0:
        speak("I am ready")
        text = get_audio()
        date = get_date(text)
        if date != None:
            get_google_events(date, service)

    for phrase in note_making_str:
        if phrase in text:
            speak("What would you like me to write down?")
            print("Speak Now\n")
            note_text = get_audio()
            make_a_note(note_text)
            speak("I have made a note of that")

    if open_application in text:
        speak("Which application do you want to open:-\n")
        application_name=get_audio()
        get_value = dict_path.get(application_name)
        if application_name in dict_path:
            if colon in get_value:
                try:
                    subprocess.Popen(get_value)
                except:
                    speak("Try again , some error has occured")
            else:
                try:
                    subprocess.Popen(eval(get_value),shell=True)
                except:
                    speak("Try again , some error has occured")
        else:
            speak("Application that you are trying to open doesn't exist in this program")

