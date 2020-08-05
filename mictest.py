from __future__ import print_function
from wit import Wit
from geopy.geocoders import Nominatim
from datetime import date, timedelta, datetime, time
import speech_recognition as sr
import paths

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors

import pyaudio
import wave
import webbrowser
import http.client
import json
import geocoder
import pyttsx3
import time
import subprocess
import os
import signal


#Wit.ai settings
API_ENDPOINT = 'https://api.wit.ai/speech'
#Client access token
wit_access_token = "RE2XUQMFJZNENNLIJQG6RXYCUENJWF6E"

#read-only scope for google api
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

#Browser settings 
chrome_path = paths.chrome_path
webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))


r = sr.Recognizer()
source = sr.Microphone()
data = ""

#Get current location
location = geocoder.ip('me')

messages = []

service = ""
cal_service = ""

#Get Speech data and return intent from WIT
def RecogniseSpeech():

    #Parse speech
    r.adjust_for_ambient_noise(source)
    print("Listening")
    try:
        audio = r.listen(source, timeout=2)
        text = r.recognize_wit(audio, key=wit_access_token, show_all=True)
        actions(text)
        return None
    except sr.WaitTimeoutError:
        return print("No voice input recognised")
            

def get_weather(d, lat,lon):
    ##Get weather forecast from datapoint metoffice API
    conn = http.client.HTTPSConnection("api-metoffice.apiconnect.ibmcloud.com")
    headers = {'x-ibm-client-id': paths.client,
               'x-ibm-client-secret': paths.clientsecret,
               'accept':"application/json"}

    conn.request("GET", "/metoffice/production/v0/forecasts/point/daily?excludeParameterMetadata=true&includeLocationName=true&latitude={}&longitude={}".format(lat, lon),
                 headers=headers)

    res = conn.getresponse()
    data = res.read()
    weather_record_6_day = data.decode("utf-8")
    weather_record_6_day = json.loads(weather_record_6_day)
    date = d

    ##runs correctly and returns data within 6 days, add error checking for out of index range, add location name data
    choice_date = 0
    for i in weather_record_6_day['features'][0]['properties']['timeSeries']:
        if i['time'][0:10]==date:
            break
        choice_date+=1

    try:
        weather = weather_record_6_day['features'][0]['properties']['timeSeries'][choice_date]
    except IndexError:
        result = "Could not find weather information for that date" 
        return result

    #Based on codes returned from API call
    weather_types = {1:'Sunny',3:'Partly Cloudy', 5:'Misty',6:'Foggy',7:'Cloudy',8:'Overcast',
                     10:'Light Showers',11:'Drizzle',12:'Light Rain',14:'Heavy Showers',
                     15:'Heavy Rain',17:'Sleet Showers',18:'Sleet',20:'Hail Showers',
                     21:'Hail',23:'Light Snow Showers',24:'Light Snow',26:'Heavy Snow Showers',
                     27:'Heavy Snow', 29:'Thunder Shower',30:'Thunder'}

    
    day_weather = weather_types.get(weather['daySignificantWeatherCode'])
    day_highs = round(float(weather['dayMaxScreenTemperature']))
    day_lows = round(float(weather['nightMinScreenTemperature']))
    feels_like = round(float(weather['dayMaxFeelsLikeTemp']))

    
    result = "Weather will be {}, highs of {}, lows of {}, feels like {}".format(day_weather, day_highs, day_lows,
                                                                              feels_like)
    return result

#Text to speech 
def speech(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

#Offline voice recognition constant
def begin():
    r.listen_in_background(source, callback)
    time.sleep(1000000)
    sys.exit(0)

#Listens for keywords and activates Wit voice recognition or ends program
def callback(recognizer, audio):
    try:
        s_text = r.recognize_sphinx(audio, keyword_entries=[("lion",1),("end",1)])
        print(s_text)

        if "lion" in s_text:
            data = RecogniseSpeech()
        elif "end" in s_text:
            pid = os.getpid()
            os.kill(pid,signal.SIGTERM)

    except sr.UnknownValueError:
        print("Didn't catch that")

def set_mail_service():
    global service
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

    service = build('gmail', 'v1', credentials=creds)

def set_cal_service():
    global cal_service

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('caltoken.pickle'):
        with open('caltoken.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'calcredentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('caltoken.pickle', 'wb') as token:
            pickle.dump(creds, token)

    cal_service = build('calendar', 'v3', credentials=creds)
    
def email_check(data, service):
    global messages

    # Call the Gmail API
    try:
        response = service.users().messages().list(userId='me',
                                        q="is:unread").execute()

        if 'messages' in response:
            messages.extend(response['messages'])
        print("You have",len(messages),"new messages")
        msg_follow = True
    except errors.HttpError as e:
        print("An error occurred: {}".format(e))

def new_emails(msg_id, service):
    #Check if messages contains entries
    if msg_id:
        #Loop through message IDs
        for m in msg_id:
            message = service.users().messages().get(userId='me', id=m['id']).execute()
            sender = ""
            #Check for from header
            for header in message['payload']['headers']:
                if header['name']=='From':
                    sender = header['value']
            print ("From: {}".format(sender))
    else:
        print("No emails in list")

def date_event_check(service, date, end_date, week_range=False):
    #if single day
    if not week_range:
        print('Getting the events for the day')
        #Return dict of all events max 20
        events_result = service.events().list(calendarId='primary', timeMax = end_date, timeMin=date,
                                            maxResults=20, singleEvents=True,
                                            orderBy='startTime').execute()
        #Get items or return empty list
        events = events_result.get('items', [])
    

        if not events:
            speech('No events found for that day.')
        #Loop events and print time and summary per event
        for event in events:
            start = event['start'].get('dateTime')
            hour, minute = start[11:16].split(":")
            print(hour, minute)
            speech(hour+minutes(minute)+" "+ str(event['summary']))
            
    if week_range:
        print("Getting events for that range")
        #Get events for range, max 200
        events_result = service.events().list(calendarId='primary', timeMax = end_date, timeMin=date,
                                            maxResults=200, singleEvents=True,
                                            orderBy='startTime').execute()
        events = events_result.get('items', [])

        if not events:
            speech('No events found for that range.')
        #Loop events and print day, time and summary per event
        for event in events:
            start = event['start'].get('dateTime')
            hour, minute = start[11:16].split(":")
            day = start[8:10]+start[5:7]
            print(hour, minute)
            print(day)
            speech(days(day)+hour+minutes(minute)+" "+ str(event['summary']))

#Converting minutes to string to be used correctly in text to speech function
def minutes(date):
    words = ["hundred", "o one", "o two", "o three", "o four", "o five", "o six", "o seven",
             "o eight", "o nine"]
    if int(date)<10:
        return words[int(date)]

#Convert day and month section of date into strings for text to speech function
def days(date):
    final = ""
    endings = {"1":"st", "2":"nd", "3":"rd"}
    if date[1] in endings:
        final = date[1]+endings[date[1]]
    months = {"01":"January","02":"February","03":"March","04":"April","05":"May",
              "06":"June","07":"July","08":"August","09":"September","10":"October",
              "11":"November","12":"December"}
    final += " "+months[date[2:4]]+" "
    return final
    

def actions(data):
    if data:
        print(str(data))

        ##Open site in new tab
        if "website:website" in data['entities']:
            webbrowser.get('chrome').open(data['entities']['website:website'][0]['value'], new=0, autoraise=True)

        ##prototype for google search from voice
        elif "wit$search_query:search_query" in data['entities']:
            result = data['entities']['wit$search_query:search_query'][0]['value']
            webbrowser.get('chrome').open("http://www.google.com/#q="+result)

        ##if weather request, get date asked within 6 days, location if said. return weather for day/location
        elif "weather:weather" in data['entities']:
            if "wit$datetime:datetime" in data['entities']:
                if data['entities']['wit$datetime:datetime'][0]['type']=="value":
                    for entry in data['entities']['wit$datetime:datetime']:
                        chosen_date = entry['value'][:10]
                        if "wit$location:location" in data['entities']:
                            lat = data['entities']['wit$location:location'][0]['resolved']['values'][0]['coords']['lat']
                            lng = data['entities']['wit$location:location'][0]['resolved']['values'][0]['coords']['long']
                            text = ("In "+data['entities']['wit$location:location'][0]['resolved']['values'][0]['name']+" the ")
                            text += get_weather(chosen_date,lat,lng)
                            speech(text)
                        else:
                            location = geocoder.ip('me')
                            lat = location.lat
                            lng = location.lng
                            text = get_weather(chosen_date, lat, lng)
                            speech(text)
                elif data['entities']['wit$datetime:datetime'][0]['type']=="interval":
                    sat = datetime.strptime(data['entities']['wit$datetime:datetime'][0]['values'][0]['from']['value'][:10],'%Y-%m-%d').date()+timedelta(days=1)
                    sat = sat.strftime("%Y-%m-%d")
                    sun = datetime.strptime(data['entities']['wit$datetime:datetime'][0]['values'][0]['from']['value'][:10],'%Y-%m-%d').date()+timedelta(days=2)
                    sun = sun.strftime("%Y-%m-%d")
                    if "wit$location:location" in data['entities']:
                        lat = data['entities']['wit$location:location'][0]['resolved']['values'][0]['coords']['lat']
                        lng = data['entities']['wit$location:location'][0]['resolved']['values'][0]['coords']['long']
                        text = "On Saturday "
                        text += ("in "+data['entities']['wit$location:location'][0]['resolved']['values'][0]['name']+" the ")
                        text += get_weather(sat,lat,lng)
                        speech(text)
                        text = "On Sunday "
                        text += get_weather(sun,lat,lng)
                        speech(text)
                    else:
                        location = geocoder.ip('me')
                        lat = location.lat
                        lng = location.lng
                        text = "On Saturday the "
                        text += get_weather(sat,lat,lng)
                        speech(text)
                        text = "On Sunday the "
                        text += get_weather(sun,lat,lng)
                        speech(text)
            else:
                today = date.today().strftime("%Y-%m-%d")
                if "wit$location:location" in data['entities']:
                    lat = data['entities']['wit$location:location'][0]['resolved']['values'][0]['coords']['lat']
                    lng = data['entities']['wit$location:location'][0]['resolved']['values'][0]['coords']['long']
                    text = get_weather(today,lat,lng)
                    speech(text)
                else:
                    location = geocoder.ip('me')
                    lat = location.lat
                    lng = location.lng
                    text = get_weather(today,lat,lng)
                    speech(text)
        #Open specific programs, paths specified in another file, computer dependent
        elif "choice:choice" in data['entities']:
            choice = data['intents'][0]['name']
            if choice in paths.programs:
                subprocess.Popen(paths.programs[choice])
        elif "email:email" in data['entities']:
            #Checks if asking if new emails exist, sets up google API then runs query
            if data['intents'][0]['name'] == "new_check":
                set_mail_service()
                email_check(data, service)
            #Requires google API to already be ran from new_check then retrives from information for new messages
            elif data['intents'][0]['name'] == "new_check_from":
                new_emails(messages, service)
        #Checks for google calendar events for date/range specified
        elif "calendar:calendar" in data['entities']:
            if data['entities']['wit$datetime:datetime'][0]['type']=="value":
                rdate = data['entities']['wit$datetime:datetime'][0]['value']
                start_date = datetime.fromisoformat(rdate)
                end_date = start_date.replace(hour=23, minute=59, second=59)
                print(start_date)
                print(end_date)
                set_cal_service()
                date_event_check(cal_service, start_date.isoformat(), end_date.isoformat())
            elif data['entities']['wit$datetime:datetime'][0]['type']=="interval":
                start_date = data['entities']['wit$datetime:datetime'][0]['from']['value']
                end_date = data['entities']['wit$datetime:datetime'][0]['to']['value']
                print(start_date)
                print(end_date)
                set_cal_service()
                date_event_check(cal_service, start_date, end_date, week_range=True)


        else:
            print("No command recognised for that input")
    
if __name__ == "__main__":
    begin()
    print("Finished")

