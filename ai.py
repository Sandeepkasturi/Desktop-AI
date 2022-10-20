# coded by -- SANDEEP KASTURI --
# follow me for more..
import subprocess
from tkinter.ttk import Progressbar
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import pywhatkit
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir !")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir !")

    else:
        speak("Good Evening Sir !")

    assname = ("Rolex")
    phone = ('9550337558')
    speak("I am your Assistant")
    speak(assname)


def username():
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns

    print("#####################".center(columns))
    print("Welcome Mr.", uname.center(columns))
    print("#####################".center(columns))

    speak("How can i Help you, Sir")


def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to take your voice.")
        return "None"

    return query


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()

    # Enable low security in gmail
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()


if __name__ == '__main__':
    clear = lambda: os.system('cls')

    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    username()

    while True:

        query = takeCommand().lower()

        # All the commands said by user will be
        # stored here in 'query' and will be
        # converted to lower case for easily
        # recognition of command
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")

        elif 'play music' in query or "play song" in query:
            speak("Here you go with music")
            song = query.replace('tony play', '')
            speak('playing ' + song)
            pywhatkit.playonyt(song)
            

        elif 'what is the time' in query:
            strTime = datetime.datetime.now().strftime("%I:%M %p")
            speak(f"Sir, the time is {strTime}")

        elif 'open browser' in query:
            codePath = r"C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\brave.exe"
            os.startfile(codePath)

        elif 'email to sandeep' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "Receiver email address"
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("whome should i send")
                to = input()
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")

        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query

        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")

        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me", assname)

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Sandeep Kasturi.")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif "calculate " in query:
            pywhatkit.search(query)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = query.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)
        elif "who am i" in query:
            speak("If you talk then definitely your human.")

        elif "why you came to world" in query:
            speak("Thanks to Sandeep. further It's a secret")

        elif 'open cmd' in query:
            speak("opening command prompt")
            power = r"C:\Users\HP\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\System Tools"
            os.startfile(power)

        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")

        elif "who are you" in query:
            speak("I am your virtual assistant created by Sandeep Kasturi")

        elif 'reason for you' in query:
            speak("I was created as a Minor project by Sandeep kasturi ")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,
                                                       0,
                                                       "Location of wallpaper",
                                                       0)
            speak("Background changed successfully")

        elif 'open my c complier' in query:
            appli = r"C:\\TC\\BIN\\DB.EXE"
            os.startfile(appli)

        elif 'what is news today' in query:
            speak("Today news is ")
            query = query.replace('Hm ', '')
            speak('playing news ' + query)
            pywhatkit.search(query)
            speak(pywhatkit.search(query))
            
        elif 'search ' or 'what ' or 'who ' or 'how ' in query:
            person = query.replace('query ', '')
            info = wikipedia.summary(person, 1)
            print(info)
            speak(info)

        elif 'lock window' in query:
            speak('ok sir')
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak('ok sir')
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif 'empty recycle bin' in query:
            speak('ok sir')
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query:
            speak('ok sir')
            speak("for how much time you want to stop rolex from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")

        elif "camera" in query or "take a photo" in query:
            speak('ok sir')
            ec.capture(0, "rolex Camera ", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/s"])

        elif "hibernate" in query or "sleep" in query:
            speak('ok sir')
            speak("Hibernating") or speak('sleeping the system')
            subprocess.call("shutdown /s")

        elif "log off" in query or "sign out" in query:
            speak('ok sir')
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            speak('ok sir')
            speak("What should i write, sir")
            note = takeCommand()
            file = open('rolex.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in query:
            speak('ok sir')
            speak("Showing Notes")
            file = open("rolex.txt", "r")
            print(file.read())
            speak(file.read(6))

        elif "update assistant" in query:
            speak('ok sir')
            speak("After downloading file please replace this file with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream=True)

            with open("Voice.py", "wb") as Pypdf:

                total_length = int(r.headers.get('content-length'))

                for ch in Progressbar.bar(r.iter_content(chunk_size=2391975),
                                          expected_size=(total_length / 1024) + 1):
                    if ch:
                        Pypdf.write(ch)

        # NPPR9-FWDCX-D2C8J-H872K-2YT43
        elif "rolex" in query:

            wishMe()
            speak("rolex 2 point o in your service sir")

        elif "what is the weather" in query:

            # Google Open weather website
            # to get API of Open weather
            api_key = "Api key"
            base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
            speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()

            if x["code"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " + str(
                    current_temperature) + "\n atmospheric pressure (in hPa unit) =" + str(
                    current_pressure) + "\n humidity (in percentage) = " + str(
                    current_humidiy) + "\n description = " + str(weather_description))

            else:
                speak(" City Not Found ")

        elif "Good Morning" in query:
            speak("A warm" + query)
            speak("How are you Mister")
            speak(assname)

        # most asked question from google Assistant
        elif "will you be my gf" in query or "will you be my bf" in query:
            speak("I'm not sure about, may be you should give me some time")

        elif "how are you" in query:
            speak("I'm fine, glad you me that")

        elif "i love you" in query:
            speak("It's hard to understand")

        elif "what is " in query or "who is " in query:
            speak("Here you go")
            query = query.replace('showing query', '')
            speak('showing results ' + query)
            pywhatkit.search(query)
            
        # elif "" in query:
        # Command go here
        # For adding more commands
        # syntax from standard books like oxford python programming
        # coded by  -- SANDEEP KASTURI --
