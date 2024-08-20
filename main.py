import pyttsx3
import requests
import speech_recognition as sr
import keyboard
import os
import subprocess as sp
from datetime import datetime
from decouple import config
from random import choice
from conv import random_text
from online import find_my_ip, search_on_google, search_on_wikipedia, youtube, send_email, get_news, weather_forecast

engine = pyttsx3.init('sapi5')
engine.setProperty('volume',1.5)
engine.setProperty('rate',225)
voices = engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)

USER = config('USER')
HOSTNAME = config('BOT')


def speak(text):
    engine.say(text)
    engine.runAndWait()


def greet_me():
    hour = datetime.now().hour
    if(hour>=6) and (hour < 12):
        speak(f"Good morning {USER}")
    elif (hour >= 12) and (hour <=16 ):
        speak(f"Good afternoon {USER}")
    elif (hour >= 16) and (hour < 19):
        speak(f"Good evening {USER}")
    speak(f"Hi I am {HOSTNAME}. How may I assist you? {USER}")


listening =  False


def start_listening():
    global listening
    listening = True
    print("started listening ")


def pause_listening():
    global listening
    listening = False
    print("stopped listening")


keyboard.add_hotkey('ctrl+alt+k',start_listening)
keyboard.add_hotkey('ctrl+alt+p',pause_listening)


def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing....")
        queri = r.recognize_google(audio,language='en-in')
        print(queri)
        if not 'stop' in queri or 'exit' in queri:
            speak(choice(random_text))
        else:
            hour  =datetime.now().hour
            if hour >= 21 and hour < 6:
                speak("Good night sir, take care!")
            else:
                speak("Have a good day sir!")
            exit()

    except Exception:
        speak("Sorry I couldn't understand. Can you please repeat that?")
        queri = 'None'
    return queri


if __name__ == '__main__':
    greet_me()
    while True:
        if listening:
            query = take_command().lower()
            if "how are you" in query:
                speak("I am absolutely fine sir. What about you")
            elif "open command prompt" in query:
                speak("Opening command prompt")
                os.system('start cmd')


            elif "open camera" in query:
                speak("Opening camera sir")
                sp.run('start microsoft.windows.camera:',shell=True)

            elif "open notepad" in query:
                speak("Opening Notepad for you sir")
                notepad_path = "C:\Windows\System32\notepad.exe"
                os.startfile(notepad_path)

            elif "open discord" in query:
                speak("Opening Discord for you sir")
                discord_path = "C:\\Users\\harsh\\AppData\\Local\\Discord\\app-1.0.9157"
                os.startfile(discord_path)

            elif "open excel" in query:
                speak("Opening Excel for you sir")
                excel_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
                os.startfile(excel_path)

            elif "open powerpoint" in query:
                speak("Opening Powerpoint for you sir")
                powerpoint_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
                os.startfile(powerpoint_path)

            elif "open onenote" in query:
                speak("Opening OneNote for you sir")
                onenote_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\ONENOTE.EXE"
                os.startfile(onenote_path)

            elif "open outlook" in query:
                speak("Opening Outlook for you sir")
                outlook_path = "C:\\Program Files\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"
                os.startfile(outlook_path)

            elif "open openshot video editor" in query:
                speak("Opening Outlook for you sir")
                openshotvideoeditor_path = "C:\\Program Files\\OpenShot Video Editor\\openshot-qt.exe"
                os.startfile(openshotvideoeditor_path)

            elif "open word" in query:
                speak("Opening Word for you sir")
                word_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
                os.startfile(word_path)

            elif "open microsoft edge" in query:
                speak("Opening Microsoft Edge for you sir")
                microsoftedge_path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
                os.startfile(microsoftedge_path)

            elif "open vmware horizon view client" in query:
                speak("Opening VMware Horizon View Client for you sir")
                vmwarehorizonclient_path = "C:\\Program Files\\VMware\\VMware Horizon View Client\\vmware-view.exe"
                os.startfile(vmwarehorizonclient_path)

            elif "open google chrome" in query:
                speak("Opening Google Chrome for you sir")
                googlechrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                os.startfile(googlechrome_path)

            elif "ip address" in query:
                ip_address = find_my_ip()
                speak(
                    f"your ip address is {ip_address}"
                )
                print(f"your ip address is {ip_address}")

            elif "open youtube" in query:
                speak("What do you want to play on youtube sir?")
                video = take_command().lower()
                youtube(video)

            elif "open google" in query:
                speak("What do you want to search on google?")
                query = take_command().lower()
                search_on_google(query)

            elif "wikipedia" in query:
                speak("What do you want to search on wikipedia sir?")
                search = take_command().lower()
                results = search_on_wikipedia(search)
                speak(f"According to wikipedia,{results}")
                speak("I am printing in on terminal")
                print(results)


            elif "send an email" in query:
                speak("On what email address do you want to send sir?. Please enter in the terminal")
                receiver_add = input("Email address:")
                speak("What should be the subject sir?")
                subject = take_command().capitalize()
                speak("What is the message ?")
                message = take_command().capitalize()
                if send_email(receiver_add, subject, message):
                    speak("I have sent the email sir")
                    print("I have sent the email sir")

                else:
                    speak("something went wrong Please check the error log")

            elif "give me news" in query:
                speak(f"I am reading out the latest headlines of today,sir")
                speak(get_news())
                speak("I am printing it on screen sir")
                print(*get_news(),sep='\n')


            elif 'weather' in query:
                ip_address = ip_address()
                city = requests.get(f"https://ipapi.co/{ip_address}/city/").text
                speak(f"Getting weather report for your city {city}")
                weather, temp, feels_like = weather_forecast(city)
                speak(f"The current temperature is {temp}, but it feels like {feels_like}")
                speak(f"Also, the weather report talks about {weather}")
                speak("For your convenience, I am printing it on the screen sir.")
                print(f"Description: {weather}\nTemperature: {temp}\nFeels like: {feels_like}")



