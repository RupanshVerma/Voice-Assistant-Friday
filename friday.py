import pyttsx3
from pyttsx3.drivers import sapi5
import wikipedia
import speech_recognition as sr
import datetime
import webbrowser
import os
import smtplib
import pyjokes
import wolframalpha
import subprocess
import json 
import ctypes
import operator
import winshell
import time 
import requests
import shutil
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
engine.setProperty('voices',voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good morning!")
    elif hour>=12 and hour<16:
        speak("Good Afternoon!")
    elif hour>=16 and hour<20:
        speak("Good Evening!")
    else :
        speak("Good Night!")
    speak("Hello. I am Friday . An Artificial Intelligence voice assistant. How may i help you sir?")

def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening.....")
        r.pause_threshold = 1
        audio=r.listen(source)
    try:
        print("Recognising......")
        query=r.recognize_google(audio,language='en-in')
        print(f"User said : {query}\n")   
    
    except Exception as e:
        print("Sorry Sir. I did not get what you said. Can you repeat please....")
        return "None"
    return query

def sendEmail(to,content):
    server=smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('2019648435.rupansh@ug.sharda.ac.in','Assassin@30')
    server.sendmail('2019648435.rupansh@ug.sharda.ac.in',to,content)
    server.close()

if __name__=="__main__":
    clear=lambda: os.system('cls')
    wishMe()
    while True:
        query=takeCommand().lower()

        if 'wikipedia' in query:
            speak('Searching wikipedia...')
            query=query.replace("wikipedia","")
            results=wikipedia.summary(query,sentences=3)
            speak("According to wikipedia")
            print(results)
            speak(results)
        elif 'open youtube' in query:
            webbrowser.open("youtube.com")
        elif 'open google' in query:
            webbrowser.open("google.com")    
        elif 'open facebook' in query:
            webbrowser.open("facebook.com")
        elif 'open fb' in query:
            webbrowser.open("facebook.com")
        elif 'open twitter' in query:
            webbrowser.open("twitter.com")
        elif 'play music' in query:
            music_dir='F:\\Me\\music'
            songs=os.listdir(music_dir)
            os.startfile(os.path.join(music_dir,songs[0]))
        elif 'play song' in query:
            music_dir='F:\\Me\\music'
            songs=os.listdir(music_dir)
            os.startfile(os.path.join(music_dir,songs[0]))
        elif 'the time' in query:
            strTime=datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, The time right now is {strTime}")
        elif 'open visual studio' in query:
            codePath="C:\\Users\\Assassin\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
        elif 'email to rajat' in query:
            try:
                speak("What should i say ?")
                content=takeCommand()
                to="2019502614.rajat@ug.sharda.ac.in"
                sendEmail(to,content)
                speak("Email has been sent")
            except Exception as e:
                speak("Sorry sir, the email was not sent..")
        elif 'how are you' in query: 
            speak("I am fine, Thank you") 
            speak("How are you, Sir")
        elif 'fine' in query or "good" in query: 
            speak("It's good to know that your fine") 
        elif "what's your name" in query or "What is your name" in query: 
            speak("My friends call me Friday") 
        elif "who made you" in query or "who created you" in query:  
            speak("I have been created by Rupansh and Rajat") 
        elif "calculate" in query:  
            app_id = '69K6WV-RP9U6736RL'
            client = wolframalpha.Client(app_id) 
            indx = query.lower().split().index('calculate')  
            query = query.split()[indx + 1:]  
            res = client.query(' '.join(query))  
            answer = next(res.results).text 
            print("The answer is " + answer)  
            speak("The answer is " + answer)  
        elif 'What is love' in query: 
            speak("It is 7th sense that destroy all other senses") 
        elif 'lock window' in query: 
            speak("locking the device") 
            ctypes.windll.user32.LockWorkStation() 
        elif 'shutdown system' in query: 
            speak("Hold On a Sec ! Your system is on its way to shut down")  
            os.system("shutdown /s /t 1")
        elif 'empty recycle bin' in query: 
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True) 
            speak("Recycle Bin Recycled") 
        elif "where is" in query: 
            query = query.replace("where is", "") 
            location = query 
            speak("User asked to Locate") 
            speak(location) 
            webbrowser.open("https://www.google.nl/maps/place/"+location) 
        elif "restart computer" in query: 
            subprocess.call(["shutdown", "/r"]) 
        elif "restart the computer" in query: 
            subprocess.call(["shutdown", "/r"]) 
        elif "restart system" in query: 
            subprocess.call(["shutdown", "/r"]) 
        elif "weather" in query:  
            api_key = "37e5faaec67d6e67a57cacb29c52b096" 
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak(" City name ") 
            print("City name : ") 
            city_name = takeCommand() 
            complete_url = base_url + "q="+ city_name + "&appid=" + api_key 
            response = requests.get(complete_url)  
            x = response.json()  
            if x["cod"] != "404":  
                y = x["main"]  
                current_temperature = y["temp"]  
                current_pressure = y["pressure"]  
                current_humidiy = y["humidity"]  
                z = x["weather"]  
                weather_description = z[0]["description"]  
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))  
              
            else:  
                speak(" City Not Found ")  
        elif "write a note" in query: 
            speak("What should i write, sir") 
            note = takeCommand() 
            file = open('friday.txt', 'w') 
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
            speak("Showing Notes") 
            file = open("friday.txt", "r")  
            print(file.read()) 
            speak(file.read(6)) 
        elif "how are you" in query: 
            speak("I'm fine, glad you asked me that") 
        elif "i love you" in query: 
            speak("It's hard to understand")
        elif "what is" in query or "who is" in query: 
            client = wolframalpha.Client("69K6WV-RP9U6736RL") 
            res = client.query(query) 
            try: 
                print (next(res.results).text) 
                speak (next(res.results).text) 
            except StopIteration: 
                print ("No results")
        elif 'tell me a joke' in query:
            joke=pyjokes.get_joke()
            print(joke)
            speak(joke)
        elif 'tell me another joke' in query:
            joke=pyjokes.get_joke()
            print(joke)
            speak(joke)
        elif 'friday shutdown' in query:
            exit()
        elif 'friday quit' in query:
            exit()
        elif 'quit friday' in query:
            exit()
        elif 'shutdown friday' in query:
            exit()
        elif 'Bye friday' in query:
            exit()
        elif 'friday stop' in query:
            exit()
input("press enter to close")
