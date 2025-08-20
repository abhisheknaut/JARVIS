import wikipedia
import webbrowser
import os
import win32com.client 
import time
import speech_recognition as sr
import pyaudio

#  speaker set-up
engine= win32com.client.Dispatch('SAPI.SpVoice')


def wishMe():
    samay = time.strftime("%H")
    if(samay>="00" and samay<="12"):
        engine.speak("Good morning sir")

    elif(samay>"12" and samay<="16"):
        engine.speak("Good afternoon sir")
    elif(samay>"16" and samay<="18"):
        engine.speak("good evening sir")
    else:
        engine.speak("goood night sir")

    engine.speak("i am jarvis sir please tell me how can i help you")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening")
        r.pause_threshold =1
        audio =r.listen(source)
    try:
        print("Recogniging")
        query = r.recognize_google(audio)
        print(query)
        # print(f"user said : {query}\n")
    except Exception as e:
        print("say that again please")
        return "None"
    return query

if __name__ == "__main__":
    wishMe()
    
    while True:
     query  = takeCommand().lower()
    #  query = input("enter : ")

     if 'wikipedia' in query:
         engine.speak("searching wikipedia")
         query = query.replace("wikipedia ","")
         result = wikipedia.summary(query,sentences=2)
         print(result)
         engine.speak(f"According to wikipedia {result}")

     elif 'open youtube ' in query:
        #  os.startfile("C:\\Users\\MODERN 14\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Chrome Apps\\YouTube.lnk")
         webbrowser.open("youtube.com")

     elif 'open google' in query:
         webbrowser.open("google.com")

     elif 'open instagram' in query:
         webbrowser.open("instagram.com")

     elif 'open whatsapp' in query:
         webbrowser.open("web.whatsapp.com")

     elif 'play music' in query:
       os.startfile("C:\\Users\\MODERN 14\\Music\\Playlists\\New folder\\abhishek.audio.unknown")
        
     elif 'show time' in query:
         strtime  = time.strftime("%H:%M:%S")
         print(strtime)
         engine.speak(strtime)

     elif 'open vscode' in query:
         codePath = "C:\\Users\\MODERN 14\\AppData\\Local\\Programs\\Python\\Python313\\python.exe" "c:/Users/MODERN 14/.vscode/python"
         os.startfile(codePath)
     elif 'break program' in query:
         break
     else:
         engine.speak("sorry sir i am not able to understand your command")

