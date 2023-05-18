import speech_recognition as sr
import pyttsx3
import wikipedia
import datetime
import os

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voices',voices[1].id)


def speak(audio):
    print(" ")
    print(f": {audio}")
    engine.say(audio)
    engine.runAndWait()
    print(" ")
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en')
        print(f"User said: {query}\n")
    except:
        print("Sorry, I didn't hear that. Please say that again.")
        return ""
    return query.lower()


def ai_tasks():

    while True:

        command = takeCommand()
        if 'what is' in command.lower():
            search_query = command.lower().replace('what is', '')
            wikipedia_summary = wikipedia.summary(search_query, sentences=1)
            print(wikipedia_summary)
            speak(wikipedia_summary)

        elif 'what time is it' in command.lower():
            current_time = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"The current time is {current_time}")
            speak(f"The current time is {current_time}")

        elif 'power point presentation' or 'power point' in command.lower():
            speak("opening Power Point presentation")
            power = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
            os.startfile(power)

        elif 'excel' or 'microsoft excel' in command.lower():
            speak("opening Microsoft Excel")
            excel = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
            os.startfile(excel)

        elif 'onenote' in command:
            speak("opening onenote")
            note = r"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE"
            os.startfile(note)

        elif 'word' or 'microsoft word' in command:
            speak("opening Microsoft word")
            word = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
            os.startfile(word)
        elif 'stop' in command.lower() or 'exit' in command.lower() or 'quit' in command.lower():
            print("Goodbye!")
            speak("Goodbye!")
            exit()
        else:
            speak("Sorry, I don't know how to do that yet.")