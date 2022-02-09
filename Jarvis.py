# Part 1: Importing modules & initialising variables

import speech_recognition as sr # Speech recognition is the process of converting spoken words to text.
import pyttsx3  # pyttsx3 is a text-to-speech conversion library in Python, which converts the entered text into speech.
import os   # The os module in Python provides functions for interacting with the operating system.
import winshell # The winshell module provides functionality for accessing folders, copying them, renaming & deleting.
import datetime # Python Datetime module supplies classes to work with date and time.
import webbrowser   # The webbrowser module provides a basic interface to the system's standard web browser.
import wikipedia    # Wikipedia is a Python library that makes it easy to access and parse data from Wikipedia.
import random   # Python Random module is an in-built module of Python which is used to generate random numbers
import pyjokes # Pyjokes is a python library that is used to create one-line jokes.

engine = pyttsx3.init('sapi5')  # We will use this function to get voices.
# Microsoft Speech API (SAPI5) is the technology for voice recognition and synthesis provided by Microsoft.
voices = engine.getProperty('voices')   # Initialising voices as the property of the engine.
engine.setProperty('voice', voices[1].id)   # Setting the voice property of the engine & selecting the voice for our assistant.
r = sr.Recognizer() # Declaring te recognizer function as 'r'.

# Part 2: Declaring Functions

def takeCommand():

    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.5)
        print("\nListening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print("User said: " + query + "\n")

    except Exception as e:
        print(e)
        speak("I am unable to recognize your voice sir. Kindly repeat.")
        print("Unable to recognize your voice.")
        return "None"

    return query


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak("Good Morning Sir !")

    elif 12 <= hour < 18:
        speak("Good Afternoon Sir !")

    else:
        speak("Good Evening Sir !")

    speak("I am your Voice Assistant.")


global uname


def username():
    speak("What should i call you sir?")
    uname = takeCommand()
    speak("Welcome Mister " + uname)
    print("Welcome Mr." + uname)

    speak("How can i Help you Sir?")
    return uname

# Part 3: Declaring the main function
if __name__ == '__main__':
    clear = lambda: os.system('cls')

    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    username()

    while True:

        query = takeCommand().lower()

        # All the commands said by user will be stored here in 'query' and will be converted to lower case for easily recognition of command
        if 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you sir")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")

        elif "open wikipedia" in query:
            speak('Opening Wikipedia for you sir')
            webbrowser.open("https://www.wikipedia.org/")

        elif 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("https://www.youtube.com/")

        elif 'google' in query:
            speak("Here you go to Google\n")
            programName = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            os.startfile(programName)

        elif 'chrome' in query:
            speak("Here you go to chrome\n")
            programName = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            os.startfile(programName)

        elif 'vs code' in query or 'visual studio' in query:
            speak("opening VS Code\n")
            programName = r"C:\Users\DELL\AppData\Local\Programs\Microsoft VS Code\Code.exe"
            os.startfile(programName)


        elif 'play music' in query or "play song" in query:
            speak("Here you go with music")
            music_dir = "C:\Piyush\Music"
            songs = os.listdir(music_dir)
            n = random.randint(0, 20)
            print("Playing: " + songs[n])
            random = os.startfile(os.path.join(music_dir, songs[n]))

        elif 'time' in query:
            time = datetime.datetime.now().strftime("%I:%M %p")
            print(time)
            speak("Sir, the time is" + time)

        elif "change your name" in query:
            speak("What would you like to call me sir? ")
            assname = takeCommand()
            speak("Thanks for naming me")

        elif "your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me", assname)

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Mohit, Aditya & Piyush.")

        elif "who are you" in query:
            speak("I am a virtual assistant created by Mohit, Aditya & Piyush")

        elif 'joke' in query:
            joke1 = pyjokes.get_joke()
            print('Joke: ' + joke1)
            speak(joke1)

        elif "who am i" in query:
            speak("You are my creator sir")

        elif 'powerpoint' in query:
            speak("opening MS PowerPoint")
            power = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
            os.startfile(power)

        elif 'excel' in query:
            speak("opening MS Excel")
            power = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
            os.startfile(power)

        elif 'word' in query:
            speak("opening MS Word")
            power = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
            os.startfile(power)

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/" + location + "")

        elif "write a note" in query:
            speak("What should i write sir")
            note = takeCommand()
            print(note)
            file = open('jarvis.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%I:%M %p")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
                speak("Alright sir")
            else:
                speak("Alright sir")
                file.write(note)

        elif "show note" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r")
            print(file.read())
            speak(file.read(6))

        elif 'exit' in query or 'bye' in query:
            speak("Thankyou for giving me your time sir.")
            exit()