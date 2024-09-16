from AppOpener import open, close
import pyttsx3
import speech_recognition as sr
import os
import time
import webbrowser

# Using time and pyttsx3 module for greetings------>
engine = pyttsx3.init()
current_hour = int(time.strftime('%H'))
if current_hour < 12:
    print("Good morning sir, I am Jarvis, how can i help you?")
    engine.say("Good morning sir, I am Jarvis, how can i help you?")
    engine.runAndWait()
elif 12 <= current_hour < 16:
    print("Good after noon sir, I am Jarvis, how can i help you?")
    engine.say("Good after noon sir, I am Jarvis, how can i help you?")
    engine.runAndWait()
elif current_hour >= 16:
    print("Good evening sir, I am Jarvis, how can i help you?")
    engine.say("Good evening sir, I am Jarvis, how can i help you?")
    engine.runAndWait()

# Storing recognizer function in recognizer variable------>
recognizer = sr.Recognizer()


# Function for listening command------->
def listen_for_command():
    with sr.Microphone() as source:  # Using microphone named as source
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)  # Adjusts the ambient noise
        audio = recognizer.listen(source)  # Listens the voice
    try:
        print("Recognizing...")
        cmd = recognizer.recognize_google(audio)  # Recognize the voice
        print("User:", cmd)
        return cmd.lower()
    except sr.UnknownValueError:
        print("Sorry I didn't catch that")
        engine.say("Sorry I didn't catch that")
        engine.runAndWait()
        return ""
    except sr.RequestError:
        print("Could not request results. Please check your internet connection")
        engine.say("Could not request results. Please check your internet connection")
        engine.runAndWait()
        return ""


# Function for executing command------>
def execute_command(cmd):
    if "open google" in cmd:                       # Change the name if you use some other browser
        print("Opening browser")
        engine.say("Opening google")
        engine.runAndWait()
        webbrowser.open("https://www.google.com")  # Opens google, Set the URL as per your browser you use

    elif "close browser" in cmd:
        print("closing browser")
        engine.say("closing browser")
        engine.runAndWait()
        os.system("taskkill /f /im chrome.exe")  # Closes google, use the the browser name if you use instead of chrome

    elif "open youtube" in cmd:
        print("opening youtube")
        engine.say("opening youtube")
        engine.runAndWait()
        webbrowser.open("https://www.youtube.com")  # Opens youtube
    elif "open whatsapp" in cmd:
        print("opening whatsapp")
        engine.say("opening whatsapp")
        engine.runAndWait()
        open("Whatsapp")  # opens whatsapp
    elif "close whatsapp" in cmd:
        print("closing whatsapp")
        engine.say("closing whatsapp")
        engine.runAndWait()
        close("Whatsapp")  # Closes whatsapp
    elif "open notepad" in cmd:
        print("opening notepad")
        engine.say("opening notepad")
        engine.runAndWait()
        open("notepad")  # opens notepad
    elif "close notepad" in cmd:
        print("closing notepad")
        engine.say("closing notepad")
        engine.runAndWait()
        close("notepad")  # closes notepad
    elif "shutdown the computer" in cmd:
        print("shutting down")
        engine.say("shutting down")
        engine.runAndWait()
        os.system("shutdown /s /t 1")  # Shuts down the pc
    elif "meet my family" in cmd:
        print("Hello dear siblings! nice to meet you all. I am Jarvis")
        engine.say("Hello dear siblings! nice to meet you all. I am Jarvis")  # Greets family
        engine.runAndWait()
    elif "meet my friends" in cmd:
        print("Hello buddies! nice to meet you all. I am Jarvis")
        engine.say("Hello buddies! nice to meet you all. I am Jarvis")  # Greets friends
        engine.runAndWait()
    elif "what is current time" in cmd:
        current_time = time.localtime(time.time())
        l_time = time.asctime(current_time)
        say_time = "Current time is " + l_time
        print(say_time)
        engine.say(say_time)
        engine.runAndWait()
    elif "search for" in cmd:
        print("Sure sir!")
        engine.say("Sure sir!")
        engine.runAndWait()
        search_query = cmd.split("search for ")[1].strip()  # Remove leading and trailing whitespace
        search_query = search_query.replace(" ", "+")  # Replace spaces with '+'
        search_url = "https://www.google.com/search?q=" + search_query
        webbrowser.open(search_url)

    else:
        print("Command not recognized")
        engine.say("Command not recognized")
        engine.runAndWait()


# Loop to keep program executing------>
while True:
    command = listen_for_command()
    if command:
        execute_command(command)
