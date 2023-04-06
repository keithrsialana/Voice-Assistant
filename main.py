import speech_recognition as sr
import pyttsx3
import os
import subprocess
import webbrowser
import datetime
import win32com.client as wincl
import sys
import re
from datetime import datetime as dt

# Initialize text-to-speech engine
engine = pyttsx3.init()

# Initialize speech recognition engine
r = sr.Recognizer()

# use a specific microphone device by index
mic = sr.Microphone(device_index=2)

# adjust for ambient noise using the specified microphone device
with mic as source:
    r.adjust_for_ambient_noise(source)

# Set the default audio source
# with sr.Microphone() as source:
#     r.adjust_for_ambient_noise(source)

# Define a function to speak the response
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Define a function to recognize speech
def recognize_speech():
    with mic as source:
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio)
            print("You said: " + query)
            if query == "cancel" or query == "never mind" or query == "nevermind":
                query == ""
        except sr.UnknownValueError:
            query = ""
        except sr.RequestError:
            print("Sorry, I could not connect to the internet.")
            query = ""
    return query

# Define a function to set an alarm TODO
def set_alarm():
    alarm_name = ""
    alarm_time = ""
    snoozeTime = ""
    cmd = f'ms-switchalarms.exe create "{alarm_name}" /time "{alarm_time}" /snooze {snoozeTime} /sound "Alarm"'

    speak("What time do you want to set the alarm for?")
    alarm_time = recognize_speech()

    # Replaces to proper AM/PM format
    if "a.m." in alarm_time:
        alarm_time = alarm_time.replace("a.m.","AM")
    elif "p.m." in alarm_time:
        alarm_time = alarm_time.replace("p.m.", "PM")
    
    # Check if the time is in the correct format
    if not re.match(r"\d{1,2}:\d{2} (AM|PM)", alarm_time):
        speak("Sorry, that time format is not valid. Please try again.")
        return
    
    # Converts 12-hr time to 24-hr time
    # time_object = dt.strptime(alarm_time, "%I:%M %p")
    # alarm_time = dt.strftime(time_object, "%H:%M")
    
    # Set Alarm Name
    speak("What would you like to call the alarm?")
    alarm_name = recognize_speech()

    # Set snooze time
    snoozeTime = ""
    speak("Set the snooze time?")
    if recognize_speech() == "yes":
        speak("How many minutes?")
        snoozeTime = recognize_speech()
        if "minute" in snoozeTime:
            snoozeTime = snoozeTime.replace("minute", "")
        if "miuntes" in snoozeTime:
            snoozeTime = snoozeTime.replace("minutes", "")
        snoozeTime.strip()
    else:
        snoozeTime = "5"

    # Set repeat
    speak("Would you like it to repeat?")
    if recognize_speech() == "yes":
        speak("Which days would you like it to repeat on?")
        days = recognize_speech()
        foundDays = ""
        if "monday" in days:
            foundDays = foundDays + "M"
        if "tuesday" in days:
            foundDays = foundDays + ",T"
        if "wednesday" in days:
            foundDays = foundDays + ",W"
        if "thursday" in days:
            foundDays = foundDays + ",Th"
        if "friday" in days:
            foundDays = foundDays + ",F"
        if "saturday" in days:
            foundDays = foundDays + ",S"
        if "sunday" in days:
            foundDays = foundDays + ",Su"
        cmd = cmd + f' /repeat "{foundDays}"'
        
    # Execute the command to set the alarm
    # cmd = f'ms-switchalarms.exe create "Alarm Name" /time "8:00 AM" /snooze 5 /sound "Alarm" /repeat "M,T,W,Th,F,S,Su"'
    print("Executing command:", cmd)
    subprocess.run(cmd, shell=True)

    # subprocess.Popen(['cmd.exe', '/c', 'start', 'ms-clock:', f'/alarm {alarm_time}'])
    return

# Define a function to add an event to the calendar TODO
def add_event():
    speak("What is the name of the event?")
    event_name = recognize_speech()
    speak("When is the event? Please say the date and time.")
    event_time = recognize_speech()
    subprocess.Popen(['outlook.exe', '/c', 'IPM.Appointment', f'/subject "{event_name}"', f'/start "{event_time}"'])

# Define a function to remove an event from the calendar TODO
def remove_event():
    speak("What is the name of the event you want to remove?")
    event_name = recognize_speech()
    outlook = wincl.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')
    calendar = namespace.GetDefaultFolder(9)
    for item in calendar.Items:
        if item.Subject == event_name:
            item.Delete()
            speak(f"{event_name} has been removed from the calendar.")

# Define a function to open an application TODO
def open_application():
    speak("Which application do you want to open?")
    app_name = recognize_speech().lower()

# Define a function to close an application TODO
def close_application():
    speak("Which application do you want to close?")
    app_name = recognize_speech().lower()
    os.system(f"taskkill /f /im {app_name}.exe")

# Define a function to search the web TODO
def search_web():
    speak("What do you want to search for?")
    query = recognize_speech().lower()
    url = f"https://www.google.com/search?q={query}"
    webbrowser.get().open(url)
    speak(f"Here are the search results for {query}.")

# Define a function to set the user's name
def set_user_name():
    speak("What is your name?")
    name = recognize_speech().lower().replace("my name is", "").replace("my name's'", "").replace("i'm", "").strip()
    with open("user_name.txt", "w") as f:
        f.write(name)
    speak(f"Okay, from now on, I will call you {name}")

# Define a function to get the user's name
def get_user_name():
    try:
        with open("user_name.txt", "r") as f:
            name = f.read()
            return name     
    except FileNotFoundError:
        set_user_name()
        name = get_user_name()
        return name

# Define a function to set the name of the voice assistant
def set_name():
    speak("What would you like to name me?")
    name = recognize_speech().lower()
    with open("name.txt", "w") as f:
        f.write(name)
    speak(f"From now on, you can call me {name}.")

# Define a function to get the name of the voice assistant
def get_name():
    try:
        with open("name.txt", "r") as f:
            name = f.read()
            return name     
    except FileNotFoundError:
        set_name()
    
# Define a function to check if the voice assistant's name was called or if the user asked for the name
def is_called(query):
    name = get_name()
    if name in query:
        return True
    elif "assistant change your name" in query:
        set_name()
    elif "assistant what's your name" in query:
        speak(f"My name is {name}.")
        return False
    else:
        return False

# Define a function to start listening for commands only when the name is called
def start_listening():
    userName = get_user_name()
    speak(f"Hello {userName}, How can I help you?")
    listening = True
    while True:
        if listening:
            query = recognize_speech().lower()
            if is_called(query):
                query = query.replace(get_name(), "").strip()
                if "set an alarm" in query or "set alarm" in query or "add an alarm" in query or "add alarm" in query:
                    set_alarm()
                elif "add event" in query or "add an event" in query:
                    add_event()
                elif "remove event" in query or "remove an event" in query:
                    remove_event()
                elif "open application" in query or "open an application" in query:
                    open_application()
                elif "close application" in query or "close an application " in query:
                    close_application()
                elif "search web" in query or "google something for me" in query or "look something up" in query:
                    search_web()
                elif "change your name" in query:
                    set_name()
                elif "what's my name" in query:
                    foundUserName = get_user_name()
                    speak(f"You are {foundUserName}")
                elif "change my name" in query:
                    set_user_name()
                elif "hi" in query or "are you here" in query or "hello" in query or "are you still here" in query or "can you hear me" in query or "are you there" in query:
                    speak("Yes, I'm still here")
                elif "you can sleep" in query or "sleep" in query:
                    speak("Okay, let me know if you need anything.")
                    listening = False
                elif "close yourself" in query or "exit" in query or "turn off" in query or "power down" in query or "shut down" in query or "shutdown" in query or "shut off" in query:
                    speak("Goodbye!")
                    sys.exit()
                else:
                    speak("Sorry, I did not understand that. Please try again.")
            else:
                print("I heard something, but my name was not called")
        else:
            if is_called(recognize_speech().lower()):
                speak(f"Hello {userName}! I'm back and ready to listen to your commands.")
                listening = True

# Define a main function to handle user input
def main():
    start_listening()

if __name__ == "__main__":
    main()