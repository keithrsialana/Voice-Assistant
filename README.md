# Voice Assistant

## Set up
Before running the code, make sure to install the right libraries using pip. Run these commands in a terminal in order:
```
python -m venv venv
```
```
venv/Scripts/activate
```
```
pip install -r requirements.txt
```

## Running the Application
With the current version of the application, running the code is simply just like running any other Python application.
```
py ./main.py
```
or
```
python ./main.py
```

## About
This voice assistant is a personal passion project that mainly tests my capabilities to learn and understand external libraries to integrate onto a modern project that would be deemed quite interesting. I myself have always wanted to know how to create a fully automated assistant like Siri and Bigsby.

## Current Functionalities
These are Functionalities that have been completed:
- searching the web ("Search something up on Google for me")
- naming the assistant ("Change your name")
- naming the user ("Change my name")
- sleep mode ("Go to sleep")

## Known bugs
- depending on the mic, the assistant will hear background sounds. Can interfere with multi-input tasks like setting an alarm or creating an event on the calendar.
- creating an alarm doesn't call the right bash command, resulting in the alarm not being set.

## Future Goals
I can only work on creating this application when I have the free time to do so. That being said, these are the listed goals I have for this project:
- creating and removing Alarms using the Microsoft Clock application
- creating and removing Events using the Microsoft Calendar application
- simple calculations
- advanced calculations
- opening and closing other apps
