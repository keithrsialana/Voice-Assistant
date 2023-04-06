import speech_recognition as sr

# get a list of available audio devices
print(sr.Microphone.list_microphone_names())

# use a specific microphone device by index
mic = sr.Microphone(device_index=2)

# adjust for ambient noise using the specified microphone device
with mic as source:
    r = sr.Recognizer()
    r.adjust_for_ambient_noise(source)

def recognize_speech():
    with mic as source:
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio)
            print("You said: " + query)
        except sr.UnknownValueError:
            print("I heard something.")
            query = ""
        except sr.RequestError:
            print("Sorry, I could not connect to the internet.")
            query = ""
    
# recognize speech from the audio source
while True:
    try:
        recognize_speech()
    except sr.UnknownValueError:
        print("Unable to recognize speech")

