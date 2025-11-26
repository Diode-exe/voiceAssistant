import speech_recognition as sr
import datetime
import win32api, win32con
import os

# Source - https://stackoverflow.com/a/10441322
# Posted by Skurmedel, modified by community. See post 'Timeline' for change history
# Retrieved 2025-11-26, License - CC BY-SA 4.0
def playpause():
    hwcode = win32api.MapVirtualKey(win32con.VK_MEDIA_PLAY_PAUSE, 0)
    win32api.keybd_event(win32con.VK_MEDIA_PLAY_PAUSE, hwcode)

stop = False

while not stop:
    r = sr.Recognizer()

    with sr.Microphone() as source:
        print("listening...")
        audio = r.listen(source)
    text = r.recognize_google(audio).lower()

    if "pause" or "play" in text:
        playpause()
    elif "open" in text and "chrome" in text:
        os.system("start chrome")
    elif "time" in text:
        print(datetime.datetime.now())
    elif "stop" in text:
        stop = True