import speech_recognition as sr
import datetime
import win32api, win32con
import wmi
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import webbrowser
from battery import Batteries
import pyautogui
import time
import pyttsx3
import threading
import re
import queue

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
   IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

c = wmi.WMI()
t = wmi.WMI(moniker = "//./root/wmi")

r = sr.Recognizer()
engine = pyttsx3.init()
speech_queue = queue.Queue()

stop = False

def speak(text):
    speech_queue.put(text)

def tts_worker():
    while True:
        text = speech_queue.get()
        if text is None:
            break
        engine.say(text)
        engine.runAndWait()
        speech_queue.task_done()

threading.Thread(target=tts_worker, daemon=True).start()

def playpause():
    hwcode = win32api.MapVirtualKey(win32con.VK_MEDIA_PLAY_PAUSE, 0)
    win32api.keybd_event(win32con.VK_MEDIA_PLAY_PAUSE, hwcode)

def next_track():
    hwcode = win32api.MapVirtualKey(win32con.VK_MEDIA_NEXT_TRACK, 0)
    win32api.keybd_event(win32con.VK_MEDIA_NEXT_TRACK, hwcode)

def prev_track():
    hwcode = win32api.MapVirtualKey(win32con.VK_MEDIA_PREV_TRACK, 0)
    win32api.keybd_event(win32con.VK_MEDIA_PREV_TRACK, hwcode)

def decreaseVol(step=0.06):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    vol = cast(interface, POINTER(IAudioEndpointVolume))
    current = vol.GetMasterVolumeLevelScalar()
    new = max(current - step, 0.0)
    vol.SetMasterVolumeLevelScalar(new, None)

def increaseVol(step=0.06):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    vol = cast(interface, POINTER(IAudioEndpointVolume))
    current = vol.GetMasterVolumeLevelScalar()
    new = max(current + step, 0.0)
    vol.SetMasterVolumeLevelScalar(new, None)

def parseSwitchWindows():
    numbers = {
        "one": 1, "two": 2, "three": 3, "four": 4,
        "five": 5, "six": 6, "seven": 7, "eight": 8,
        "nine": 9, "ten": 10
    }

    match = re.search(r'\d+', text)
    if match:
        times = int(match.group())
    else:
        # check for spelled out number
        for word, num in numbers.items():
            if word in text:
                times = num
                break
        else:
            times = 1

    for _ in range(times):
        switchWindows()

def switchWindows():
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    time.sleep(0.1)
    pyautogui.keyUp('alt')
    pyautogui.keyUp('tab')

while not stop:
    try:
        with sr.Microphone() as source:
            print("listening...")
            audio = r.listen(source)
        text = r.recognize_google(audio).lower()

        print(f"I heard {text}")
        speak(f"I heard {text}")

        if "pause" in text or "play" in text:
            msg = "Ok, pausing"
            print(msg)
            speak(msg)
            playpause()
        elif "next" in text:
            next_track()
        elif "previous" in text:
            prev_track()
        elif "open" in text and "browser" in text:
            msg = "Ok, opening your browser"
            print(msg)
            speak(msg)
            webbrowser.open_new_tab("https://blank.org")
        elif "time" in text:
            msg = f"Ok, the time is {datetime.datetime.now()}"
            print(msg)
            speak(msg)
        elif "stop" in text:
            msg = "Ok, exiting"
            print(msg)
            speak(msg)
            stop = True
        elif "volume down" in text:
            decreaseVol()
        elif "volume up" in text:
            increaseVol()
        elif "test" in text:
            msg = "Test received. It's working."
            print(msg)
            speak(msg)
        elif "battery" in text:
            Batteries.getBattery()
        elif "start menu" in text:
            pyautogui.press('win')
        elif "switch windows" in text or "switch window" in text:
            parseSwitchWindows()
        else:
            msg = "I don't understand..."
            print(msg)
            speak(msg)
    except sr.exceptions.UnknownValueError:
        msg = "Sorry, I didn't catch that"
        print(msg)
        speak(msg)
