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

# Source - https://stackoverflow.com/a/10441322
# Posted by Skurmedel, modified by community. See post 'Timeline' for change history
# Retrieved 2025-11-26, License - CC BY-SA 4.0

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
   IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

c = wmi.WMI()
t = wmi.WMI(moniker = "//./root/wmi")

r = sr.Recognizer()

stop = False


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

def switchWindows():
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    time.sleep(0.1)
    pyautogui.keyUp('alt')

while not stop:
    try:
        with sr.Microphone() as source:
            print("listening...")
            audio = r.listen(source)
        text = r.recognize_google(audio).lower()

        print(f"I heard {text}")

        if "pause" in text or "play" in text:
            print("Ok, pausing")
            playpause()
        elif "next" in text:
            next_track()
        elif "previous" in text:
            prev_track()
        elif "open" in text and "browser" in text:
            print("Ok, opening your browser")
            webbrowser.open_new_tab("https://blank.org")
        elif "time" in text:
            print(f"Ok, the time is {datetime.datetime.now()}")
        elif "stop" in text:
            print("Ok, exiting")
            stop = True
        elif "volume down" in text:
            decreaseVol()
        elif "volume up" in text:
            increaseVol()
        elif "test" in text:
            print("Test received. It's working.")
        elif "battery" in text:
            Batteries.getBattery()
        elif "start menu" in text:
            pyautogui.press('win')
        elif "switch windows" in text or "switch window" in text:
            switchWindows()
        else:
            print("I don't understand...")
    except sr.exceptions.UnknownValueError:
        print("Sorry, I didn't catch that")