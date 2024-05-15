import datetime
from typing import TextIO
import wmi
import speech_recognition as SR
import sys
import re
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import pyttsx3

def FindNumers(s):
    numbers = re.findall(r'\d+', s)
    if(len(numbers) > 0):
    # numbers is a list of strings, convert to integers
        return int(numbers[0])
    else:
        return None


def main():
    rec = SR.Recognizer()
    writeToFile=0;
    with open('log.txt', 'a')  as LogFile:

        while True:
            with SR.Microphone() as source:
                text = rec.listen(source)
                print("Processing...");

            try:
                Command = rec.recognize_google(text)
                print(rec.recognize_google(text))
                if(Command=="save text"):
                    writeToFile=1

                if(Command=="do not save"):
                    writeToFile=0

                if(writeToFile):
                    LogFile.write(Command + "\n");

                elif("brightness" in Command):
                    if("max" in Command):
                        monitors = wmi.WMI(namespace="wmi")
                        for monitor in monitors.WmiMonitorBrightnessMethods():
                            monitor.WmiSetBrightness(96,0)
                    elif("min" in Command):
                        monitors = wmi.WMI(namespace="wmi")
                        for monitor in monitors.WmiMonitorBrightnessMethods():
                            monitor.WmiSetBrightness(5, 0)
                    elif(FindNumers(Command)):
                        monitors = wmi.WMI(namespace="wmi")
                        for monitor in monitors.WmiMonitorBrightnessMethods():
                            monitor.WmiSetBrightness(FindNumers(Command), 0)

                elif("volume" in Command):
                    if("max" in Command):
                        sessions = AudioUtilities.GetAllSessions()
                        for s in sessions :
                            vol = s._ctl.QueryInterface(ISimpleAudioVolume)
                            vol.SetMasterVolume(96/100.0, None)
                    elif ("min" in Command):
                        sessions = AudioUtilities.GetAllSessions()
                        for s in sessions:
                            vol = s._ctl.QueryInterface(ISimpleAudioVolume)
                            vol.SetMasterVolume(3/100.0, None)
                    elif(FindNumers(Command)):
                        sessions = AudioUtilities.GetAllSessions()
                        for s in sessions:
                            vol = s._ctl.QueryInterface(ISimpleAudioVolume)
                            vol.SetMasterVolumeLevel(FindNumers(Command)/100.0, None)
                    elif(Command == "read the log"):
                        print("READING LOG")
                        engine = pyttsx3.init()
                        with open('log.txt', 'r') as LogFile:
                            text = LogFile.read()
                            engine.say(text)
                            engine.runAndWait()

            except:
                print("Translation could not be performed")

if __name__ == "__main__":
    main()
