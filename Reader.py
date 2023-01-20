from win32com.client import Dispatch

speak = Dispatch("Sapi.SpVoice")
speak.Speak("Good Morning Sir , I am your Personal Assistant ")
speak.Speak("Tell me something to do")
