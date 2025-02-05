import win32com.client

def speak(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

if __name__ == "__main__":
    text_to_say = "Hello, this is a test of SAPI in Python."
    speak(text_to_say)

