import win32com.client
import time

# Text-to-Speech (TTS)
def speak(text, rate=0, volume=100, voice=None):  # Added voice parameter
    try:
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        speaker.Rate = rate
        speaker.Volume = volume

        if voice:  # Set a specific voice if provided
            for v in speaker.GetVoices():
                if voice.lower() in v.GetDescription().lower():  # Case-insensitive matching
                    speaker.Voice = v
                    break

        speaker.Speak(text)
    except Exception as e:
        print(f"Error in TTS: {e}")


# Speech Recognition (SR)
def recognize_speech(grammar_file=None):
    try:
        recognizer = win32com.client.Dispatch("SAPI.SpRecognizer")
        grammar = win32com.client.Dispatch("SAPI.SpGrammar")

        if grammar_file:
            try:
                rule = grammar.Rules.Add("rule1")  # Add a rule
                rule.LoadFromFile(grammar_file)  # Load from file
                rule.Activate()  # Activate the rule
                grammar.CmdLoadFromFile(grammar_file)  # Command load
                grammar.CmdSetRuleState("rule1", 1)  # Set rule state
                grammar.CmdCommit()  # Commit changes
                recognizer.LoadGrammar(grammar) # load grammar to the recognizer
            except Exception as e:
                print(f"Error loading grammar: {e}")

        print("Listening...")
        while True:
            try:
                events = recognizer.GetRecognizeStream().GetEvents(1000) #Get events with timeout
                for event in events:
                    if event.EventId == 2:  # SPEI_RECOGNITION
                        recognized_text = event.RecoResult.PhraseInfo.GetText()
                        print("You said:", recognized_text)
                        return recognized_text # Return the recognized text
                        #break  # Exit loop after one recognition (you can remove this)
            except Exception as e:
                print(f"Recognition error: {e}")
                break # Break on error
            time.sleep(0.1) # small delay so that the CPU is not overloaded
    except Exception as e:
        print(f"Error initializing recognizer: {e}")
    return None  # Return None if no recognition


if __name__ == "__main__":
    # TTS Example
    speak("Hello, this is a test.  Now with a different voice.", rate=1, voice="Microsoft Zira Desktop")  # Example with voice selection
    speak("This is another test", voice="Microsoft David Desktop")
    