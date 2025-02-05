import win32com.client

def recognize_speech():
    recognizer = win32com.client.Dispatch("SAPI.SpRecognizer")
    grammar = win32com.client.Dispatch("SAPI.SpGrammar")
    rule = grammar.Rules.Add("rule1")
    rule.LoadFromFile("mygrammar.xml") # Load grammar from file
    rule.Activate()
    grammar.CmdLoadFromFile("mygrammar.xml")
    grammar.CmdSetRuleState("rule1", 1)
    grammar.CmdCommit()

    recognizer.LoadGrammar(grammar)

    print("Listening...")
    while True:
        try:
            event = recognizer.GetRecognizeStream().GetEvent()
            if event.EventId == 2: # SPEI_RECOGNITION
                print("You said:", event.RecoResult.PhraseInfo.GetText())
                break
        except Exception as e:
            print("Error", e)
            break

if __name__ == "__main__":
    recognize_speech()