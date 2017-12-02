
'fixture file for SpeechSynthesis.spec.vbs
'expected outcome:
'   a message box appears informing that 
'   invalid voice is an invalid voice

With CreateObject("VBScripting.SpeechSynthesis")
    .voice = "invalid voice"
End With
