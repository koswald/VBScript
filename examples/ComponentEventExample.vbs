' .vbs usage demo for ComponentEventExample.wsc
Set cee = WScript.CreateObject("VBScripting.EventExample", "cee_")
cee.FireTestEvent
Sub cee_TestEvent
    MsgBox "Event fired successfully.", vbInformation, WScript.ScriptName
End Sub