' .vbs usage demo for ComponentEventExample.wsc
Set cee = WScript.CreateObject("VBScripting.EventExample", "cee_")
cee.FireTestEvent
Sub cee_TestEvent
    WScript.Echo "Event fired successfully."
End Sub