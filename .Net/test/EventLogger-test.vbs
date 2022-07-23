
'manual test that EventLogger.cs compiled correctly
'and that EventLogger.dll registered correctly

'expected outcome:
'A log entry is made with the current time and date, and
'is accessible with EventVwr.exe:
'expand Windows Logs, click on Application

Set logger = CreateObject( "VBScripting.EventLogger" )
Set net = CreateObject("WScript.Network")
Set sh = CreateObject( "WScript.Shell" )
With CreateObject( "VBScripting.Includer" )
    Execute .Read( "WoWChecker" )
End With
With New WowChecker
    If .isWoW Then
        bitness = " (32-bit)"
    Else bitness = " (64-bit)"
    End If
End With

logger.log _
    "who  " & vbTab & "username: " & net.UserName & vbLf & _
    "what " & vbTab & "Testing the .NET to COM logger" & bitness & vbLf & _
    "where" & vbTab & WScript.ScriptFullName & vbLf & _
    "when " & vbTab & FormatDateTime(Now, vbLongDate) & ", " & _
                      FormatDateTime(Now, vbLongTime)

msg = "Done logging. Open the event viewer?"
title = WScript.ScriptName
mode = vbSystemModal + vbInformation + vbOKCancel
response = MsgBox(msg, mode, title)
If vbOK = response Then
    msg = "After the Event Viewer opens, expand Windows Logs " & _
            "and select Application."
    mode = mode - vbOKCancel + vbOKOnly
    sh.Run "EventVwr.msc"
    MsgBox msg, mode, title
End If
Set logger = Nothing
Set net = Nothing
Set sh = Nothing
