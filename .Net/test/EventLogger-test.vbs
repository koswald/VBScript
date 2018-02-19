
'manual test that EventLogger.cs compiled correctly
'and that EventLogger.dll registered correctly

'expected outcome:
'A log entry is made with the current time and date, and
'is accessible with EventVwr.exe:
'expand Windows Logs, click on Application

Main
Sub Main
    Set incl = CreateObject("VBScripting.Includer")
    Execute(incl.read("WoWChecker"))
    Set incl = Nothing
    Set wow = New WowChecker
    If wow.isWoW Then
        bitness = " (32-bit)"
    Else bitness = " (64-bit)"
    End If

    Set logger = CreateObject("VBScripting.EventLogger")
    logger.log "info:   Testing the .NET to COM logger" & bitness & vbLf & _
        "script: " & WScript.ScriptFullName & vbLf & _
        "date:   " & FormatDateTime(Now, vbLongDate) & vbLf & _
        "time:   " & FormatDateTime(Now, vbLongTime)
    Set logger = Nothing
    Set sh = CreateObject("WScript.Shell")

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
    Set sh = Nothing
End Sub

