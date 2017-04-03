
'manual test that the .cs file compiled correctly
'and that the .dll registered correctly

Main
Sub Main
    Set incl = CreateObject("includer")
    Execute(incl.read("WoWChecker"))
    Set incl = Nothing
    Set wow = New WowChecker
    If wow.isWoW Then
        bitness = " (32-bit)"
    Else bitness = " (64-bit)"
    End If
    Set logger = CreateObject("WSHEventLogger")
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
        sh.Run "EventVwr.msc"
    End If
    Set sh = Nothing
End Sub

