
With CreateObject("WSHEventLogger")
    .log "" _
        & "info:   Testing the .NET to COM logger" & vbLf _
        & "script: " & WScript.ScriptFullName & vbLf _
        & "date:   " & FormatDateTime(Now, vbLongDate) & vbLf _
        & "time:   " & FormatDateTime(Now, vbLongTime) & vbLf _
        & ""
End With
With CreateObject("WScript.Shell")

    msg = "Done logging. Open the event viewer?"

    title = WScript.ScriptName
    mode = vbSystemModal + vbInformation + vbOKCancel
    If vbOK = .PopUp(msg, 20, title, mode) Then
        .Run "EventVwr.msc"
    End If
End With

