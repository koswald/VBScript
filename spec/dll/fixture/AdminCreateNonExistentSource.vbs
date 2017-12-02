
'fixture for Admin.spec.vbs

With CreateObject("VBScripting.Admin")
    On Error Resume Next
        .CreateEventSource("VBScripting2")
    On Error Goto 0
End With
