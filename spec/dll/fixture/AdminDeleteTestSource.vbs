
'fixture for Admin.spec.elev.vbs

'delete an existing source

With CreateObject("VBScripting.Admin")
    On Error Resume Next
        .DeleteEventSource("VBScripting2")
    On Error Goto 0
End With
