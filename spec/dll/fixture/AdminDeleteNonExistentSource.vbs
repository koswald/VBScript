
'fixture for Admin.spec.vbs

'attempt to delete a non-existent source

With CreateObject("VBScripting.Admin")
    On Error Resume Next
        .DeleteEventSource("VBScripting2")
    On Error Goto 0
End With
