
'fixture for Admin.spec.vbs

'attempt to delete an existing source
'intended only for running without elevated privileges,
'in which case it should fail

With CreateObject("VBScripting.Admin")
    On Error Resume Next
        .DeleteEventSource("WSH")
    On Error Goto 0
End With
