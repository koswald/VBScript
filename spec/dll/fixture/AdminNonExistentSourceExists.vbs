
'fixture for Admin.spec.vbs

'attempt to find non-existent source

With CreateObject("VBScripting.Admin")
    On Error Resume Next
        .SourceExists("VBScripting2")
    On Error Goto 0
End With
