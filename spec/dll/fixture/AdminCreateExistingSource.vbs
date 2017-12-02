
'fixture for Admin.spec.vbs

'attempt to create an existing source

With CreateObject("includer")
    Execute .read("../spec/dll/Admin.spec.config")
End With
With CreateObject("VBScripting.Admin")
    .CreateEventSource(source)
End With
