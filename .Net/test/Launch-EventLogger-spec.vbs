
'launch the spec for testing EventLogger.dll

Dim args : args = "/k echo " & _
    "cscript //nologo EventLogger.spec.vbs & echo. & " & _
    "cscript //nologo EventLogger.spec.vbs"

With CreateObject("Shell.Application")
   .ShellExecute "cmd", args, "..\..\spec\dll"
End With
