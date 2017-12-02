
'launch the spec for testing FileChooser.dll

Dim args : args = "/k echo " & _
    "cscript //nologo FileChooser.spec.vbs & echo. & " & _
    "cscript //nologo FileChooser.spec.vbs"

With CreateObject("Shell.Application")
   .ShellExecute "cmd", args, "..\..\spec\dll"
End With
