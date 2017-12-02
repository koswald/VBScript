
'launch Admin.spec.vbs

Dim args : args = "/k echo " & _
    "cscript //nologo Admin.spec.vbs & echo. & " & _
    "cscript //nologo Admin.spec.vbs"

With CreateObject("Shell.Application")
   .ShellExecute "cmd", args, "..\..\spec\dll"
End With
