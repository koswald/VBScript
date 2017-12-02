
'launch Admin.spec.elev.vbs with elevated privileges

args = "/k cd """ & Resolve("..\..\spec\dll") & """ & echo " & _
    "cscript /nologo Admin.spec.elev.vbs & echo. & " & _
    "cscript /nologo Admin.spec.elev.vbs"

With CreateObject("Shell.Application")
   .ShellExecute "cmd", args,, "runas"
End With

Function Resolve(folder)
    With CreateObject("Scripting.FileSystemObject")
        Resolve = .GetAbsolutePathName(folder)
    End With
End Function
