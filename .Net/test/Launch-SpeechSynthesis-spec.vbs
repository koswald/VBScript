
'launch the spec for testing SpeechSynthesis.dll

Dim args : args = "/k echo " & _
    "cscript //nologo SpeechSynthesis.spec.vbs & echo. & " & _
    "cscript //nologo SpeechSynthesis.spec.vbs"

With CreateObject( "Shell.Application" )
   .ShellExecute "cmd", args, "..\..\spec\dll"
End With